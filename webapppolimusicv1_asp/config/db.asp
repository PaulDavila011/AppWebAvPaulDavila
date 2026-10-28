<%
'==============================================================================
' PoliMusic - Capa de acceso a datos (ADO / ADODB)
'==============================================================================
' Todas las consultas se ejecutan con parametros enlazados (ADODB.Command).
' Nunca se concatena entrada del usuario dentro del SQL: eso elimina la
' inyeccion SQL por diseno, no por filtrado de caracteres.
'==============================================================================

' --- Constantes ADO ----------------------------------------------------------
Const adCmdText          = 1
Const adParamInput       = 1
Const adInteger          = 3
Const adBoolean          = 11
Const adNumeric          = 131
Const adDBTimeStamp      = 135
Const adVarWChar         = 202
Const adOpenStatic       = 3
Const adLockReadOnly     = 1
Const adUseClient        = 3

'==============================================================================
' Abre una conexion a la base de datos (con respaldo automatico de proveedor).
'==============================================================================
Function OpenConnection()
    Dim cn, primaryErr

    Set cn = Server.CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 15
    cn.CommandTimeout = 30

    On Error Resume Next
    cn.Open ConnectionString()
    If Err.Number <> 0 Then
        primaryErr = Err.Description
        Err.Clear
        cn.Close
        ' Reintenta con el proveedor heredado
        cn.Open ConnectionStringFallback()
        If Err.Number <> 0 Then
            Dim msg
            msg = "No se pudo conectar a la base de datos." & vbCrLf & _
                  "Driver moderno: " & primaryErr & vbCrLf & _
                  "Driver heredado: " & Err.Description
            On Error GoTo 0
            Err.Raise vbObjectError + 1001, "OpenConnection", msg
        End If
    End If
    On Error GoTo 0

    Set OpenConnection = cn
End Function

'==============================================================================
' Cierra la conexion si sigue abierta.
'==============================================================================
Sub CloseConnection(cn)
    On Error Resume Next
    If IsObject(cn) Then
        If cn.State <> 0 Then cn.Close
        Set cn = Nothing
    End If
    On Error GoTo 0
End Sub

'==============================================================================
' Ejecuta un SELECT y devuelve un Recordset desconectado (client-side).
' Desconectado = seguro de usar despues de cerrar la conexion.
'
' OJO: la asignacion usa Set. Sin Set, VBScript no devuelve el objeto sino su
' propiedad predeterminada (para un Recordset eso es Fields) y las paginas
' fallan con "El objeto no acepta esta propiedad o metodo: 'EOF'".
'==============================================================================
Function ExecuteQuery(sql, params)
    Set ExecuteQuery = ExecuteQueryTyped(sql, params, Empty)
End Function

'==============================================================================
' Igual que ExecuteQuery pero permite forzar el tipo de cada parametro.
' paramTypes puede ser Empty (autodeteccion) o un array paralelo a params.
'==============================================================================
Function ExecuteQueryTyped(sql, params, paramTypes)
    Dim cn, cmd, i, rs
    Set cn = OpenConnection()

    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql

    If IsArray(params) Then
        For i = 0 To UBound(params)
            Dim pType
            If IsArray(paramTypes) Then
                pType = paramTypes(i)
            Else
                pType = ResolveParamType(params(i))
            End If
            cmd.Parameters.Append CreateParam(cmd, "p" & i, pType, params(i))
        Next
    End If

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockReadOnly

    ' Con CursorLocation = adUseClient los datos ya viven en memoria. Para que
    ' el Recordset sobreviva hay que soltar la conexion ANTES de cerrarla:
    ' se pone ActiveConnection a Nothing dentro de un manejador de errores,
    ' porque algunos proveedores (SQLOLEDB) no admiten esa asignacion.
    On Error Resume Next
    Set rs.ActiveConnection = Nothing
    Err.Clear
    On Error GoTo 0

    CloseConnection cn
    Set cmd = Nothing
    Set ExecuteQueryTyped = rs
End Function

'==============================================================================
' Ejecuta INSERT/UPDATE/DELETE y devuelve el numero de filas afectadas.
'==============================================================================
Function ExecuteNonQuery(sql, params)
    Dim cn, cmd, i, affected
    Set cn = OpenConnection()

    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql

    If IsArray(params) Then
        For i = 0 To UBound(params)
            cmd.Parameters.Append CreateParam(cmd, "p" & i, ResolveParamType(params(i)), params(i))
        Next
    End If

    cmd.Execute affected
    affected = CLng(affected)
    If affected < 0 Then affected = 0

    CloseConnection cn
    ExecuteNonQuery = affected
End Function

'==============================================================================
' Ejecuta un INSERT y devuelve el ID generado (SCOPE_IDENTITY).
'==============================================================================
Function ExecuteInsert(sql, params)
    Dim cn, cmd, i, newId
    Set cn = OpenConnection()

    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    ' Se agrega el SELECT en la misma sentencia para leer la identidad recien creada
    cmd.CommandText = sql & vbCrLf & "SELECT CAST(SCOPE_IDENTITY() AS int) AS NewID"

    If IsArray(params) Then
        For i = 0 To UBound(params)
            cmd.Parameters.Append CreateParam(cmd, "p" & i, ResolveParamType(params(i)), params(i))
        Next
    End If

    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockReadOnly

    newId = 0
    If Not rs.EOF Then
        If Not IsNull(rs("NewID")) Then newId = CLng(rs("NewID"))
    End If
    rs.Close
    Set rs = Nothing

    CloseConnection cn
    ExecuteInsert = newId
End Function

'==============================================================================
' Devuelve un unico valor escalar (primera columna de la primera fila).
'==============================================================================
Function ExecuteScalar(sql, params)
    Dim rs, value
    Set rs = ExecuteQuery(sql, params)
    value = Null
    If Not rs.EOF Then
        If Not IsNull(rs(0)) Then value = rs(0)
    End If
    rs.Close
    Set rs = Nothing
    ExecuteScalar = value
End Function

'==============================================================================
' Fabrica un parametro ADO con el tipo y tamano correctos.
'==============================================================================
Function CreateParam(cmd, name, pType, value)
    Dim p, size
    size = ResolveParamSize(value)
    Set p = cmd.CreateParameter(name, pType, adParamInput, size)

    If IsNull(value) Or IsEmpty(value) Then
        p.Value = Null
    ElseIf pType = adInteger Or pType = adNumeric Or pType = adBoolean Then
        p.Value = value
    ElseIf pType = adDBTimeStamp Then
        p.Value = CDate(value)
    Else
        p.Value = CStr(value)
    End If

    Set CreateParam = p
End Function

'==============================================================================
' Deduce el tipo ADO a partir del valor de VBScript.
'==============================================================================
Function ResolveParamType(value)
    Dim t
    ResolveParamType = adVarWChar

    If IsNull(value) Or IsEmpty(value) Then
        ResolveParamType = adVarWChar
        Exit Function
    End If

    t = VarType(value)
    If t = 11 Then                       ' vbBoolean
        ResolveParamType = adBoolean
    ElseIf t = 2 Or t = 3 Then           ' vbInteger / vbLong
        ResolveParamType = adInteger
    ElseIf t = 4 Or t = 5 Or t = 6 Then  ' vbSingle / vbDouble / vbCurrency
        ResolveParamType = adNumeric
    ElseIf t = 7 Then                    ' vbDate
        ResolveParamType = adDBTimeStamp
    ElseIf t = 8 Then                    ' vbString
        ResolveParamType = adVarWChar
    Else
        ResolveParamType = adVarWChar
    End If
End Function

'==============================================================================
' Deduce el tamano del parametro (1..4000).
'==============================================================================
Function ResolveParamSize(value)
    Dim n
    n = 1
    If Not IsNull(value) And Not IsEmpty(value) Then
        If VarType(value) = 8 Then n = Len(CStr(value))
    End If
    If n < 1 Then n = 1
    If n > 4000 Then n = 4000
    ResolveParamSize = n
End Function
%>
