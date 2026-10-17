<%
' =====================================================================
'  Acceso a datos: SIEMPRE con ADODB.Command parametrizado (marcadores ?)
'  para evitar inyección SQL.
' =====================================================================

Const adInteger = 3, adVarChar = 200, adParamInput = 1, adCmdText = 1, adExecuteNoRecords = 128

Dim gConn, gConnOpen
gConnOpen = False

Function Db()
    If Not gConnOpen Then
        Set gConn = Server.CreateObject("ADODB.Connection")
        gConn.ConnectionTimeout = 15
        gConn.CommandTimeout = 30
        On Error Resume Next
        gConn.Open DB_CONN
        If Err.Number <> 0 Then
            Dim dbErr : dbErr = Err.Description
            On Error GoTo 0
            RenderFatal "No se pudo conectar con la base de datos", "Verifique que SQL Server esté en ejecución y que el inicio de sesión usr_polimusic tenga acceso a BDD_PoliMusic.", dbErr
        End If
        On Error GoTo 0
        gConnOpen = True
    End If
    Set Db = gConn
End Function

Sub DbClose()
    If gConnOpen Then
        If gConn.State = 1 Then gConn.Close
        Set gConn = Nothing
        gConnOpen = False
    End If
End Sub

' Crea un comando con parámetros posicionales. params = Array(v1, v2, ...)
Function NewCmd(sql, params)
    Dim cmd, i
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = Db()
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    If IsArray(params) Then
        For i = 0 To UBound(params)
            cmd.Parameters.Append MakeParam(cmd, params(i))
        Next
    End If
    Set NewCmd = cmd
End Function

Function MakeParam(cmd, v)
    Dim s, size
    Select Case VarType(v)
        Case vbInteger, vbLong, vbByte
            Set MakeParam = cmd.CreateParameter("", adInteger, adParamInput, , CLng(v))
        Case vbNull, vbEmpty
            Set MakeParam = cmd.CreateParameter("", adVarChar, adParamInput, 1, Null)
        Case Else
            s = CStr(v)
            size = Len(s)
            If size < 1 Then size = 1
            Set MakeParam = cmd.CreateParameter("", adVarChar, adParamInput, size, s)
    End Select
End Function

' Devuelve un arreglo de Scripting.Dictionary (una entrada por fila).
Function QList(sql, params)
    Dim rs, arr(), n, d, f
    ReDim arr(-1)
    n = -1
    Set rs = NewCmd(sql, params).Execute()
    Do While Not rs.EOF
        Set d = Server.CreateObject("Scripting.Dictionary")
        d.CompareMode = vbTextCompare
        For Each f In rs.Fields
            d(f.Name) = f.Value
        Next
        n = n + 1
        ReDim Preserve arr(n)
        Set arr(n) = d
        rs.MoveNext
    Loop
    rs.Close
    QList = arr
End Function

' Primera fila como Dictionary, o Nothing.
Function QRow(sql, params)
    Dim rows
    rows = QList(sql, params)
    If UBound(rows) >= 0 Then
        Set QRow = rows(0)
    Else
        Set QRow = Nothing
    End If
End Function

' Primer valor de la primera fila, o Null.
Function QScalar(sql, params)
    Dim rs
    Set rs = NewCmd(sql, params).Execute()
    If rs.State = 1 Then
        If Not rs.EOF Then
            QScalar = rs.Fields(0).Value
        Else
            QScalar = Null
        End If
        rs.Close
    Else
        QScalar = Null
    End If
End Function

' Ejecuta INSERT/UPDATE/DELETE. Devuelve filas afectadas.
Function QExec(sql, params)
    Dim ra
    ra = 0
    NewCmd(sql, params).Execute ra, , adExecuteNoRecords
    QExec = ra
End Function

' Ejecuta un INSERT y devuelve el IDENTITY generado.
Function QInsert(sql, params)
    QInsert = CLng(QScalar("SET NOCOUNT ON; " & sql & "; SELECT CAST(SCOPE_IDENTITY() AS int) AS NEW_ID;", params))
End Function

' Ejecuta un lote atómico (BEGIN TRAN ... COMMIT con XACT_ABORT).
' El lote puede asignar la variable @out (p. ej. SET @out = SCOPE_IDENTITY()).
' Devuelve "" si todo fue bien o un mensaje de error legible; outVal recibe @out.
Function TryBatch(sql, params, ByRef outVal)
    Dim rs, errText
    outVal = Null
    errText = ""
    On Error Resume Next
    Set rs = NewCmd("SET NOCOUNT ON; SET XACT_ABORT ON; DECLARE @out int; SET @out = NULL; BEGIN TRAN; " & sql & " COMMIT TRAN; SELECT @out AS OUT_VAL;", params).Execute()
    If Err.Number = 0 Then
        If rs.State = 1 Then
            If Not rs.EOF Then outVal = rs.Fields(0).Value
            rs.Close
        End If
    End If
    If Err.Number <> 0 Then
        errText = Err.Description
        Err.Clear
        Db().Execute "IF @@TRANCOUNT > 0 ROLLBACK TRAN", , adExecuteNoRecords
        Err.Clear
    End If
    On Error GoTo 0
    If errText <> "" Then
        TryBatch = FriendlyDbError(errText)
    Else
        TryBatch = ""
    End If
End Function

Function FriendlyDbError(msg)
    If InStr(1, msg, "REFERENCE", vbTextCompare) > 0 Or InStr(1, msg, "FOREIGN KEY", vbTextCompare) > 0 Then
        FriendlyDbError = "No se pudo completar la operación porque existen registros relacionados."
    ElseIf InStr(1, msg, "truncated", vbTextCompare) > 0 Or InStr(1, msg, "truncar", vbTextCompare) > 0 Then
        FriendlyDbError = "Uno de los valores excede la longitud permitida por la base de datos."
    Else
        FriendlyDbError = "Error de base de datos: " & msg
    End If
End Function

' "(?),(?),(?)" para insertar listas de IDs con parámetros
Function ValuesList(n)
    Dim i, s
    s = ""
    For i = 1 To n
        If i > 1 Then s = s & ","
        s = s & "(?)"
    Next
    ValuesList = s
End Function

' INSERT de relación N:M usando @out como dueño: (ownerCol=@out, otherCol=cada id)
Function RelSql(table, ownerCol, otherCol, ids)
    If UBound(ids) < 0 Then
        RelSql = ""
    Else
        RelSql = "INSERT INTO " & table & " (" & ownerCol & ", " & otherCol & ") SELECT @out, v.x FROM (VALUES " & ValuesList(UBound(ids) + 1) & ") v(x); "
    End If
End Function

' Concatena dos arreglos de parámetros
Function ConcatArr(a, b)
    Dim r(), i, n
    n = UBound(a) + UBound(b) + 1
    ReDim r(n)
    For i = 0 To UBound(a)
        r(i) = a(i)
    Next
    For i = 0 To UBound(b)
        r(UBound(a) + 1 + i) = b(i)
    Next
    ConcatArr = r
End Function

' Elimina una canción con todas sus relaciones (y su archivo subido). Devuelve "" o el error.
Function DeleteSong(id)
    Dim oldMedia, dummy
    oldMedia = QScalar("SELECT SONG_PATH FROM TBL_SONG WHERE ID_SONG = ?", Array(id))
    DeleteSong = TryBatch("SET @out = ?; DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = @out; DELETE FROM TBL_TRACKLIST WHERE ID_SONG = @out; " & _
                          "DELETE FROM TBL_PLAYER WHERE ID_SONG = @out; DELETE FROM TBL_SONG WHERE ID_SONG = @out;", Array(id), dummy)
    If DeleteSong = "" Then DeleteUploadedMedia oldMedia
End Function

' Bloque SQL reutilizable: canciones con artistas y portada del álbum.
Function SongSelectSql()
    SongSelectSql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS,0) AS PLAYS, " & _
        "ISNULL((SELECT STRING_AGG(LTRIM(RTRIM(a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME,''))), ', ') WITHIN GROUP (ORDER BY a.ARTIST_NAME) " & _
        "  FROM TBL_PLAYER p JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG), '') AS ARTISTS, " & _
        "(SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST t JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS COVER, " & _
        "(SELECT TOP 1 al.TITLE_ALBUM FROM TBL_TRACKLIST t JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS ALBUM_TITLE, " & _
        "(SELECT TOP 1 t.ID_ALBUM FROM TBL_TRACKLIST t WHERE t.ID_SONG = s.ID_SONG ORDER BY t.ID_ALBUM) AS ID_ALBUM " & _
        "FROM TBL_SONG s "
End Function

' Etiquetas del catálogo (p. ej. tipos de usuario o de playlist) como Dictionary cod -> valor
Function CatalogMap(parentCode)
    Dim d, rows, r
    Set d = Server.CreateObject("Scripting.Dictionary")
    rows = QList("SELECT COD_CATALOG, VALUE FROM TBL_CATALOG WHERE COD_CATALOG_PARENT = ? ORDER BY COD_CATALOG", Array(CStr(parentCode)))
    For Each r In rows
        d(CStr(r("COD_CATALOG"))) = r("VALUE")
    Next
    Set CatalogMap = d
End Function
%>
