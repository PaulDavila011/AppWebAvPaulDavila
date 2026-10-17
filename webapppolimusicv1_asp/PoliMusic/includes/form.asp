<%
' =====================================================================
'  Acceso unificado a formularios (urlencoded y multipart/form-data)
'  ASP clásico no procesa multipart: se implementa un lector binario
'  con ADODB.Stream. Nunca usar Request("x") ni mezclar Request.Form
'  con Request.BinaryRead en la misma petición.
' =====================================================================

Dim gFormReady, gFormMultipart, gFormFields, gFormFiles, gFormRaw, gFormError
gFormReady = False
gFormMultipart = False
gFormError = ""

Sub FormInit()
    Dim ct
    If gFormReady Then Exit Sub
    gFormReady = True
    ct = Request.ServerVariables("CONTENT_TYPE")
    If InStr(1, ct, "multipart/form-data", vbTextCompare) > 0 Then
        gFormMultipart = True
        Set gFormFields = Server.CreateObject("Scripting.Dictionary")
        gFormFields.CompareMode = vbTextCompare
        Set gFormFiles = Server.CreateObject("Scripting.Dictionary")
        gFormFiles.CompareMode = vbTextCompare
        ParseMultipart ct
    End If
End Sub

Function FormValue(name)
    FormInit
    If gFormMultipart Then
        If gFormFields.Exists(name) Then
            FormValue = gFormFields(name)(0)
        Else
            FormValue = ""
        End If
    Else
        If Request.Form(name).Count > 0 Then
            FormValue = Request.Form(name)(1)
        Else
            FormValue = ""
        End If
    End If
End Function

Function FormText(name)
    FormText = Trim(FormValue(name))
End Function

' Todos los valores de un campo repetido (checkboxes) como arreglo
Function FormValues(name)
    Dim arr(), i, n
    FormInit
    If gFormMultipart Then
        If gFormFields.Exists(name) Then
            FormValues = gFormFields(name)
        Else
            ReDim arr(-1)
            FormValues = arr
        End If
    Else
        n = Request.Form(name).Count
        ReDim arr(n - 1)
        For i = 1 To n
            arr(i - 1) = Request.Form(name)(i)
        Next
        FormValues = arr
    End If
End Function

' Valores enteros únicos de un campo repetido
Function FormIds(name)
    Dim vals, v, d, id
    Set d = Server.CreateObject("Scripting.Dictionary")
    vals = FormValues(name)
    For Each v In vals
        id = ToInt(v, 0)
        If id > 0 Then d(CStr(id)) = id
    Next
    FormIds = d.Items
End Function

Function FormHasFile(name)
    FormInit
    FormHasFile = False
    If gFormMultipart Then
        If gFormFiles.Exists(name) Then
            FormHasFile = (gFormFiles(name)("length") > 0 And gFormFiles(name)("filename") <> "")
        End If
    End If
End Function

Function FileExt(filename)
    Dim p
    p = InStrRev(filename, ".")
    If p > 0 Then
        FileExt = LCase(Mid(filename, p))
    Else
        FileExt = ""
    End If
End Function

' Guarda el archivo subido en /folder con un nombre aleatorio.
' Devuelve la ruta "~/folder/archivo.ext" o "" (y errMsg) si falla.
Function FormSaveFile(name, folder, allowedExts, maxBytes, ByRef errMsg)
    Dim f, ext, fname, outStream, physical
    FormSaveFile = ""
    errMsg = ""
    If Not FormHasFile(name) Then Exit Function
    Set f = gFormFiles(name)
    ext = FileExt(f("filename"))
    If ext = "" Or InStr("," & allowedExts & ",", "," & ext & ",") = 0 Then
        errMsg = "Tipo de archivo no permitido (" & ext & "). Permitidos: " & Replace(allowedExts, ",", ", ")
        Exit Function
    End If
    If f("length") > maxBytes Then
        errMsg = "El archivo supera el tamaño máximo de " & FmtNum(maxBytes \ 1048576) & " MB."
        Exit Function
    End If
    Randomize
    fname = LCase(folder) & "_" & Year(Now) & Pad2(Month(Now)) & Pad2(Day(Now)) & Pad2(Hour(Now)) & Pad2(Minute(Now)) & Pad2(Second(Now)) & "_" & Hex(Int(Rnd * 16777215)) & ext
    physical = Server.MapPath(folder & "/" & fname)

    Set outStream = Server.CreateObject("ADODB.Stream")
    outStream.Type = 1
    outStream.Open
    gFormRaw.Position = f("start")
    gFormRaw.CopyTo outStream, f("length")
    On Error Resume Next
    outStream.SaveToFile physical, 2
    If Err.Number <> 0 Then
        errMsg = "No se pudo guardar el archivo en el servidor (permisos de escritura en /" & folder & ")."
        Err.Clear
    Else
        FormSaveFile = "~/" & folder & "/" & fname
    End If
    On Error GoTo 0
    outStream.Close
End Function

' Elimina del disco un archivo subido por la aplicación (nombre generado por FormSaveFile).
' Nunca toca archivos con otro patrón (p. ej. los del script original de la BD).
Sub DeleteUploadedMedia(path)
    Dim p, fso, physical
    p = Nz(path, "")
    If Not ReTest("^~/(songFiles|albumImages|ArtistImages|userImages|playlistImages)/[a-z]+_\d{14}_[0-9A-F]+\.[a-z0-9]+$", p) Then Exit Sub
    ' Solo si ningún otro registro sigue usando el archivo
    If CLng(QScalar("SELECT (SELECT COUNT(*) FROM TBL_SONG WHERE SONG_PATH = ?) + (SELECT COUNT(*) FROM TBL_ALBUM WHERE ALBUM_IMAGE_PATH = ?) + " & _
        "(SELECT COUNT(*) FROM TBL_ARTIST WHERE ARTIST_IMAGE = ?) + (SELECT COUNT(*) FROM TBL_USER WHERE USER_PHOTO = ?) + (SELECT COUNT(*) FROM TBL_PLAYLIST WHERE PHOTO = ?)", _
        Array(p, p, p, p, p))) > 0 Then Exit Sub
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    physical = Server.MapPath(Mid(p, 3))
    On Error Resume Next
    If fso.FileExists(physical) Then fso.DeleteFile physical, True
    Err.Clear
    On Error GoTo 0
End Sub

' ---------------- Implementación del lector multipart ----------------

Sub ParseMultipart(ct)
    Dim total, remaining, size, chunk, boundary, p, text
    Dim delim, pos, partStart, hdrEnd, dataStart, nextPos, headers, fieldName, info, arr

    total = Request.TotalBytes
    If total > MAX_UPLOAD_BYTES + 65536 Then
        gFormError = "La petición supera el tamaño máximo permitido."
        Exit Sub
    End If

    p = InStr(1, ct, "boundary=", vbTextCompare)
    If p = 0 Then Exit Sub
    boundary = Mid(ct, p + 9)
    If InStr(boundary, ";") > 0 Then boundary = Left(boundary, InStr(boundary, ";") - 1)
    boundary = Replace(Trim(boundary), """", "")

    Set gFormRaw = Server.CreateObject("ADODB.Stream")
    gFormRaw.Type = 1
    gFormRaw.Open
    remaining = total
    Do While remaining > 0
        size = 524288
        If size > remaining Then size = remaining
        chunk = Request.BinaryRead(size)
        If size <= 0 Then Exit Do
        gFormRaw.Write chunk
        remaining = remaining - size
    Loop

    ' Vista de texto 1 byte = 1 carácter para localizar delimitadores
    gFormRaw.Position = 0
    gFormRaw.Type = 2
    gFormRaw.Charset = "iso-8859-1"
    text = gFormRaw.ReadText
    gFormRaw.Position = 0
    gFormRaw.Type = 1

    delim = "--" & boundary
    pos = InStr(1, text, delim, vbBinaryCompare)
    Do While pos > 0
        partStart = pos + Len(delim)
        If Mid(text, partStart, 2) = "--" Then Exit Do
        partStart = partStart + 2
        hdrEnd = InStr(partStart, text, vbCrLf & vbCrLf, vbBinaryCompare)
        If hdrEnd = 0 Then Exit Do
        headers = Mid(text, partStart, hdrEnd - partStart)
        dataStart = hdrEnd + 4
        nextPos = InStr(dataStart, text, vbCrLf & delim, vbBinaryCompare)
        If nextPos = 0 Then Exit Do

        fieldName = HeaderAttr(headers, "name")
        If fieldName <> "" Then
            If ReTest("\bfilename=""", headers) Then
                Set info = Server.CreateObject("Scripting.Dictionary")
                info("filename") = HeaderAttr(headers, "filename")
                info("start") = dataStart - 1
                info("length") = nextPos - dataStart
                If Not gFormFiles.Exists(fieldName) Then Set gFormFiles(fieldName) = info
            Else
                If gFormFields.Exists(fieldName) Then
                    arr = gFormFields(fieldName)
                    ReDim Preserve arr(UBound(arr) + 1)
                Else
                    ReDim arr(0)
                End If
                arr(UBound(arr)) = SliceUtf8(dataStart - 1, nextPos - dataStart)
                gFormFields(fieldName) = arr
            End If
        End If
        pos = nextPos + 2
    Loop
End Sub

Function HeaderAttr(headers, attr)
    Dim re, m
    Set re = New RegExp
    re.Pattern = "\b" & attr & "=""([^""]*)"""
    re.IgnoreCase = True
    Set m = re.Execute(headers)
    If m.Count > 0 Then
        HeaderAttr = m(0).SubMatches(0)
    Else
        HeaderAttr = ""
    End If
End Function

Function SliceUtf8(offset, length)
    Dim tmp
    If length <= 0 Then
        SliceUtf8 = ""
        Exit Function
    End If
    Set tmp = Server.CreateObject("ADODB.Stream")
    tmp.Type = 1
    tmp.Open
    gFormRaw.Position = offset
    gFormRaw.CopyTo tmp, length
    tmp.Position = 0
    tmp.Type = 2
    tmp.Charset = "utf-8"
    SliceUtf8 = tmp.ReadText
    tmp.Close
End Function
%>
