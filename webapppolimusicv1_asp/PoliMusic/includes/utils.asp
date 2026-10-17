<%
' =====================================================================
'  Utilidades generales: codificación, conversión, formato, UI helpers
' =====================================================================

' Codificación HTML segura (contenido y atributos)
Function H(v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        H = ""
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    H = s
End Function

Function JsonEsc(v)
    Dim s
    s = Nz(v, "")
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    s = Replace(s, "<", "<")
    JsonEsc = s
End Function

Function Nz(v, def)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = def
    Else
        Nz = CStr(v)
    End If
End Function

Function ReTest(pattern, s)
    Dim re
    Set re = New RegExp
    re.Pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    ReTest = re.Test(Nz(s, ""))
End Function

' Entero validado (sin errores de conversión)
Function ToInt(v, def)
    Dim s
    s = Trim(Nz(v, ""))
    If ReTest("^-?\d{1,9}$", s) Then
        ToInt = CLng(s)
    Else
        ToInt = def
    End If
End Function

Function QS(name)
    QS = Trim(Nz(Request.QueryString(name), ""))
End Function

Function IsPost()
    IsPost = (UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST")
End Function

' Convierte rutas tipo "~/carpeta/archivo" (como las guarda la BD) en URL relativa al sitio
Function MediaUrl(path)
    Dim p
    p = Trim(Nz(path, ""))
    If p = "" Then
        MediaUrl = ""
    ElseIf Left(p, 2) = "~/" Then
        MediaUrl = Replace(Mid(p, 3), " ", "%20")
    ElseIf LCase(Left(p, 7)) = "http://" Or LCase(Left(p, 8)) = "https://" Then
        MediaUrl = p
    Else
        MediaUrl = Replace(p, " ", "%20")
    End If
End Function

' ¿Existe físicamente el archivo de una ruta "~/..."?
Function MediaExists(path)
    Dim p, fso
    p = Trim(Nz(path, ""))
    MediaExists = False
    If Left(p, 2) = "~/" And InStr(p, "..") = 0 Then
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        MediaExists = fso.FileExists(Server.MapPath(Mid(p, 3)))
        On Error GoTo 0
    ElseIf LCase(Left(p, 4)) = "http" Then
        MediaExists = True
    End If
End Function

Function Pad2(n)
    Pad2 = Right("0" & CStr(n), 2)
End Function

Function IsoDate(d)
    If IsDate(d) Then
        IsoDate = Year(d) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d))
    Else
        IsoDate = ""
    End If
End Function

Function FmtDate(d)
    Dim months
    months = Array("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")
    If IsDate(d) Then
        FmtDate = Day(d) & " " & months(Month(d) - 1) & " " & Year(d)
    Else
        FmtDate = "—"
    End If
End Function

' 12345 -> "12.345"
Function FmtNum(n)
    Dim s, out, i, c
    s = CStr(ToInt(n, 0))
    out = ""
    c = 0
    For i = Len(s) To 1 Step -1
        out = Mid(s, i, 1) & out
        c = c + 1
        ' VBScript no cortocircuita And: se anidan las condiciones
        If c Mod 3 = 0 And i > 1 Then
            If Mid(s, i - 1, 1) <> "-" Then out = "." & out
        End If
    Next
    FmtNum = out
End Function

Function Plural(n, one, many)
    If CLng(n) = 1 Then
        Plural = n & " " & one
    Else
        Plural = FmtNum(n) & " " & many
    End If
End Function

Function Initials(name)
    Dim parts, s, p
    s = ""
    parts = Split(Trim(Replace(Replace(Nz(name, "?"), ".", " "), "_", " ")), " ")
    For Each p In parts
        If Len(p) > 0 And Len(s) < 2 Then s = s & UCase(Left(p, 1))
    Next
    If s = "" Then s = "?"
    Initials = s
End Function

Function Greeting()
    Dim hr
    hr = Hour(Now())
    If hr < 12 Then
        Greeting = "Buenos días"
    ElseIf hr < 19 Then
        Greeting = "Buenas tardes"
    Else
        Greeting = "Buenas noches"
    End If
End Function

Function ArtistFullName(r)
    ArtistFullName = Trim(Nz(r("ARTIST_NAME"), "") & " " & Nz(r("ARTIST_LASTNAME"), ""))
End Function

' Icono del sprite SVG (definido en layout)
Function Icon(name)
    Icon = "<svg class=""ico"" aria-hidden=""true""><use href=""#i-" & name & """/></svg>"
End Function

' Imagen con respaldo automático si el archivo no existe
Function Img(path, kind, alt, cls)
    Dim fb
    Select Case kind
        Case "artist" : fb = "assets/img/fallback-artist.svg"
        Case "user"   : fb = "assets/img/fallback-user.svg"
        Case Else     : fb = "assets/img/fallback-cover.svg"
    End Select
    Dim src
    src = MediaUrl(path)
    If src = "" Then src = fb
    Img = "<img src=""" & H(src) & """ data-fallback=""" & fb & """ alt=""" & H(alt) & """ class=""" & cls & """ loading=""lazy"">"
End Function

' ---------- Mensajes flash (patrón Post/Redirect/Get) ----------
Sub SetFlash(kind, msg)
    Session("flash_kind") = kind
    Session("flash_msg") = msg
End Sub

Sub RenderFlash()
    Dim k, m
    m = Session("flash_msg")
    If Nz(m, "") <> "" Then
        k = Session("flash_kind")
        Response.Write "<div class=""alert alert-" & H(k) & """ role=""status"">"
        If k = "success" Then
            Response.Write Icon("check")
        Else
            Response.Write Icon("alert")
        End If
        Response.Write "<span>" & H(m) & "</span></div>"
        Session("flash_msg") = ""
        Session("flash_kind") = ""
    End If
End Sub

Sub RedirectTo(url)
    DbClose
    Response.Redirect url
    Response.End
End Sub

' Solo URLs internas del tipo pagina.asp?x=y
Function SafeReturnUrl(u)
    If ReTest("^[a-z0-9_]+\.asp(\?[a-z0-9_=&%.\-]*)?$", u) Then
        SafeReturnUrl = u
    Else
        SafeReturnUrl = "default.asp"
    End If
End Function

' Página de error autónoma (sin layout) para fallos graves
Sub RenderFatal(title, message, detail)
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.Write "<!DOCTYPE html><html lang=""es""><head><meta charset=""utf-8""><meta name=""viewport"" content=""width=device-width, initial-scale=1"">"
    Response.Write "<title>" & H(title) & " · " & APP_NAME & "</title><link rel=""stylesheet"" href=""assets/css/app.css""></head>"
    Response.Write "<body class=""page-center""><div class=""error-card""><div class=""error-code"">!</div><h1>" & H(title) & "</h1><p>" & H(message) & "</p>"
    Dim addr : addr = Request.ServerVariables("REMOTE_ADDR")
    If detail <> "" And (addr = "127.0.0.1" Or addr = "::1") Then
        Response.Write "<pre class=""error-detail"">" & H(detail) & "</pre>"
    End If
    Response.Write "<a class=""btn btn-primary"" href=""default.asp"">Volver al inicio</a></div></body></html>"
    Response.End
End Sub
%>
