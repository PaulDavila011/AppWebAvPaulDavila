<%
'==============================================================================
' PoliMusic - Funciones utilitarias compartidas
'==============================================================================
' Nada aqui toca la base de datos: son helpers puros de presentacion y de
' validacion. Se incluyen en todas las paginas via includes/bootstrap.asp.
'==============================================================================

'==============================================================================
' H() - Escapa HTML. TODA salida proveniente de la base de datos o del usuario
' debe pasar por aqui antes de imprimirse. Evita XSS almacenado.
'==============================================================================
Function H(value)
    Dim s
    If IsNull(value) Or IsEmpty(value) Then
        H = ""
        Exit Function
    End If
    s = CStr(value)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    H = s
End Function

'==============================================================================
' Ha() - Escapa para usar dentro de un atributo HTML (igual que H, alias
' semantico para dejar claro el contexto).
'==============================================================================
Function Ha(value)
    Ha = H(value)
End Function

'==============================================================================
' JsArg() - Escapa una cadena para incrustarla como literal JavaScript entre
' comillas simples.
'==============================================================================
Function JsArg(value)
    Dim s
    If IsNull(value) Or IsEmpty(value) Then
        JsArg = ""
        Exit Function
    End If
    s = CStr(value)
    s = Replace(s, "\", "\\")
    s = Replace(s, "'", "\'")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, "<", "\u003C")
    s = Replace(s, ">", "\u003E")
    JsArg = s
End Function

'==============================================================================
' JsonEscape() - Escapa una cadena para embeberla en JSON.
'==============================================================================
Function JsonEscape(value)
    Dim s
    If IsNull(value) Or IsEmpty(value) Then
        JsonEscape = ""
        Exit Function
    End If
    s = CStr(value)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

'==============================================================================
' MediaUrl() - Traduce una ruta guardada en la BD ("~/songFiles/x.mp3") a una
' URL servible por IIS ("assets/media/songFiles/x.mp3").
' Si la ruta viene vacia o nula se devuelve el fallback indicado.
'==============================================================================
Function MediaUrl(dbPath, fallback)
    Dim s
    s = CStr(dbPath & "")

    If Len(Trim(s)) = 0 Then
        MediaUrl = fallback
        Exit Function
    End If

    ' Normaliza las barras invertidas que a veces quedan en la BD
    s = Replace(s, "\", "/")

    If Left(s, 2) = MEDIA_PREFIX Then
        s = Mid(s, 3)
    ElseIf Left(s, 1) = "/" Then
        s = Mid(s, 2)
    End If

    ' Caso especial: las imagenes de artista en la BD usan "ArtistImages"
    ' (mayuscula) pero la carpeta real es "artistImages".
    If LCase(Left(s, 13)) = "artistimages/" Then
        s = "artistImages/" & Mid(s, 14)
    End If

    MediaUrl = MEDIA_BASE_URL & s
End Function

'==============================================================================
' UserPhotoUrl() - URL de la foto de perfil de un usuario.
'==============================================================================
Function UserPhotoUrl(dbPath)
    Dim s
    s = CStr(dbPath & "")
    If Len(Trim(s)) = 0 Then
        UserPhotoUrl = MEDIA_BASE_URL & "userImages/default.jpg"
    Else
        UserPhotoUrl = MediaUrl(s, MEDIA_BASE_URL & "userImages/default.jpg")
    End If
End Function

'==============================================================================
' CoverUrl() - URL de portada con respaldo generico.
'==============================================================================
Function CoverUrl(dbPath)
    CoverUrl = MediaUrl(dbPath, MEDIA_BASE_URL & "albumImages/default.png")
End Function

'==============================================================================
' CleanText() - Normaliza entrada de texto: recorta espacios y limita a 4000.
'==============================================================================
Function CleanText(value)
    Dim s
    If IsNull(value) Or IsEmpty(value) Then
        CleanText = ""
        Exit Function
    End If
    s = CStr(value)
    s = Trim(s)
    If Len(s) > 4000 Then s = Left(s, 4000)
    CleanText = s
End Function

'==============================================================================
' Truncate() - Recorta un texto agregando puntos suspensivos.
'==============================================================================
Function Truncate(value, maxLen)
    Dim s
    s = CStr(value & "")
    If Len(s) <= maxLen Then
        Truncate = s
    Else
        Truncate = Left(s, maxLen - 1) & Chr(8230)   ' …
    End If
End Function

'==============================================================================
' FmtDate() - Formatea una fecha como dd/mm/aaaa.
'==============================================================================
Function FmtDate(value)
    If IsNull(value) Or IsEmpty(value) Then
        FmtDate = "-"
        Exit Function
    End If
    On Error Resume Next
    FmtDate = Right("0" & Day(value), 2) & "/" & Right("0" & Month(value), 2) & "/" & Year(value)
    If Err.Number <> 0 Then
        FmtDate = "-"
        Err.Clear
    End If
    On Error GoTo 0
End Function

'==============================================================================
' FmtDuration() - Convierte segundos a "m:ss".
'==============================================================================
Function FmtDuration(totalSeconds)
    Dim secs, m, s
    If IsNull(totalSeconds) Or IsEmpty(totalSeconds) Then
        FmtDuration = "0:00"
        Exit Function
    End If
    secs = CLng(totalSeconds)
    If secs < 0 Then secs = 0
    m = secs \ 60
    s = secs Mod 60
    FmtDuration = m & ":" & Right("0" & s, 2)
End Function

'==============================================================================
' IsBlank() - True si la cadena esta vacia o solo tiene espacios.
'==============================================================================
Function IsBlank(value)
    If IsNull(value) Or IsEmpty(value) Then
        IsBlank = True
    Else
        IsBlank = (Len(Trim(CStr(value))) = 0)
    End If
End Function

'==============================================================================
' IsValidEmail() - Validacion basica de formato de correo.
'==============================================================================
Function IsValidEmail(value)
    Dim s, atPos, dotPos, i, ch
    IsValidEmail = False
    If IsBlank(value) Then Exit Function

    s = Trim(CStr(value))
    If Len(s) < 6 Or Len(s) > 100 Then Exit Function
    If InStr(s, " ") > 0 Then Exit Function

    atPos = InStr(s, "@")
    If atPos < 2 Then Exit Function
    If InStr(atPos + 1, s, "@") > 0 Then Exit Function

    dotPos = InStrRev(s, ".")
    If dotPos <= atPos + 1 Then Exit Function
    If dotPos >= Len(s) Then Exit Function

    ' Solo caracteres razonables
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If Not ( _
            (ch >= "a" And ch <= "z") Or _
            (ch >= "A" And ch <= "Z") Or _
            (ch >= "0" And ch <= "9") Or _
            ch = "@" Or ch = "." Or ch = "_" Or ch = "-" Or ch = "+" Or ch = "%" _
        ) Then
            Exit Function
        End If
    Next

    IsValidEmail = True
End Function

'==============================================================================
' ParseId() - Convierte un valor de querystring/formulario en entero positivo.
' Devuelve 0 si no es un ID valido.
'==============================================================================
Function ParseId(value)
    Dim s, n, i, ch, ok
    ParseId = 0
    If IsNull(value) Or IsEmpty(value) Then Exit Function

    s = Trim(CStr(value))
    If Len(s) = 0 Or Len(s) > 9 Then Exit Function

    ok = True
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch < "0" Or ch > "9" Then
            ok = False
            Exit For
        End If
    Next
    If Not ok Then Exit Function

    n = CLng(s)
    If n > 0 Then ParseId = n
End Function

'==============================================================================
' Flash messages - Mensajes cortos que sobreviven una redireccion.
'==============================================================================
Sub SetFlash(kind, message)
    Session("pm_flash_kind") = kind
    Session("pm_flash_msg") = message
End Sub

Function TakeFlashKind()
    TakeFlashKind = CStr(Session("pm_flash_kind") & "")
    Session("pm_flash_kind") = ""
End Function

Function TakeFlashMsg()
    TakeFlashMsg = CStr(Session("pm_flash_msg") & "")
    Session("pm_flash_msg") = ""
End Function

'==============================================================================
' Anti-CSRF - Token por sesion incrustado en cada formulario POST.
'==============================================================================
Function CsrfToken()
    Dim t
    t = CStr(Session("pm_csrf") & "")
    If Len(t) = 0 Then
        t = NewToken()
        Session("pm_csrf") = t
    End If
    CsrfToken = t
End Function

'==============================================================================
' NewToken() - Genera un token aleatorio (sin dependencias externas).
'==============================================================================
Function NewToken()
    Dim g, s, i, rndVal
    Randomize
    s = ""
    For i = 1 To 32
        rndVal = Int(Rnd() * 16)
        s = s & Hex(rndVal)
    Next
    ' Mezcla con datos del reloj y del servidor para no depender solo de Rnd
    s = s & Hex(Timer() * 1000) & Hex(Len(Request.ServerVariables("HTTP_USER_AGENT") & ""))
    NewToken = s
End Function

'==============================================================================
' CsrfField() - Input oculto listo para pegar dentro de un <form>.
'==============================================================================
Function CsrfField()
    CsrfField = "<input type=""hidden"" name=""pm_csrf"" value=""" & Ha(CsrfToken()) & """>"
End Function

'==============================================================================
' ValidateCsrf() - Comprueba que el POST trae el token correcto.
'==============================================================================
Function ValidateCsrf()
    Dim sent, expected
    sent = CStr(Request.Form("pm_csrf") & "")
    expected = CStr(Session("pm_csrf") & "")
    If Len(expected) = 0 Then
        ValidateCsrf = False
    ElseIf StrComp(sent, expected, 0) = 0 Then
        ValidateCsrf = True
    Else
        ValidateCsrf = False
    End If
End Function

'==============================================================================
' DenyAccess() - Pagina 403 (autenticado pero sin permisos).
'==============================================================================
Sub DenyAccess(message)
    Response.Status = "403 Forbidden"
    Response.Write "<!DOCTYPE html><html lang=""es""><head><meta charset=""utf-8"">"
    Response.Write "<meta name=""viewport"" content=""width=""device-width, initial-scale=1"">"
    Response.Write "<title>Acceso denegado | PoliMusic</title>"
    Response.Write "<link rel=""stylesheet"" href=""" & AppRoot() & "assets/css/polimusic.css"">"
    Response.Write "</head><body class=""pm-error-page"">"
    Response.Write "<div class=""pm-error-box"">"
    Response.Write "<span class=""pm-error-code"">403</span>"
    Response.Write "<h1>Acceso denegado</h1>"
    Response.Write "<p>" & H(message) & "</p>"
    Response.Write "<p class=""pm-error-actions"">"
    Response.Write "<a class=""pm-btn pm-btn-primary"" href=""" & AppRoot() & "default.asp"">Volver al inicio</a>"
    Response.Write "</p></div></body></html>"
    Response.End
End Sub

'==============================================================================
' NotFound() - Pagina 404.
'==============================================================================
Sub NotFound(message)
    Response.Status = "404 Not Found"
    Response.Write "<!DOCTYPE html><html lang=""es""><head><meta charset=""utf-8"">"
    Response.Write "<meta name=""viewport"" content=""width=""device-width, initial-scale=1"">"
    Response.Write "<title>No encontrado | PoliMusic</title>"
    Response.Write "<link rel=""stylesheet"" href=""" & AppRoot() & "assets/css/polimusic.css"">"
    Response.Write "</head><body class=""pm-error-page"">"
    Response.Write "<div class=""pm-error-box"">"
    Response.Write "<span class=""pm-error-code"">404</span>"
    Response.Write "<h1>Pagina no encontrada</h1>"
    Response.Write "<p>" & H(message) & "</p>"
    Response.Write "<p class=""pm-error-actions"">"
    Response.Write "<a class=""pm-btn pm-btn-primary"" href=""" & AppRoot() & "default.asp"">Volver al inicio</a>"
    Response.Write "</p></div></body></html>"
    Response.End
End Sub

'==============================================================================
' AppRoot() - Raiz de la aplicacion ("/" si esta en la raiz del sitio).
'==============================================================================
Function AppRoot()
    Dim p
    p = CStr(Request.ServerVariables("APPL_MD_PATH") & "")
    AppRoot = "/"
End Function
%>
