<%
' =====================================================================
'  Autenticación, control de acceso basado en roles (RBAC) y CSRF
' =====================================================================

' SQL que compara la contraseña: texto plano (comparación sensible a mayúsculas)
' o hash SHA1, formato que ya tiene 'username1' en el script original.
Const SQL_PASSWORD_MATCH = "(u.PASSWORD = CONVERT(varchar(40), HASHBYTES('SHA1', CAST(? AS varchar(50))), 2) OR u.PASSWORD COLLATE Latin1_General_BIN = ?)"
' Por ahora las contraseñas nuevas se guardan sin cifrar (texto plano).
Const SQL_PASSWORD_VALUE = "CAST(? AS varchar(50))"

Dim gUserChecked
gUserChecked = False

Function IsLoggedIn()
    IsLoggedIn = (ToInt(Session("uid"), 0) > 0)
End Function

Function CurrentUserId()
    CurrentUserId = ToInt(Session("uid"), 0)
End Function

Function IsAdmin()
    IsAdmin = IsLoggedIn() And (ToInt(Session("role"), -1) = ROLE_ADMIN)
End Function

Sub SignIn(u)
    Session("uid") = CLng(u("ID_USER"))
    Session("uname") = u("USERNAME")
    Session("email") = u("EMAIL")
    Session("role") = CLng(u("USER_TYPE"))
    Session("photo") = Nz(u("USER_PHOTO"), "")
    Session("login_fails") = 0
    Session("csrf") = ""
End Sub

Sub SignOut()
    Session.Contents.RemoveAll
    Session.Abandon
End Sub

' Revalida al usuario contra la BD en cada petición: si fue eliminado
' o cambió de rol, la sesión se actualiza de inmediato.
Function RefreshSessionUser()
    Dim u
    RefreshSessionUser = False
    If Not IsLoggedIn() Then Exit Function
    If gUserChecked Then
        RefreshSessionUser = True
        Exit Function
    End If
    Set u = QRow("SELECT ID_USER, USERNAME, EMAIL, USER_TYPE, USER_PHOTO FROM TBL_USER WHERE ID_USER = ?", Array(CurrentUserId()))
    If u Is Nothing Then
        SignOut
        Exit Function
    End If
    Session("uname") = u("USERNAME")
    Session("email") = u("EMAIL")
    Session("role") = CLng(u("USER_TYPE"))
    Session("photo") = Nz(u("USER_PHOTO"), "")
    gUserChecked = True
    RefreshSessionUser = True
End Function

Function CurrentPageUrl()
    Dim page, q
    page = Request.ServerVariables("SCRIPT_NAME")
    page = Mid(page, InStrRev(page, "/") + 1)
    q = Request.ServerVariables("QUERY_STRING")
    If q <> "" Then page = page & "?" & q
    CurrentPageUrl = page
End Function

' --- Guardas de acceso (llamar al inicio de cada página protegida) ---

Sub RequireLogin()
    If Not RefreshSessionUser() Then
        Dim r
        r = ""
        If UCase(Request.ServerVariables("REQUEST_METHOD")) = "GET" Then r = "?r=" & Server.URLEncode(CurrentPageUrl())
        RedirectTo "login.asp" & r
    End If
End Sub

Sub RequireAdmin()
    RequireLogin
    If Not IsAdmin() Then RenderForbidden
End Sub

' Variante para endpoints AJAX: responde JSON en lugar de redirigir
Sub RequireLoginApi()
    If Not RefreshSessionUser() Then
        Response.Clear
        Response.Status = "401 Unauthorized"
        Response.ContentType = "application/json"
        Response.Write "{""ok"":false,""error"":""unauthorized""}"
        DbClose
        Response.End
    End If
End Sub

Sub RenderForbidden()
    Response.Clear
    Response.Status = "403 Forbidden"
    Response.Write "<!DOCTYPE html><html lang=""es""><head><meta charset=""utf-8""><meta name=""viewport"" content=""width=device-width, initial-scale=1"">"
    Response.Write "<title>Acceso denegado · " & APP_NAME & "</title><link rel=""icon"" href=""assets/img/logo.svg""><link rel=""stylesheet"" href=""assets/css/app.css""></head>"
    Response.Write "<body class=""page-center""><main id=""content"" class=""error-card""><div class=""error-code"">403</div>"
    Response.Write "<h1>Acceso denegado</h1><p>Tu cuenta (<strong>" & H(Session("uname")) & "</strong>) no tiene permisos de administrador para esta sección.</p>"
    Response.Write "<a class=""btn btn-primary"" href=""default.asp"" data-no-pjax>Volver al inicio</a></main></body></html>"
    DbClose
    Response.End
End Sub

' ---------------------------- CSRF ----------------------------

Function CsrfToken()
    If Nz(Session("csrf"), "") = "" Then
        ' Aleatorio criptográfico generado por SQL Server
        Session("csrf") = QScalar("SELECT CONVERT(varchar(64), CRYPT_GEN_RANDOM(32), 2)", Array())
    End If
    CsrfToken = Session("csrf")
End Function

Function CsrfField()
    CsrfField = "<input type=""hidden"" name=""csrf"" value=""" & H(CsrfToken()) & """>"
End Function

Function CsrfValid()
    Dim t
    t = FormValue("csrf")
    CsrfValid = (Nz(Session("csrf"), "") <> "" And t = Session("csrf"))
End Function

Sub RequireCsrf()
    If Not CsrfValid() Then
        SetFlash "error", "La sesión del formulario expiró. Inténtalo de nuevo."
        RedirectTo SafeReturnUrl(CurrentPageUrl())
    End If
End Sub
%>
