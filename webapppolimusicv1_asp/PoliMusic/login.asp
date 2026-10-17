<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
Dim errMsg, loginVal, returnUrl, u, fails, lockUntil, isLocked, okMsg
pageTitle = "Iniciar sesión"
errMsg = ""
loginVal = ""

If RefreshSessionUser() Then RedirectTo "default.asp"

returnUrl = SafeReturnUrl(QS("r"))
okMsg = ""
If QS("bye") = "1" Then okMsg = "Cerraste sesión correctamente. ¡Hasta pronto!"
If QS("r") <> "" And Not IsPost() Then okMsg = "Inicia sesión para continuar."

If IsPost() Then
    returnUrl = SafeReturnUrl(FormText("r"))
    loginVal = FormText("username")
    lockUntil = Session("lock_until")
    isLocked = False
    If IsDate(lockUntil) Then isLocked = (Now() < CDate(lockUntil))
    If isLocked Then
        errMsg = "Demasiados intentos fallidos. Espera " & LOGIN_LOCK_MINUTES & " minutos e inténtalo de nuevo."
    ElseIf Not CsrfValid() Then
        errMsg = "La sesión del formulario expiró. Vuelve a intentarlo."
    ElseIf loginVal = "" Or FormValue("password") = "" Then
        errMsg = "Ingresa tu usuario (o correo) y tu contraseña."
    Else
        Set u = QRow("SELECT TOP 1 u.ID_USER, u.USERNAME, u.EMAIL, u.USER_TYPE, u.USER_PHOTO FROM TBL_USER u " & _
                     "WHERE (u.USERNAME = ? OR u.EMAIL = ?) AND " & SQL_PASSWORD_MATCH, _
                     Array(loginVal, loginVal, FormValue("password"), FormValue("password")))
        If u Is Nothing Then
            fails = ToInt(Session("login_fails"), 0) + 1
            Session("login_fails") = fails
            If fails >= LOGIN_MAX_ATTEMPTS Then
                Session("lock_until") = DateAdd("n", LOGIN_LOCK_MINUTES, Now())
                Session("login_fails") = 0
                errMsg = "Demasiados intentos fallidos. Tu acceso se bloqueó por " & LOGIN_LOCK_MINUTES & " minutos."
            Else
                errMsg = "Usuario o contraseña incorrectos."
            End If
        Else
            SignIn u
            Session("lock_until") = ""
            SetFlash "success", Greeting() & ", " & u("USERNAME") & ". ¡Qué bueno verte!"
            RedirectTo returnUrl
        End If
    End If
End If
%>
<!--#include file="includes/auth_top.asp"-->
      <h1>Inicia sesión</h1>
      <p class="muted">Ingresa a tu cuenta para seguir escuchando.</p>
      <% If errMsg <> "" Then %><div class="alert alert-error" role="alert"><%= Icon("alert") %><span><%= H(errMsg) %></span></div><% End If %>
      <% If errMsg = "" And okMsg <> "" Then %><div class="alert alert-info" role="status"><%= Icon("check") %><span><%= H(okMsg) %></span></div><% End If %>
      <form method="post" action="login.asp" class="form" data-busy novalidate>
        <%= CsrfField() %>
        <input type="hidden" name="r" value="<%= H(returnUrl) %>">
        <div class="field">
          <label for="username">Usuario o correo electrónico</label>
          <input class="input" id="username" name="username" type="text" autocomplete="username" required maxlength="50" value="<%= H(loginVal) %>" autofocus>
        </div>
        <div class="field">
          <label for="password">Contraseña</label>
          <div class="input-group">
            <input class="input" id="password" name="password" type="password" autocomplete="current-password" required maxlength="50">
            <button type="button" class="input-addon" data-toggle-password="password" aria-pressed="false">Mostrar</button>
          </div>
        </div>
        <button type="submit" class="btn btn-primary btn-block btn-lg">Iniciar sesión</button>
      </form>
      <p class="auth-switch">¿No tienes cuenta? <a href="register.asp">Regístrate gratis</a></p>
<!--#include file="includes/auth_bottom.asp"-->
