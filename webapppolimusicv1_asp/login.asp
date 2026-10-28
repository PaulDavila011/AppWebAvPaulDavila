<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
'==============================================================================
' PoliMusic - Inicio de sesion
'==============================================================================
Dim errMsg, fUser, registered
Dim connErr

errMsg = ""
fUser = ""
registered = (ParseId(Request.QueryString("registered")) = 1)

' Si ya hay sesion activa, directo al inicio
If IsLoggedIn() Then
    Response.Redirect "default.asp"
    Response.End
End If

' --- Procesa el formulario ---
If UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) = "POST" Then

    fUser = CleanText(Request.Form("username"))
    Dim fPass
    fPass = CStr(Request.Form("password") & "")

    If Not ValidateCsrf() Then
        errMsg = "La sesion del formulario expiro. Vuelve a intentarlo."
    ElseIf IsBlank(fUser) Or Len(fPass) = 0 Then
        errMsg = "Escribe tu nombre de usuario y tu contrasena."
    Else
        On Error Resume Next
        Dim uid
        uid = Authenticate(fUser, fPass)
        If Err.Number <> 0 Then
            connErr = Err.Description
            Err.Clear
            uid = 0
        End If
        On Error GoTo 0

        If Len(CStr(connErr & "")) > 0 Then
            errMsg = "No se pudo comprobar las credenciales: " & connErr
        ElseIf uid > 0 Then
            ' Regresa a la pagina solicitada originalmente, si la hay
            Dim target
            target = CStr(Session("pm_return_to") & "")
            Session("pm_return_to") = ""

            If Len(target) > 0 And InStr(target, "login.asp") = 0 Then
                Response.Redirect target
            Else
                Response.Redirect "default.asp"
            End If
            Response.End
        Else
            errMsg = "Usuario o contrasena incorrectos. Verifica tus datos."
        End If
    End If
End If

Dim pageAction
pageAction = ""
%>
<!DOCTYPE html>
<html lang="es">
<head>
<% RenderHead "Iniciar sesion" %>
</head>
<body class="pm-auth">

    <!-- Panel decorativo -->
    <section class="pm-auth-aside" aria-hidden="true">
        <h2>Tu musica,<br>tu biblioteca,<br>tu ritmo.</h2>
        <p>PoliMusic reune todo el catalogo de la Escuela Politecnica Nacional
           en un solo lugar. Inicia sesion y sigue escuchando donde lo dejaste.</p>

        <ul class="pm-auth-features">
            <li>
                <span class="pm-feat-ico"><% Response.Write Icon("headphones") %></span>
                <div>
                    <strong>Reproductor continuo</strong>
                    <span>La musica no se corta mientras navegas por el catalogo.</span>
                </div>
            </li>
            <li>
                <span class="pm-feat-ico"><% Response.Write Icon("album") %></span>
                <div>
                    <strong>Albumes y artistas</strong>
                    <span>Consulta la discografia completa con un solo clic.</span>
                </div>
            </li>
            <li>
                <span class="pm-feat-ico"><% Response.Write Icon("playlist") %></span>
                <div>
                    <strong>Playlists propias</strong>
                    <span>Arma listas publicas o privadas con tus canciones favoritas.</span>
                </div>
            </li>
        </ul>
    </section>

    <!-- Panel del formulario -->
    <section class="pm-auth-panel">
        <div class="pm-auth-panel-inner">

            <div class="pm-auth-logo">
                <% Response.Write BrandMark() %>
                <div class="pm-brand-text">
                    <span class="pm-brand-name">PoliMusic</span>
                    <span class="pm-brand-sub">Escuela Politecnica Nacional</span>
                </div>
            </div>

            <h1 class="pm-auth-title">Bienvenido de vuelta</h1>
            <p class="pm-auth-sub">Inicia sesion para acceder a tu biblioteca musical.</p>

            <% If registered Then %>
            <div class="pm-flash pm-flash-success" role="status">
                <% Response.Write Icon("check") %>
                <span>Cuenta creada correctamente. Ya puedes iniciar sesion.</span>
            </div>
            <% End If %>

            <% If Len(errMsg) > 0 Then %>
            <div class="pm-flash pm-flash-error" role="alert">
                <% Response.Write Icon("alert") %>
                <span><% Response.Write H(errMsg) %></span>
            </div>
            <% End If %>

            <form method="post" action="login.asp" novalidate>
                <% Response.Write CsrfField() %>

                <div class="pm-field">
                    <label class="pm-label" for="username">
                        Nombre de usuario <span class="pm-req">*</span>
                    </label>
                    <input class="pm-input" type="text" id="username" name="username"
                           value="<% Response.Write H(fUser) %>"
                           placeholder="tu nombre de usuario"
                           autocomplete="username" required maxlength="50" autofocus>
                </div>

                <div class="pm-field">
                    <label class="pm-label" for="password">
                        Contrasena <span class="pm-req">*</span>
                    </label>
                    <input class="pm-input" type="password" id="password" name="password"
                           placeholder="tu contrasena"
                           autocomplete="current-password" required maxlength="50">
                </div>

                <button class="pm-btn pm-btn-primary pm-btn-block pm-btn-lg" type="submit">
                    <% Response.Write Icon("login") %>
                    <span>Iniciar sesion</span>
                </button>
            </form>

            <p class="pm-auth-foot">
                Todavia no tienes cuenta? <a href="register.asp">Registrate aqui</a>
            </p>

            <div class="pm-demo-creds">
                <p class="pm-demo-title">Cuentas de prueba</p>
                <div class="pm-demo-list">
                    <button class="pm-demo-btn" type="button"
                            onclick="pmFillLogin('userAdmin','12345678')">
                        <span><strong>userAdmin</strong> &middot; administrador</span>
                        <code>12345678</code>
                    </button>
                    <button class="pm-demo-btn" type="button"
                            onclick="pmFillLogin('paul.davila','12234567')">
                        <span><strong>paul.davila</strong> &middot; usuario</span>
                        <code>12234567</code>
                    </button>
                </div>
            </div>

        </div>
    </section>

<script src="assets/js/player.js"></script>
</body>
</html>
