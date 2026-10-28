<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
'==============================================================================
' PoliMusic - Registro de usuario
'==============================================================================
' Crea una cuenta con rol USER (USER_TYPE = 0). El registro publico nunca
' permite crear administradores: esa tarea queda reservada al panel de
' administracion.
'==============================================================================

Dim errors(), errCount
Dim fUser, fEmail, fBirth, fPass, fPass2

ReDim errors(20)
errCount = 0
fUser = "" : fEmail = "" : fBirth = "" : fPass = "" : fPass2 = ""

' Si ya hay sesion, no tiene sentido registrarse
If IsLoggedIn() Then
    Response.Redirect "default.asp"
    Response.End
End If

Sub AddError(msg)
    errors(errCount) = msg
    errCount = errCount + 1
End Sub

' --- Procesa el formulario ---
If UCase(CStr(Request.ServerVariables("REQUEST_METHOD"))) = "POST" Then

    fUser  = CleanText(Request.Form("username"))
    fEmail = CleanText(Request.Form("email"))
    fBirth = CleanText(Request.Form("birthday"))
    fPass  = CStr(Request.Form("password") & "")
    fPass2 = CStr(Request.Form("password2") & "")

    If Not ValidateCsrf() Then
        AddError "La sesion del formulario expiro. Vuelve a intentarlo."
    End If

    ' --- Nombre de usuario ---
    If IsBlank(fUser) Then
        AddError "El nombre de usuario es obligatorio."
    ElseIf Len(fUser) < 4 Then
        AddError "El nombre de usuario debe tener al menos 4 caracteres."
    ElseIf Len(fUser) > 50 Then
        AddError "El nombre de usuario no puede superar los 50 caracteres."
    ElseIf UsernameTaken(fUser, 0) Then
        AddError "Ese nombre de usuario ya esta en uso."
    End If

    ' --- Correo electronico ---
    If IsBlank(fEmail) Then
        AddError "El correo electronico es obligatorio."
    ElseIf Not IsValidEmail(fEmail) Then
        AddError "El correo electronico no tiene un formato valido."
    ElseIf Len(fEmail) > 50 Then
        AddError "El correo electronico no puede superar los 50 caracteres."
    ElseIf EmailTaken(fEmail, 0) Then
        AddError "Ya existe una cuenta registrada con ese correo."
    End If

    ' --- Contrasena ---
    If Len(fPass) < 6 Then
        AddError "La contrasena debe tener al menos 6 caracteres."
    ElseIf Len(fPass) > 50 Then
        AddError "La contrasena no puede superar los 50 caracteres."
    ElseIf StrComp(fPass, fPass2, 0) <> 0 Then
        AddError "Las dos contrasenas no coinciden."
    End If

    ' --- Fecha de nacimiento ---
    If Len(fBirth) > 0 Then
        If Not IsDate(fBirth) Then
            AddError "La fecha de nacimiento no es valida."
        ElseIf CDate(fBirth) > Date() Then
            AddError "La fecha de nacimiento no puede estar en el futuro."
        End If
    End If

    ' --- Alta del usuario ---
    If errCount = 0 Then
        Dim photoPath, newId

        photoPath = "~/userImages/default.jpg"

        ' Foto opcional
        If Request.Files.Count > 0 Then
            Dim uploaded
            uploaded = SaveUploadedFile("photo", "userImages", _
                                        "|jpg|jpeg|png|gif|webp|", "~/userImages/")
            If Len(uploaded) > 0 Then photoPath = uploaded
        End If

        newId = UserCreate(fUser, fPass, fEmail, fBirth, photoPath, ROLE_USER)

        If newId > 0 Then
            Response.Redirect "login.asp?registered=1"
            Response.End
        Else
            AddError "No se pudo crear la cuenta. Intentalo de nuevo."
        End If
    End If
End If
%>
<% Response.Write "<!DOCTYPE html>" & vbCrLf %>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Crear cuenta | PoliMusic</title>
<link rel="stylesheet" href="assets/css/polimusic.css">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Crect width='24' height='24' rx='6' fill='%237C3AED'/%3E%3Cpath d='M9 17V6l10-2v11' stroke='white' stroke-width='2' fill='none' stroke-linecap='round'/%3E%3Ccircle cx='7' cy='17' r='2.5' fill='white'/%3E%3Ccircle cx='17' cy='15' r='2.5' fill='white'/%3E%3C/svg%3E">
</head>
<body class="pm-auth">

    <section class="pm-auth-aside" aria-hidden="true">
        <h2>Crea tu cuenta<br>y empieza a escuchar.</h2>
        <p>Con una cuenta de PoliMusic puedes armar tus propias playlists,
           guardar tus canciones favoritas y consultar todo el catalogo.</p>

        <ul class="pm-auth-features">
            <li>
                <span class="pm-feat-ico"><% Response.Write Icon("sparkle") %></span>
                <div>
                    <strong>Registro en un minuto</strong>
                    <span>Solo necesitas un nombre de usuario, tu correo y una contrasena.</span>
                </div>
            </li>
            <li>
                <span class="pm-feat-ico"><% Response.Write Icon("heart") %></span>
                <div>
                    <strong>Tus playlists</strong>
                    <span>Organiza la musica como quieras, en listas publicas o privadas.</span>
                </div>
            </li>
            <li>
                <span class="pm-feat-ico"><% Response.Write Icon("album") %></span>
                <div>
                    <strong>Todo el catalogo</strong>
                    <span>Explora albumes y artistas y reproduce cualquier cancion.</span>
                </div>
            </li>
        </ul>
    </section>

    <section class="pm-auth-panel">
        <div class="pm-auth-panel-inner">

            <div class="pm-auth-logo">
                <% Response.Write BrandMark() %>
                <div class="pm-brand-text">
                    <span class="pm-brand-name">PoliMusic</span>
                    <span class="pm-brand-sub">Escuela Politecnica Nacional</span>
                </div>
            </div>

            <h1 class="pm-auth-title">Crear cuenta</h1>
            <p class="pm-auth-sub">Completa tus datos para registrarte.</p>

            <% If errCount > 0 Then %>
            <div class="pm-flash pm-flash-error" role="alert">
                <% Response.Write Icon("alert") %>
                <div>
                    <% If errCount = 1 Then %>
                    <strong><% Response.Write H(errors(0)) %></strong>
                    <% Else %>
                    <strong>Revisa los siguientes puntos:</strong>
                    <ul style="margin:6px 0 0;padding-left:18px">
                        <% Dim ei
                           For ei = 0 To errCount - 1 %>
                        <li><% Response.Write H(errors(ei)) %></li>
                        <% Next %>
                    </ul>
                    <% End If %>
                </div>
            </div>
            <% End If %>

            <form method="post" action="register.asp" enctype="multipart/form-data" novalidate>
                <% Response.Write CsrfField() %>

                <div class="pm-field">
                    <label class="pm-label" for="username">
                        Nombre de usuario <span class="pm-req">*</span>
                    </label>
                    <input class="pm-input" type="text" id="username" name="username"
                           value="<% Response.Write H(fUser) %>"
                           placeholder="minimo 4 caracteres"
                           autocomplete="username" required maxlength="50" autofocus>
                    <p class="pm-hint">Asi te veran los demas y con esto iniciaras sesion.</p>
                </div>

                <div class="pm-field">
                    <label class="pm-label" for="email">
                        Correo electronico <span class="pm-req">*</span>
                    </label>
                    <input class="pm-input" type="email" id="email" name="email"
                           value="<% Response.Write H(fEmail) %>"
                           placeholder="tu.correo@epn.edu.ec"
                           autocomplete="email" required maxlength="50">
                </div>

                <div class="pm-field">
                    <label class="pm-label" for="birthday">Fecha de nacimiento</label>
                    <input class="pm-input" type="date" id="birthday" name="birthday"
                           value="<% Response.Write H(fBirth) %>"
                           max="<% Response.Write Year(Date()) & "-" & Right("0" & Month(Date()),2) & "-" & Right("0" & Day(Date()),2) %>">
                    <p class="pm-hint">Opcional.</p>
                </div>

                <div class="pm-form-row">
                    <div class="pm-field">
                        <label class="pm-label" for="password">
                            Contrasena <span class="pm-req">*</span>
                        </label>
                        <input class="pm-input" type="password" id="password" name="password"
                               placeholder="minimo 6 caracteres"
                               autocomplete="new-password" required maxlength="50">
                    </div>

                    <div class="pm-field">
                        <label class="pm-label" for="password2">
                            Repite la contrasena <span class="pm-req">*</span>
                        </label>
                        <input class="pm-input" type="password" id="password2" name="password2"
                               placeholder="vuelve a escribirla"
                               autocomplete="new-password" required maxlength="50">
                    </div>
                </div>

                <div class="pm-field">
                    <label class="pm-label" for="photo">Foto de perfil</label>
                    <label class="pm-file" for="photo">
                        <% Response.Write Icon("upload") %>
                        <span class="pm-file-text">
                            <span class="pm-file-name">Ningun archivo seleccionado</span>
                            <br>JPG, PNG, GIF o WEBP. Maximo 10 MB.
                        </span>
                        <input type="file" id="photo" name="photo" accept="image/*">
                    </label>
                </div>

                <button class="pm-btn pm-btn-primary pm-btn-block pm-btn-lg" type="submit">
                    <% Response.Write Icon("user") %>
                    <span>Crear mi cuenta</span>
                </button>
            </form>

            <p class="pm-auth-foot">
                Ya tienes cuenta? <a href="login.asp">Inicia sesion</a>
            </p>

        </div>
    </section>

<script src="assets/js/player.js"></script>
</body>
</html>
