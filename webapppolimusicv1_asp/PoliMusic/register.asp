<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
Dim errs, uname, email, bday, pwd, pwd2, newId, u
pageTitle = "Crear cuenta"
Set errs = Server.CreateObject("Scripting.Dictionary")

If RefreshSessionUser() Then RedirectTo "default.asp"

If IsPost() Then
    uname = FormText("username")
    email = LCase(FormText("email"))
    bday = FormText("birthday")
    pwd = FormValue("password")
    pwd2 = FormValue("password2")

    If Not CsrfValid() Then errs("form") = "La sesión del formulario expiró. Vuelve a intentarlo."

    If Not ReTest("^[A-Za-z0-9._-]{3,50}$", uname) Then
        errs("username") = "Usa entre 3 y 50 caracteres: letras, números, punto, guion o guion bajo."
    ElseIf CLng(QScalar("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ?", Array(uname))) > 0 Then
        errs("username") = "Ese nombre de usuario ya está en uso."
    End If

    If Len(email) > 50 Or Not ReTest("^[^@\s]+@[^@\s]+\.[^@\s]+$", email) Then
        errs("email") = "Ingresa un correo electrónico válido (máx. 50 caracteres)."
    ElseIf CLng(QScalar("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ?", Array(email))) > 0 Then
        errs("email") = "Ya existe una cuenta con ese correo."
    End If

    If bday <> "" Then
        If Not ReTest("^\d{4}-\d{2}-\d{2}$", bday) Then
            errs("birthday") = "Fecha no válida."
        ElseIf Not IsDate(bday) Then
            errs("birthday") = "Fecha no válida."
        ElseIf CDate(bday) > Date() Or Year(CDate(bday)) < 1900 Then
            errs("birthday") = "La fecha de nacimiento no es válida."
        End If
    End If

    If Len(pwd) < 8 Or Len(pwd) > 50 Then
        errs("password") = "La contraseña debe tener entre 8 y 50 caracteres."
    ElseIf Not ReTest("[A-Za-z]", pwd) Or Not ReTest("\d", pwd) Then
        errs("password") = "Incluye al menos una letra y un número."
    End If
    If pwd2 <> pwd Then errs("password2") = "Las contraseñas no coinciden."

    If errs.Count = 0 Then
        Dim bdayParam
        bdayParam = Null
        If bday <> "" Then bdayParam = bday
        ' Todo registro público crea un usuario normal (USER_TYPE = 0). Solo un admin puede otorgar el rol ADMIN.
        newId = QInsert("INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) VALUES (?, " & SQL_PASSWORD_VALUE & ", ?, TRY_CONVERT(datetime, ?, 23), ?, ?)", _
                        Array(uname, pwd, email, bdayParam, DEFAULT_USER_PHOTO, ROLE_USER))
        Set u = QRow("SELECT ID_USER, USERNAME, EMAIL, USER_TYPE, USER_PHOTO FROM TBL_USER WHERE ID_USER = ?", Array(newId))
        SignIn u
        SetFlash "success", "¡Bienvenido a PoliMusic, " & uname & "! Tu cuenta fue creada."
        RedirectTo "default.asp"
    End If
End If

Function FieldErr(key)
    If errs.Exists(key) Then FieldErr = "<p class=""field-error"">" & H(errs(key)) & "</p>" Else FieldErr = ""
End Function
Function ErrCls(key)
    If errs.Exists(key) Then ErrCls = " is-invalid" Else ErrCls = ""
End Function
%>
<!--#include file="includes/auth_top.asp"-->
      <h1>Crea tu cuenta</h1>
      <p class="muted">Regístrate gratis y empieza a escuchar.</p>
      <% If errs.Exists("form") Then %><div class="alert alert-error" role="alert"><%= Icon("alert") %><span><%= H(errs("form")) %></span></div><% End If %>
      <form method="post" action="register.asp" class="form" data-busy novalidate>
        <%= CsrfField() %>
        <div class="field">
          <label for="username">Nombre de usuario</label>
          <input class="input<%= ErrCls("username") %>" id="username" name="username" type="text" autocomplete="username" required maxlength="50" pattern="[A-Za-z0-9._\-]{3,50}" value="<%= H(uname) %>" placeholder="ej. maria.lopez" autofocus>
          <%= FieldErr("username") %>
        </div>
        <div class="field">
          <label for="email">Correo electrónico</label>
          <input class="input<%= ErrCls("email") %>" id="email" name="email" type="email" autocomplete="email" required maxlength="50" value="<%= H(email) %>" placeholder="tu@correo.com">
          <%= FieldErr("email") %>
        </div>
        <div class="field">
          <label for="birthday">Fecha de nacimiento <span class="muted">(opcional)</span></label>
          <input class="input<%= ErrCls("birthday") %>" id="birthday" name="birthday" type="date" max="<%= IsoDate(Date()) %>" value="<%= H(bday) %>">
          <%= FieldErr("birthday") %>
        </div>
        <div class="field-row">
          <div class="field">
            <label for="password">Contraseña</label>
            <div class="input-group">
              <input class="input<%= ErrCls("password") %>" id="password" name="password" type="password" autocomplete="new-password" required minlength="8" maxlength="50">
              <button type="button" class="input-addon" data-toggle-password="password" aria-pressed="false">Mostrar</button>
            </div>
            <%= FieldErr("password") %>
          </div>
          <div class="field">
            <label for="password2">Confirmar contraseña</label>
            <input class="input<%= ErrCls("password2") %>" id="password2" name="password2" type="password" autocomplete="new-password" required maxlength="50">
            <%= FieldErr("password2") %>
          </div>
        </div>
        <p class="hint">Mínimo 8 caracteres, con al menos una letra y un número.</p>
        <button type="submit" class="btn btn-primary btn-block btn-lg">Crear cuenta</button>
      </form>
      <p class="auth-switch">¿Ya tienes cuenta? <a href="login.asp">Inicia sesión</a></p>
<!--#include file="includes/auth_bottom.asp"-->
