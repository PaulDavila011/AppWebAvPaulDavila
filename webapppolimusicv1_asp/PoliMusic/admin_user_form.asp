<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireAdmin
activeNav = "admin_users"

Dim origMedia, id, isNew, rec, errs, uname, email, bday, role, pwd, pwd2, photo, upErr, sql, params, outId, roles, k, isSelf, bdayParam
Set errs = Server.CreateObject("Scripting.Dictionary")
id = ToInt(QS("id"), 0)
isNew = (id = 0)
isSelf = (id = CurrentUserId())
Set roles = CatalogMap("1")

If Not isNew Then
    Set rec = QRow("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = ?", Array(id))
    If rec Is Nothing Then
        SetFlash "error", "El usuario solicitado no existe."
        RedirectTo "admin_users.asp"
    End If
    uname = rec("USERNAME") : email = rec("EMAIL") : bday = IsoDate(rec("BIRTHDAY")) : role = CLng(rec("USER_TYPE")) : photo = Nz(rec("USER_PHOTO"), "")
    origMedia = photo
Else
    uname = "" : email = "" : bday = "" : role = ROLE_USER : photo = DEFAULT_USER_PHOTO
End If

If IsPost() Then
    FormInit
    If gFormError <> "" Then errs("form") = gFormError
    If Not CsrfValid() Then errs("form") = "La sesión del formulario expiró. Vuelve a intentarlo."

    uname = FormText("username")
    email = LCase(FormText("email"))
    bday = FormText("birthday")
    role = ToInt(FormValue("role"), -1)
    pwd = FormValue("password")
    pwd2 = FormValue("password2")

    If Not ReTest("^[A-Za-z0-9._-]{3,50}$", uname) Then
        errs("username") = "Usa entre 3 y 50 caracteres: letras, números, punto, guion o guion bajo."
    ElseIf CLng(QScalar("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ? AND ID_USER <> ?", Array(uname, id))) > 0 Then
        errs("username") = "Ese nombre de usuario ya está en uso."
    End If
    If Len(email) > 50 Or Not ReTest("^[^@\s]+@[^@\s]+\.[^@\s]+$", email) Then
        errs("email") = "Correo electrónico no válido (máx. 50 caracteres)."
    ElseIf CLng(QScalar("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ? AND ID_USER <> ?", Array(email, id))) > 0 Then
        errs("email") = "Ya existe otra cuenta con ese correo."
    End If
    If bday <> "" Then
        If Not ReTest("^\d{4}-\d{2}-\d{2}$", bday) Then
            errs("birthday") = "Fecha no válida."
        ElseIf Not IsDate(bday) Then
            errs("birthday") = "Fecha no válida."
        ElseIf CDate(bday) > Date() Then
            errs("birthday") = "La fecha no puede ser futura."
        End If
    End If
    If Not roles.Exists(CStr(role)) Then errs("role") = "Selecciona un rol válido."
    If isSelf And role <> ROLE_ADMIN Then errs("role") = "No puedes quitarte el rol de administrador a ti mismo."

    If isNew Or pwd <> "" Then
        If Len(pwd) < 8 Or Len(pwd) > 50 Then
            errs("password") = "La contraseña debe tener entre 8 y 50 caracteres."
        ElseIf Not ReTest("[A-Za-z]", pwd) Or Not ReTest("\d", pwd) Then
            errs("password") = "Incluye al menos una letra y un número."
        ElseIf pwd2 <> pwd Then
            errs("password2") = "Las contraseñas no coinciden."
        End If
    End If

    If errs.Count = 0 Then
        If FormText("remove_photo") = "1" Then photo = DEFAULT_USER_PHOTO
        If FormHasFile("photo") Then
            photo = FormSaveFile("photo", "userImages", IMAGE_EXTS, MAX_IMAGE_BYTES, upErr)
            If upErr <> "" Then errs("photo") = upErr
        End If
    End If

    If errs.Count = 0 Then
        bdayParam = Null
        If bday <> "" Then bdayParam = bday
        If isNew Then
            sql = "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) VALUES (?, " & SQL_PASSWORD_VALUE & ", ?, TRY_CONVERT(datetime, ?, 23), ?, ?); SET @out = SCOPE_IDENTITY();"
            params = Array(uname, pwd, email, bdayParam, photo, role)
        Else
            sql = "SET @out = ?; UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = TRY_CONVERT(datetime, ?, 23), USER_PHOTO = ?, USER_TYPE = ? WHERE ID_USER = @out;"
            params = Array(id, uname, email, bdayParam, photo, role)
            If pwd <> "" Then
                sql = sql & " UPDATE TBL_USER SET PASSWORD = " & SQL_PASSWORD_VALUE & " WHERE ID_USER = @out;"
                params = ConcatArr(params, Array(pwd))
            End If
        End If
        errs("form") = TryBatch(sql, params, outId)
        If errs("form") = "" Then
            If Not isNew And origMedia <> photo Then DeleteUploadedMedia origMedia
            SetFlash "success", IIfStr(isNew, "Usuario «" & uname & "» creado.", "Usuario «" & uname & "» actualizado.")
            RedirectTo "admin_users.asp"
        End If
    End If
End If

pageTitle = IIfStr(isNew, "Nuevo usuario", "Editar usuario")

Function FieldErr(key)
    If errs.Exists(key) Then
        If errs(key) <> "" Then FieldErr = "<p class=""field-error"">" & H(errs(key)) & "</p>"
    End If
End Function
Function ErrCls(key)
    If errs.Exists(key) Then ErrCls = " is-invalid" Else ErrCls = ""
End Function
%>
<!--#include file="includes/layout_top.asp"-->
<a class="back-link" href="admin_users.asp"><%= Icon("back") %>Usuarios</a>
<% RenderAdminHeader pageTitle, IIfStr(isNew, "Registra una nueva cuenta y asigna su rol.", "Modifica los datos de " & uname & "."), "", "" %>
<% If errs.Exists("form") Then %><% If errs("form") <> "" Then %><div class="alert alert-error" role="alert"><%= Icon("alert") %><span><%= H(errs("form")) %></span></div><% End If %><% End If %>

<form method="post" enctype="multipart/form-data" class="form-layout" novalidate data-no-pjax>
  <%= CsrfField() %>
  <section class="panel">
    <h2 class="panel-title">Datos de la cuenta</h2>
    <div class="field-row">
      <div class="field">
        <label for="username">Nombre de usuario</label>
        <input class="input<%= ErrCls("username") %>" id="username" name="username" maxlength="50" required value="<%= H(uname) %>">
        <%= FieldErr("username") %>
      </div>
      <div class="field">
        <label for="email">Correo electrónico</label>
        <input class="input<%= ErrCls("email") %>" id="email" name="email" type="email" maxlength="50" required value="<%= H(email) %>">
        <%= FieldErr("email") %>
      </div>
    </div>
    <div class="field-row">
      <div class="field">
        <label for="birthday">Fecha de nacimiento</label>
        <input class="input<%= ErrCls("birthday") %>" id="birthday" name="birthday" type="date" max="<%= IsoDate(Date()) %>" value="<%= H(bday) %>">
        <%= FieldErr("birthday") %>
      </div>
      <div class="field">
        <label for="role">Rol</label>
        <select class="input<%= ErrCls("role") %>" id="role" name="role" <% If isSelf Then %>aria-describedby="role-hint"<% End If %>>
          <% For Each k In roles.Keys %>
          <option value="<%= H(k) %>"<% If CStr(role) = k Then %> selected<% End If %>><%= H(roles(k)) %></option>
          <% Next %>
        </select>
        <% If isSelf Then %><p class="hint" id="role-hint">Es tu propia cuenta: debes conservar el rol de administrador.</p><% End If %>
        <%= FieldErr("role") %>
      </div>
    </div>
    <h2 class="panel-title">Contraseña</h2>
    <% If Not isNew Then %><p class="hint">Déjala en blanco para conservar la contraseña actual.</p><% End If %>
    <div class="field-row">
      <div class="field">
        <label for="password"><%= IIfStr(isNew, "Contraseña", "Nueva contraseña") %></label>
        <input class="input<%= ErrCls("password") %>" id="password" name="password" type="password" autocomplete="new-password" maxlength="50" <% If isNew Then %>required<% End If %>>
        <%= FieldErr("password") %>
      </div>
      <div class="field">
        <label for="password2">Confirmar contraseña</label>
        <input class="input<%= ErrCls("password2") %>" id="password2" name="password2" type="password" autocomplete="new-password" maxlength="50">
        <%= FieldErr("password2") %>
      </div>
    </div>
    <p class="hint">Mínimo 8 caracteres, con al menos una letra y un número.</p>
  </section>
  <aside class="panel">
    <h2 class="panel-title">Foto de perfil</h2>
    <div class="media-preview round"><%= Replace(Img(photo, "user", "Vista previa", ""), "<img ", "<img id=""photo-preview"" ") %></div>
    <label class="file-drop">
      <input type="file" name="photo" accept="image/*" data-preview="#photo-preview">
      <%= Icon("upload") %><span>Subir imagen</span><small>JPG, PNG, WEBP o GIF · máx. 5 MB</small>
    </label>
    <%= FieldErr("photo") %>
    <% If Not isNew Then %><label class="check-inline"><input type="checkbox" name="remove_photo" value="1"> Restablecer foto predeterminada</label><% End If %>
  </aside>
  <div class="form-actions">
    <a class="btn btn-ghost" href="admin_users.asp">Cancelar</a>
    <button class="btn btn-primary" type="submit"><%= Icon("check") %><%= IIfStr(isNew, "Crear usuario", "Guardar cambios") %></button>
  </div>
</form>
<!--#include file="includes/layout_bottom.asp"-->
