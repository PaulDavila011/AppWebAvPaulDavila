<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireAdmin
activeNav = "admin_playlists"

Dim origMedia, id, isNew, rec, errs, title, owner, ptype, photo, upErr, sql, params, outId, songIds, selSongs, allSongs, users, types, k, x, u
Set errs = Server.CreateObject("Scripting.Dictionary")
id = ToInt(QS("id"), 0)
isNew = (id = 0)
Set types = CatalogMap("2")

If Not isNew Then
    Set rec = QRow("SELECT ID_PLAYLIST, ID_USER, TITLE, TYPE, PHOTO, CREATION_DATE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(id))
    If rec Is Nothing Then
        SetFlash "error", "La playlist solicitada no existe."
        RedirectTo "admin_playlists.asp"
    End If
    title = rec("TITLE") : owner = CLng(rec("ID_USER")) : ptype = CLng(rec("TYPE")) : photo = Nz(rec("PHOTO"), "")
    origMedia = photo
    Set selSongs = IdSet(QList("SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array(id)), "ID_SONG")
Else
    title = "" : owner = CurrentUserId() : ptype = PLAYLIST_PUBLIC : photo = DEFAULT_PLAYLIST_PHOTO
    Set selSongs = Server.CreateObject("Scripting.Dictionary")
End If

If IsPost() Then
    FormInit
    If gFormError <> "" Then errs("form") = gFormError
    If Not CsrfValid() Then errs("form") = "La sesión del formulario expiró. Vuelve a intentarlo."

    title = FormText("title")
    owner = ToInt(FormValue("owner"), 0)
    ptype = ToInt(FormValue("type"), -1)
    songIds = FormIds("songs")
    Set selSongs = Server.CreateObject("Scripting.Dictionary")
    For Each x In songIds : selSongs(CStr(x)) = True : Next

    If title = "" Or Len(title) > 50 Then errs("title") = "El título es obligatorio (máx. 50 caracteres)."
    If CLng(QScalar("SELECT COUNT(*) FROM TBL_USER WHERE ID_USER = ?", Array(owner))) = 0 Then errs("owner") = "Selecciona un propietario válido."
    If Not types.Exists(CStr(ptype)) Then errs("type") = "Selecciona la visibilidad."

    If errs.Count = 0 Then
        If FormText("remove_photo") = "1" Then photo = DEFAULT_PLAYLIST_PHOTO
        If FormHasFile("photo") Then
            photo = FormSaveFile("photo", "playlistImages", IMAGE_EXTS, MAX_IMAGE_BYTES, upErr)
            If upErr <> "" Then errs("photo") = upErr
        End If
    End If

    If errs.Count = 0 Then
        If isNew Then
            sql = "INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) VALUES (?, ?, GETDATE(), ?, ?); SET @out = SCOPE_IDENTITY(); "
            params = Array(owner, title, ptype, photo)
        Else
            sql = "SET @out = ?; UPDATE TBL_PLAYLIST SET ID_USER = ?, TITLE = ?, TYPE = ?, PHOTO = ? WHERE ID_PLAYLIST = @out; DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = @out; "
            params = Array(id, owner, title, ptype, photo)
        End If
        sql = sql & RelSql("TBL_PLAYLIST_SONG", "ID_PLAYLIST", "ID_SONG", songIds)
        params = ConcatArr(params, songIds)
        errs("form") = TryBatch(sql, params, outId)
        If errs("form") = "" Then
            If Not isNew And origMedia <> photo Then DeleteUploadedMedia origMedia
            SetFlash "success", IIfStr(isNew, "Playlist «" & title & "» creada.", "Playlist «" & title & "» actualizada.")
            RedirectTo "admin_playlists.asp"
        End If
    End If
End If

pageTitle = IIfStr(isNew, "Nueva playlist", "Editar playlist")
users = QList("SELECT ID_USER, USERNAME, USER_TYPE FROM TBL_USER ORDER BY USERNAME", Array())
allSongs = QList("SELECT x.ID_SONG AS ID, x.SONG_NAME AS LABEL, x.ARTISTS AS SUB FROM (" & SongSelectSql() & ") x ORDER BY x.SONG_NAME", Array())

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
<a class="back-link" href="admin_playlists.asp"><%= Icon("back") %>Playlists</a>
<% RenderAdminHeader pageTitle, "Título, propietario, visibilidad, portada y canciones.", "", "" %>
<% If errs.Exists("form") Then %><% If errs("form") <> "" Then %><div class="alert alert-error" role="alert"><%= Icon("alert") %><span><%= H(errs("form")) %></span></div><% End If %><% End If %>

<form method="post" enctype="multipart/form-data" class="form-layout" novalidate data-no-pjax>
  <%= CsrfField() %>
  <section class="panel">
    <h2 class="panel-title">Información</h2>
    <div class="field">
      <label for="title">Título</label>
      <input class="input<%= ErrCls("title") %>" id="title" name="title" maxlength="50" required value="<%= H(title) %>">
      <%= FieldErr("title") %>
    </div>
    <div class="field-row">
      <div class="field">
        <label for="owner">Propietario</label>
        <select class="input<%= ErrCls("owner") %>" id="owner" name="owner">
          <% For Each u In users %>
          <option value="<%= u("ID_USER") %>"<% If CLng(u("ID_USER")) = owner Then %> selected<% End If %>><%= H(u("USERNAME")) %><% If CLng(u("USER_TYPE")) = ROLE_ADMIN Then %> (admin)<% End If %></option>
          <% Next %>
        </select>
        <%= FieldErr("owner") %>
      </div>
      <div class="field">
        <label for="type">Visibilidad</label>
        <select class="input<%= ErrCls("type") %>" id="type" name="type">
          <% For Each k In types.Keys %>
          <option value="<%= H(k) %>"<% If CStr(ptype) = k Then %> selected<% End If %>><%= H(IIfStr(k = CStr(PLAYLIST_PUBLIC), "Pública", IIfStr(k = CStr(PLAYLIST_PRIVATE), "Privada", types(k)))) %></option>
          <% Next %>
        </select>
        <p class="hint">Las privadas solo las ven su propietario y los administradores.</p>
        <%= FieldErr("type") %>
      </div>
    </div>
    <h2 class="panel-title">Canciones</h2>
    <% RenderChecklist "songs", allSongs, selSongs, "No hay canciones registradas." %>
  </section>
  <aside class="panel">
    <h2 class="panel-title">Portada</h2>
    <div class="media-preview"><%= Replace(Img(photo, "cover", "Vista previa", ""), "<img ", "<img id=""img-preview"" ") %></div>
    <label class="file-drop">
      <input type="file" name="photo" accept="image/*" data-preview="#img-preview">
      <%= Icon("upload") %><span>Subir portada</span><small>JPG, PNG, WEBP o GIF · máx. 5 MB</small>
    </label>
    <%= FieldErr("photo") %>
    <% If Not isNew Then %>
    <label class="check-inline"><input type="checkbox" name="remove_photo" value="1"> Restablecer portada predeterminada</label>
    <div class="kv"><span>Creada</span><strong><%= FmtDate(rec("CREATION_DATE")) %></strong></div>
    <% End If %>
  </aside>
  <div class="form-actions">
    <a class="btn btn-ghost" href="admin_playlists.asp">Cancelar</a>
    <button class="btn btn-primary" type="submit"><%= Icon("check") %><%= IIfStr(isNew, "Crear playlist", "Guardar cambios") %></button>
  </div>
</form>
<!--#include file="includes/layout_bottom.asp"-->
