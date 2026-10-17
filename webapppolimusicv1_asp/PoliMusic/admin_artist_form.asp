<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireAdmin
activeNav = "admin_artists"

Dim origMedia, id, isNew, rec, errs, fname, lname, imgPath, upErr, sql, params, outId, songIds, selSongs, allSongs, x
Set errs = Server.CreateObject("Scripting.Dictionary")
id = ToInt(QS("id"), 0)
isNew = (id = 0)

If Not isNew Then
    Set rec = QRow("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(id))
    If rec Is Nothing Then
        SetFlash "error", "El artista solicitado no existe."
        RedirectTo "admin_artists.asp"
    End If
    fname = rec("ARTIST_NAME") : lname = Nz(rec("ARTIST_LASTNAME"), "") : imgPath = Nz(rec("ARTIST_IMAGE"), "")
    origMedia = imgPath
    Set selSongs = IdSet(QList("SELECT ID_SONG FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(id)), "ID_SONG")
Else
    fname = "" : lname = "" : imgPath = ""
    Set selSongs = Server.CreateObject("Scripting.Dictionary")
End If

If IsPost() Then
    FormInit
    If gFormError <> "" Then errs("form") = gFormError
    If Not CsrfValid() Then errs("form") = "La sesión del formulario expiró. Vuelve a intentarlo."

    fname = FormText("artist_name")
    lname = FormText("artist_lastname")
    songIds = FormIds("songs")
    Set selSongs = Server.CreateObject("Scripting.Dictionary")
    For Each x In songIds : selSongs(CStr(x)) = True : Next

    If fname = "" Or Len(fname) > 100 Then errs("artist_name") = "El nombre es obligatorio (máx. 100 caracteres)."
    If Len(lname) > 100 Then errs("artist_lastname") = "Máximo 100 caracteres."

    If errs.Count = 0 Then
        If FormText("remove_image") = "1" Then imgPath = ""
        If FormHasFile("image") Then
            imgPath = FormSaveFile("image", "ArtistImages", IMAGE_EXTS, MAX_IMAGE_BYTES, upErr)
            If upErr <> "" Then errs("image") = upErr
        End If
    End If

    If errs.Count = 0 Then
        Dim lnameParam, imgParam
        lnameParam = Null
        If lname <> "" Then lnameParam = lname
        imgParam = Null
        If imgPath <> "" Then imgParam = imgPath
        If isNew Then
            sql = "INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (?, ?, ?); SET @out = SCOPE_IDENTITY(); "
            params = Array(fname, lnameParam, imgParam)
        Else
            sql = "SET @out = ?; UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? WHERE ID_ARTIST = @out; DELETE FROM TBL_PLAYER WHERE ID_ARTIST = @out; "
            params = Array(id, fname, lnameParam, imgParam)
        End If
        sql = sql & RelSql("TBL_PLAYER", "ID_ARTIST", "ID_SONG", songIds)
        params = ConcatArr(params, songIds)
        errs("form") = TryBatch(sql, params, outId)
        If errs("form") = "" Then
            If Not isNew And origMedia <> imgPath Then DeleteUploadedMedia origMedia
            SetFlash "success", IIfStr(isNew, "Artista «" & Trim(fname & " " & lname) & "» creado.", "Artista «" & Trim(fname & " " & lname) & "» actualizado.")
            RedirectTo "admin_artists.asp"
        End If
    End If
End If

pageTitle = IIfStr(isNew, "Nuevo artista", "Editar artista")
allSongs = QList("SELECT ID_SONG AS ID, SONG_NAME AS LABEL, '' AS SUB FROM TBL_SONG ORDER BY SONG_NAME", Array())

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
<a class="back-link" href="admin_artists.asp"><%= Icon("back") %>Artistas</a>
<% RenderAdminHeader pageTitle, "Datos del artista, imagen y canciones en las que participa.", "", "" %>
<% If errs.Exists("form") Then %><% If errs("form") <> "" Then %><div class="alert alert-error" role="alert"><%= Icon("alert") %><span><%= H(errs("form")) %></span></div><% End If %><% End If %>

<form method="post" enctype="multipart/form-data" class="form-layout" novalidate data-no-pjax>
  <%= CsrfField() %>
  <section class="panel">
    <h2 class="panel-title">Información</h2>
    <div class="field-row">
      <div class="field">
        <label for="artist_name">Nombre</label>
        <input class="input<%= ErrCls("artist_name") %>" id="artist_name" name="artist_name" maxlength="100" required value="<%= H(fname) %>">
        <%= FieldErr("artist_name") %>
      </div>
      <div class="field">
        <label for="artist_lastname">Apellido <span class="muted">(opcional)</span></label>
        <input class="input<%= ErrCls("artist_lastname") %>" id="artist_lastname" name="artist_lastname" maxlength="100" value="<%= H(lname) %>">
        <%= FieldErr("artist_lastname") %>
      </div>
    </div>
    <h2 class="panel-title">Canciones</h2>
    <% RenderChecklist "songs", allSongs, selSongs, "No hay canciones registradas." %>
  </section>
  <aside class="panel">
    <h2 class="panel-title">Imagen</h2>
    <div class="media-preview round"><%= Replace(Img(imgPath, "artist", "Vista previa", ""), "<img ", "<img id=""img-preview"" ") %></div>
    <label class="file-drop">
      <input type="file" name="image" accept="image/*" data-preview="#img-preview">
      <%= Icon("upload") %><span>Subir imagen</span><small>JPG, PNG, WEBP o GIF · máx. 5 MB</small>
    </label>
    <%= FieldErr("image") %>
    <% If imgPath <> "" Then %><label class="check-inline"><input type="checkbox" name="remove_image" value="1"> Quitar imagen actual</label><% End If %>
  </aside>
  <div class="form-actions">
    <a class="btn btn-ghost" href="admin_artists.asp">Cancelar</a>
    <button class="btn btn-primary" type="submit"><%= Icon("check") %><%= IIfStr(isNew, "Crear artista", "Guardar cambios") %></button>
  </div>
</form>
<!--#include file="includes/layout_bottom.asp"-->
