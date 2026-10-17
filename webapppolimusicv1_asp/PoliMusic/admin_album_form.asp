<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireAdmin
activeNav = "admin_albums"

Dim origMedia, id, isNew, rec, errs, title, yr, imgPath, upErr, sql, params, outId, songIds, selSongs, allSongs, x, yrParam, imgParam
Set errs = Server.CreateObject("Scripting.Dictionary")
id = ToInt(QS("id"), 0)
isNew = (id = 0)

If Not isNew Then
    Set rec = QRow("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(id))
    If rec Is Nothing Then
        SetFlash "error", "El álbum solicitado no existe."
        RedirectTo "admin_albums.asp"
    End If
    title = rec("TITLE_ALBUM") : yr = Nz(rec("RELEASE_YEAR"), "") : imgPath = Nz(rec("ALBUM_IMAGE_PATH"), "")
    origMedia = imgPath
    Set selSongs = IdSet(QList("SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(id)), "ID_SONG")
Else
    title = "" : yr = CStr(Year(Date())) : imgPath = ""
    Set selSongs = Server.CreateObject("Scripting.Dictionary")
End If

If IsPost() Then
    FormInit
    If gFormError <> "" Then errs("form") = gFormError
    If Not CsrfValid() Then errs("form") = "La sesión del formulario expiró. Vuelve a intentarlo."

    title = FormText("title")
    yr = FormText("release_year")
    songIds = FormIds("songs")
    Set selSongs = Server.CreateObject("Scripting.Dictionary")
    For Each x In songIds : selSongs(CStr(x)) = True : Next

    If title = "" Or Len(title) > 50 Then errs("title") = "El título es obligatorio (máx. 50 caracteres)."
    If yr <> "" Then
        If ToInt(yr, 0) < 1900 Or ToInt(yr, 0) > Year(Date()) + 1 Then errs("release_year") = "Ingresa un año entre 1900 y " & (Year(Date()) + 1) & "."
    End If

    If errs.Count = 0 Then
        If FormText("remove_image") = "1" Then imgPath = ""
        If FormHasFile("image") Then
            imgPath = FormSaveFile("image", "albumImages", IMAGE_EXTS, MAX_IMAGE_BYTES, upErr)
            If upErr <> "" Then errs("image") = upErr
        End If
    End If

    If errs.Count = 0 Then
        yrParam = Null
        If yr <> "" Then yrParam = ToInt(yr, 0)
        imgParam = Null
        If imgPath <> "" Then imgParam = imgPath
        If isNew Then
            sql = "INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (?, ?, ?); SET @out = SCOPE_IDENTITY(); "
            params = Array(title, yrParam, imgParam)
        Else
            sql = "SET @out = ?; UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = @out; DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = @out; "
            params = Array(id, title, yrParam, imgParam)
        End If
        sql = sql & RelSql("TBL_TRACKLIST", "ID_ALBUM", "ID_SONG", songIds)
        params = ConcatArr(params, songIds)
        errs("form") = TryBatch(sql, params, outId)
        If errs("form") = "" Then
            If Not isNew And origMedia <> imgPath Then DeleteUploadedMedia origMedia
            SetFlash "success", IIfStr(isNew, "Álbum «" & title & "» creado.", "Álbum «" & title & "» actualizado.")
            RedirectTo "admin_albums.asp"
        End If
    End If
End If

pageTitle = IIfStr(isNew, "Nuevo álbum", "Editar álbum")
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
<a class="back-link" href="admin_albums.asp"><%= Icon("back") %>Álbumes</a>
<% RenderAdminHeader pageTitle, "Título, año de lanzamiento, portada y tracklist.", "", "" %>
<% If errs.Exists("form") Then %><% If errs("form") <> "" Then %><div class="alert alert-error" role="alert"><%= Icon("alert") %><span><%= H(errs("form")) %></span></div><% End If %><% End If %>

<form method="post" enctype="multipart/form-data" class="form-layout" novalidate data-no-pjax>
  <%= CsrfField() %>
  <section class="panel">
    <h2 class="panel-title">Información</h2>
    <div class="field-row">
      <div class="field">
        <label for="title">Título del álbum</label>
        <input class="input<%= ErrCls("title") %>" id="title" name="title" maxlength="50" required value="<%= H(title) %>">
        <%= FieldErr("title") %>
      </div>
      <div class="field field-sm">
        <label for="release_year">Año</label>
        <input class="input<%= ErrCls("release_year") %>" id="release_year" name="release_year" type="number" min="1900" max="<%= Year(Date()) + 1 %>" value="<%= H(yr) %>">
        <%= FieldErr("release_year") %>
      </div>
    </div>
    <h2 class="panel-title">Tracklist</h2>
    <% RenderChecklist "songs", allSongs, selSongs, "No hay canciones registradas." %>
  </section>
  <aside class="panel">
    <h2 class="panel-title">Portada</h2>
    <div class="media-preview"><%= Replace(Img(imgPath, "cover", "Vista previa", ""), "<img ", "<img id=""img-preview"" ") %></div>
    <label class="file-drop">
      <input type="file" name="image" accept="image/*" data-preview="#img-preview">
      <%= Icon("upload") %><span>Subir portada</span><small>JPG, PNG, WEBP o GIF · máx. 5 MB</small>
    </label>
    <%= FieldErr("image") %>
    <% If imgPath <> "" Then %><label class="check-inline"><input type="checkbox" name="remove_image" value="1"> Quitar portada actual</label><% End If %>
  </aside>
  <div class="form-actions">
    <a class="btn btn-ghost" href="admin_albums.asp">Cancelar</a>
    <button class="btn btn-primary" type="submit"><%= Icon("check") %><%= IIfStr(isNew, "Crear álbum", "Guardar cambios") %></button>
  </div>
</form>
<!--#include file="includes/layout_bottom.asp"-->
