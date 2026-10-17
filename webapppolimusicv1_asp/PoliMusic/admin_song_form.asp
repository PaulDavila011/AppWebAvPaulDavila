<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireAdmin
activeNav = "admin_songs"
' Volver a la Song List si se abrió desde allí
Dim backUrl
backUrl = IIfStr(QS("r") = "songs", "songs.asp", "admin_songs.asp")

Dim origMedia, id, isNew, rec, errs, sname, spath, plays, upErr, sql, params, outId, artistIds, albumIds, selArtists, selAlbums, allArtists, allAlbums
Set errs = Server.CreateObject("Scripting.Dictionary")
id = ToInt(QS("id"), 0)
isNew = (id = 0)

If Not isNew Then
    Set rec = QRow("SELECT ID_SONG, SONG_NAME, SONG_PATH, ISNULL(PLAYS,0) AS PLAYS FROM TBL_SONG WHERE ID_SONG = ?", Array(id))
    If rec Is Nothing Then
        SetFlash "error", "La canción solicitada no existe."
        RedirectTo "admin_songs.asp"
    End If
    sname = rec("SONG_NAME") : spath = rec("SONG_PATH") : plays = rec("PLAYS")
    origMedia = spath
    Set selArtists = IdSet(QList("SELECT ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = ?", Array(id)), "ID_ARTIST")
    Set selAlbums = IdSet(QList("SELECT ID_ALBUM FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(id)), "ID_ALBUM")
Else
    sname = "" : spath = "" : plays = 0
    Set selArtists = Server.CreateObject("Scripting.Dictionary")
    Set selAlbums = Server.CreateObject("Scripting.Dictionary")
End If

If IsPost() Then
    FormInit
    If gFormError <> "" Then errs("form") = gFormError
    If Not CsrfValid() Then errs("form") = "La sesión del formulario expiró. Vuelve a intentarlo."

    sname = FormText("song_name")
    If FormText("song_path") <> "" Then spath = FormText("song_path")
    artistIds = FormIds("artists")
    albumIds = FormIds("albums")
    Set selArtists = Server.CreateObject("Scripting.Dictionary")
    Set selAlbums = Server.CreateObject("Scripting.Dictionary")
    Dim x
    For Each x In artistIds : selArtists(CStr(x)) = True : Next
    For Each x In albumIds : selAlbums(CStr(x)) = True : Next

    If sname = "" Or Len(sname) > 50 Then errs("song_name") = "El título es obligatorio (máx. 50 caracteres)."

    If errs.Count = 0 And FormHasFile("audio") Then
        spath = FormSaveFile("audio", "songFiles", AUDIO_EXTS, MAX_UPLOAD_BYTES, upErr)
        If upErr <> "" Then errs("audio") = upErr
    ElseIf spath = "" Then
        errs("audio") = "Sube un archivo de audio o indica la ruta del archivo."
    ElseIf Not ReTest("^(~/[A-Za-z0-9_\-./ ]+|https?://\S+)\.(mp3|m4a|ogg|wav)$", spath) Or InStr(spath, "..") > 0 Then
        errs("audio") = "Ruta no válida. Usa el formato ~/songFiles/archivo.mp3"
    ElseIf Len(spath) > 255 Then
        errs("audio") = "La ruta es demasiado larga."
    End If

    If errs.Count = 0 Then
        If isNew Then
            sql = "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, 0); SET @out = SCOPE_IDENTITY(); "
            params = Array(sname, spath)
        Else
            sql = "SET @out = ?; UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ?" & IIfStr(FormText("reset_plays") = "1", ", PLAYS = 0", "") & " WHERE ID_SONG = @out; " & _
                  "DELETE FROM TBL_PLAYER WHERE ID_SONG = @out; DELETE FROM TBL_TRACKLIST WHERE ID_SONG = @out; "
            params = Array(id, sname, spath)
        End If
        sql = sql & RelSql("TBL_PLAYER", "ID_SONG", "ID_ARTIST", artistIds) & RelSql("TBL_TRACKLIST", "ID_SONG", "ID_ALBUM", albumIds)
        params = ConcatArr(ConcatArr(params, artistIds), albumIds)
        errs("form") = TryBatch(sql, params, outId)
        If errs("form") = "" Then
            If Not isNew And origMedia <> spath Then DeleteUploadedMedia origMedia
            SetFlash "success", IIfStr(isNew, "Canción «" & sname & "» agregada.", "Canción «" & sname & "» actualizada.")
            RedirectTo backUrl
        End If
    End If
End If

pageTitle = IIfStr(isNew, "Nueva canción", "Editar canción")
allArtists = QList("SELECT ID_ARTIST AS ID, LTRIM(RTRIM(ARTIST_NAME + ' ' + ISNULL(ARTIST_LASTNAME,''))) AS LABEL, '' AS SUB FROM TBL_ARTIST ORDER BY ARTIST_NAME, ARTIST_LASTNAME", Array())
allAlbums = QList("SELECT ID_ALBUM AS ID, TITLE_ALBUM AS LABEL, CAST(RELEASE_YEAR AS varchar(10)) AS SUB FROM TBL_ALBUM ORDER BY TITLE_ALBUM", Array())

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
<a class="back-link" href="<%= backUrl %>"><%= Icon("back") %>Canciones</a>
<% RenderAdminHeader pageTitle, "Título, archivo de audio y relaciones con artistas y álbumes.", "", "" %>
<% If errs.Exists("form") Then %><% If errs("form") <> "" Then %><div class="alert alert-error" role="alert"><%= Icon("alert") %><span><%= H(errs("form")) %></span></div><% End If %><% End If %>

<form method="post" enctype="multipart/form-data" class="form-layout" novalidate data-no-pjax>
  <%= CsrfField() %>
  <section class="panel">
    <h2 class="panel-title">Información</h2>
    <div class="field">
      <label for="song_name">Título de la canción</label>
      <input class="input<%= ErrCls("song_name") %>" id="song_name" name="song_name" maxlength="50" required value="<%= H(sname) %>">
      <%= FieldErr("song_name") %>
    </div>
    <h2 class="panel-title">Artistas</h2>
    <% RenderChecklist "artists", allArtists, selArtists, "No hay artistas registrados. Crea uno primero." %>
    <h2 class="panel-title">Álbumes (tracklist)</h2>
    <% RenderChecklist "albums", allAlbums, selAlbums, "No hay álbumes registrados." %>
  </section>
  <aside class="panel">
    <h2 class="panel-title">Archivo de audio</h2>
    <% If spath <> "" Then %>
      <% If MediaExists(spath) Then %>
        <audio class="audio-preview" controls preload="none" src="<%= H(MediaUrl(spath)) %>"></audio>
      <% Else %>
        <div class="alert alert-warning"><%= Icon("alert") %><span>El archivo actual no existe en el servidor.</span></div>
      <% End If %>
    <% End If %>
    <label class="file-drop">
      <input type="file" name="audio" accept=".mp3,.m4a,.ogg,.wav,audio/*" data-file-name>
      <%= Icon("upload") %><span>Subir audio</span><small>MP3, M4A, OGG o WAV · máx. 50 MB</small>
    </label>
    <%= FieldErr("audio") %>
    <div class="field">
      <label for="song_path">o ruta del archivo</label>
      <input class="input mono" id="song_path" name="song_path" maxlength="255" value="<%= H(spath) %>" placeholder="~/songFiles/mi-cancion.mp3">
      <p class="hint">Si subes un archivo, la ruta se completa automáticamente.</p>
    </div>
    <% If Not isNew Then %>
    <div class="kv"><span>Reproducciones</span><strong><%= FmtNum(plays) %></strong></div>
    <label class="check-inline"><input type="checkbox" name="reset_plays" value="1"> Reiniciar contador de reproducciones</label>
    <% End If %>
  </aside>
  <div class="form-actions">
    <a class="btn btn-ghost" href="<%= backUrl %>">Cancelar</a>
    <button class="btn btn-primary" type="submit"><%= Icon("check") %><%= IIfStr(isNew, "Crear canción", "Guardar cambios") %></button>
  </div>
</form>
<!--#include file="includes/layout_bottom.asp"-->