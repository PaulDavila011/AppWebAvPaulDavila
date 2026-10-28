<%
'==============================================================================
' PoliMusic - Arranque comun de paginas
'==============================================================================
' Incluye todas las capas y expone helpers de alto nivel para que cada pagina
' sea corta: configurar la sesion, imprimir cabecera, cuerpo y cierre.
'==============================================================================
%>
<!--#include virtual="/config/config.asp" -->
<!--#include virtual="/config/db.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/auth.asp" -->
<!--#include virtual="/includes/repository.asp" -->
<!--#include virtual="/includes/crud.asp" -->
<!--#include virtual="/includes/layout.asp" -->
<%
'==============================================================================
' RenderJsConfig() - Puente de datos entre el servidor y player.js.
'==============================================================================
Sub RenderJsConfig()
%>
<script>
window.PM_BASE = "./";
window.PM_CSRF = "<% Response.Write JsArg(CsrfToken()) %>";
window.PM_USER = "<% Response.Write JsArg(CurrentUserName()) %>";
window.PM_IS_ADMIN = <% If IsAdmin() Then Response.Write "true" Else Response.Write "false" %>;
window.PM_LOGGED = <% If IsLoggedIn() Then Response.Write "true" Else Response.Write "false" %>;
</script>
<%
End Sub

'==============================================================================
' JsonEscape() - (definida en functions.asp, se reutiliza aqui para claridad)
'==============================================================================

'==============================================================================
' TracksJson() - Serializa un Recordset de canciones a un array JSON para el
' reproductor. Se escapan todos los valores de texto.
'
' OJO: esta funcion RECORRE el Recordset y lo deja al final. Si la pagina
' todavia necesita iterarlo, usa TracksJsonSafe(), que guarda y restaura la
' posicion.
'==============================================================================
Function TracksJson(rs)
    TracksJson = TracksJsonCore(rs, False)
End Function

'==============================================================================
' TracksJsonSafe() - Igual que TracksJson pero deja el cursor donde estaba,
' para poder seguir recorriendo el mismo Recordset despues.
'==============================================================================
Function TracksJsonSafe(rs)
    TracksJsonSafe = TracksJsonCore(rs, True)
End Function

'==============================================================================
' TracksJsonCore() - Implementacion comun. restore=True conserva la posicion.
'==============================================================================
Function TracksJsonCore(rs, restore)
    Dim json, firstRow, id, name, path, artist, album, cover, plays
    Dim savedBookmark, hasBookmark

    json = "["
    firstRow = True
    hasBookmark = False

    If restore Then
        If Not rs.EOF And Not rs.BOF Then
            savedBookmark = rs.Bookmark
            hasBookmark = True
        End If
    End If

    If Not rs.EOF Then rs.MoveFirst

    Do While Not rs.EOF
        id     = CLng(rs("ID_SONG"))
        name   = CStr(rs("SONG_NAME") & "")
        path   = MediaUrl(rs("SONG_PATH"), "")
        artist = CStr(rs("ARTIST_NAME") & "")
        If Len(CStr(rs("ARTIST_LASTNAME") & "")) > 0 Then
            artist = artist & " " & CStr(rs("ARTIST_LASTNAME") & "")
        End If
        album  = CStr(rs("TITLE_ALBUM") & "")
        cover  = CoverUrl(rs("ALBUM_IMAGE_PATH"))
        plays  = CLng(rs("PLAYS"))

        If Not firstRow Then json = json & ","
        firstRow = False

        json = json & "{" & _
            """id"":" & id & "," & _
            """title"":""" & JsonEscape(name) & """," & _
            """src"":""" & JsonEscape(path) & """," & _
            """artist"":""" & JsonEscape(artist) & """," & _
            """album"":""" & JsonEscape(album) & """," & _
            """cover"":""" & JsonEscape(cover) & """," & _
            """plays"":" & plays & "}"

        rs.MoveNext
    Loop

    If restore Then
        If hasBookmark Then rs.Bookmark = savedBookmark
    End If

    TracksJsonCore = json & "]"
End Function

'==============================================================================
' RenderTrackHead() - Cabecera de la tabla de canciones.
'==============================================================================
Sub RenderTrackHead()
%>
<div class="pm-track-head" role="row">
    <span class="pm-track-num" aria-hidden="true">#</span>
    <span class="pm-track-main">Titulo</span>
    <span class="pm-track-artist">Artista</span>
    <span class="pm-track-album">Album</span>
    <span class="pm-track-plays">Reproducciones</span>
    <span class="pm-track-actions" aria-hidden="true"></span>
</div>
<%
End Sub

'==============================================================================
' RenderTrackRow() - Fila de cancion, clicable para reproducir.
'==============================================================================
Sub RenderTrackRow(rs, tracksJson, index)
    Dim songId, cover
    songId = CLng(rs("ID_SONG"))
    cover = CoverUrl(rs("ALBUM_IMAGE_PATH"))
%>
<div class="pm-track-row" role="row" data-play="<%= index %>" tabindex="0">
    <span class="pm-track-num">
        <span class="pm-track-num-text"><%= index + 1 %></span>
        <span class="pm-track-num-play"><% Response.Write Icon("play") %></span>
    </span>

    <span class="pm-track-main">
        <span class="pm-track-cover">
            <img src="<%= Ha(cover) %>" alt="" loading="lazy">
        </span>
        <span class="pm-track-text">
            <span class="pm-track-title"><%= H(rs("SONG_NAME")) %></span>
            <span class="pm-track-sub"><%= H(rs("ARTIST_NAME") & "") %><% If Len(CStr(rs("ARTIST_LASTNAME") & "")) > 0 Then %> <%= H(rs("ARTIST_LASTNAME")) %><% End If %></span>
        </span>
    </span>

    <span class="pm-track-artist">
        <% If Not IsNull(rs("ID_ARTIST")) And CLng(rs("ID_ARTIST")) > 0 Then %>
        <a href="artist.asp?id=<%= CLng(rs("ID_ARTIST")) %>"><%= H(rs("ARTIST_NAME") & "") %></a>
        <% Else %>-<% End If %>
    </span>

    <span class="pm-track-album">
        <% If Not IsNull(rs("ID_ALBUM")) And CLng(rs("ID_ALBUM")) > 0 Then %>
        <a href="album.asp?id=<%= CLng(rs("ID_ALBUM")) %>"><%= H(rs("TITLE_ALBUM") & "") %></a>
        <% Else %>-<% End If %>
    </span>

    <span class="pm-track-plays"><%= CLng(rs("PLAYS")) %></span>

    <span class="pm-track-actions">
        <% If IsAdmin() Then %>
        <a class="pm-icon-btn pm-icon-btn-xs" href="admin/songs.asp?edit=<%= songId %>"
           title="Editar cancion" aria-label="Editar cancion"><% Response.Write Icon("edit") %></a>
        <a class="pm-icon-btn pm-icon-btn-xs" href="playlist_add.asp?song=<%= songId %>"
           title="Agregar a una playlist" aria-label="Agregar a una playlist"><% Response.Write Icon("plus") %></a>
        <% End If %>
    </span>
</div>
<%
End Sub

'==============================================================================
' RenderAlbumCard() - Tarjeta de album.
'==============================================================================
Sub RenderAlbumCard(rs)
    Dim cover
    cover = CoverUrl(rs("ALBUM_IMAGE_PATH"))
%>
<a class="pm-media-card" href="album.asp?id=<%= CLng(rs("ID_ALBUM")) %>">
    <span class="pm-media-art">
        <img src="<%= Ha(cover) %>" alt="Portada de <%= Ha(rs("TITLE_ALBUM")) %>" loading="lazy">
        <span class="pm-media-play" aria-hidden="true"><% Response.Write Icon("play") %></span>
    </span>
    <span class="pm-media-body">
        <span class="pm-media-title"><%= H(rs("TITLE_ALBUM")) %></span>
        <span class="pm-media-sub">
            <%= IIf(IsNull(rs("RELEASE_YEAR")), "Sin ano", rs("RELEASE_YEAR")) %>
            &middot; <%= CLng(rs("SONG_COUNT")) %> <% If CLng(rs("SONG_COUNT")) = 1 Then %>cancion<% Else %>canciones<% End If %>
        </span>
    </span>
</a>
<%
End Sub

'==============================================================================
' RenderArtistCard() - Tarjeta de artista.
'==============================================================================
Sub RenderArtistCard(rs)
    Dim img, fullName
    img = MediaUrl(rs("ARTIST_IMAGE"), MEDIA_BASE_URL & "artistImages/default.jpg")
    fullName = ArtistNameFull(rs)
%>
<a class="pm-media-card pm-media-card-round" href="artist.asp?id=<%= CLng(rs("ID_ARTIST")) %>">
    <span class="pm-media-art pm-media-art-round">
        <img src="<%= Ha(img) %>" alt="<%= Ha(fullName) %>" loading="lazy">
    </span>
    <span class="pm-media-body">
        <span class="pm-media-title"><%= H(fullName) %></span>
        <span class="pm-media-sub">
            <%= CLng(rs("SONG_COUNT")) %> <% If CLng(rs("SONG_COUNT")) = 1 Then %>cancion<% Else %>canciones<% End If %>
            &middot; <%= CLng(rs("TOTAL_PLAYS")) %> reproducciones
        </span>
    </span>
</a>
<%
End Sub

'==============================================================================
' RenderPlaylistCard() - Tarjeta de playlist.
'==============================================================================
Sub RenderPlaylistCard(rs)
    Dim img, isPublic, ownerLabel
    img = MediaUrl(rs("PHOTO"), MEDIA_BASE_URL & "playlistImages/default.jpg")
    isPublic = (CLng(rs("TYPE")) = PLAYLIST_PUBLIC)

    If isPublic Then
        ownerLabel = "Publica"
    Else
        ownerLabel = "Privada"
    End If
%>
<a class="pm-media-card" href="playlist.asp?id=<%= CLng(rs("ID_PLAYLIST")) %>">
    <span class="pm-media-art">
        <img src="<%= Ha(img) %>" alt="Portada de <%= Ha(rs("TITLE")) %>" loading="lazy">
        <span class="pm-media-play" aria-hidden="true"><% Response.Write Icon("play") %></span>
        <span class="pm-media-badge">
            <% Response.Write Icon(IIf(isPublic, "globe", "lock")) %>
            <%= ownerLabel %>
        </span>
    </span>
    <span class="pm-media-body">
        <span class="pm-media-title"><%= H(rs("TITLE")) %></span>
        <span class="pm-media-sub">
            <%= CLng(rs("SONG_COUNT")) %> <% If CLng(rs("SONG_COUNT")) = 1 Then %>cancion<% Else %>canciones<% End If %>
            &middot; <%= H(rs("USERNAME") & "") %>
        </span>
    </span>
</a>
<%
End Sub

'==============================================================================
' RenderAppStart() - Apertura del armazon comun: sidebar, topbar y contenido.
'==============================================================================
Sub RenderAppStart(navKey, pageTitle, pageSubtitle, actionsHtml, needsLogin)
    Dim mustLogin
    mustLogin = False
    If IsEmpty(needsLogin) Then
        mustLogin = True
    Else
        mustLogin = CBool(needsLogin)
    End If

    If mustLogin Then RequireLogin()

    Response.Write "<!DOCTYPE html>" & vbCrLf
%>
<html lang="es">
<head>
<% RenderHead pageTitle %>
</head>
<body class="pm-app">
<a class="pm-skip" href="#pmMain">Ir al contenido principal</a>

<% RenderSidebar navKey %>

<div class="pm-main">
<% RenderTopbar pageTitle, pageSubtitle, actionsHtml %>
<main class="pm-content" id="pmMain">
<% RenderFlash %>
<%
End Sub

'==============================================================================
' RenderAppEnd() - Cierre del armazon: reproductor, footer y scripts.
'==============================================================================
Sub RenderAppEnd()
%>
</main>
<% RenderFooter %>
</div>

<% RenderPlayer %>
<div class="pm-sidebar-backdrop" data-pm-sidebar-toggle aria-hidden="true"></div>

<script src="assets/js/player.js"></script>
</body>
</html>
<%
End Sub

'==============================================================================
' RenderDeleteModal() - Marcado del modal de confirmacion reutilizable.
'==============================================================================
Sub RenderDeleteModal()
%>
<div class="pm-modal" id="pmConfirmModal" hidden aria-hidden="true" role="dialog" aria-modal="true">
    <div class="pm-modal-card">
        <span class="pm-modal-ico"><% Response.Write Icon("alert") %></span>
        <h3 class="pm-modal-title">Confirmar eliminacion</h3>
        <p class="pm-modal-text" id="pmConfirmText">Esta accion no se puede deshacer.</p>
        <div class="pm-modal-actions">
            <button class="pm-btn pm-btn-ghost" type="button" data-pm-modal-cancel>Cancelar</button>
            <button class="pm-btn pm-btn-danger" type="button" data-pm-modal-ok>Eliminar</button>
        </div>
    </div>
</div>
<%
End Sub

'==============================================================================
' GetExtension() - Extension de un nombre de archivo, en minusculas.
'==============================================================================
Function GetExtension(fileName)
    Dim s, i
    s = CStr(fileName & "")
    i = InStrRev(s, ".")
    If i <= 0 Then
        GetExtension = ""
    Else
        GetExtension = LCase(Mid(s, i + 1))
    End If
End Function

'==============================================================================
' GetBaseName() - Nombre de archivo sin extension.
'==============================================================================
Function GetBaseName(fileName)
    Dim s, i
    s = CStr(fileName & "")
    i = InStrRev(s, ".")
    If i <= 0 Then
        GetBaseName = s
    Else
        GetBaseName = Left(s, i - 1)
    End If
End Function

'==============================================================================
' SanitizeFileName() - Deja solo caracteres seguros para el sistema de archivos.
'==============================================================================
Function SanitizeFileName(fileName)
    Dim s, i, ch, out
    s = CStr(fileName & "")
    out = ""

    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or _
           (ch >= "0" And ch <= "9") Or ch = "." Or ch = "_" Or ch = "-" Then
            out = out & ch
        ElseIf ch = " " Then
            out = out & "_"
        End If
    Next

    If Len(out) = 0 Then out = "archivo"
    If Len(out) > 80 Then out = Left(out, 80)
    SanitizeFileName = out
End Function

'==============================================================================
' SaveUploadedFile() - Guarda un archivo subido en assets/media/<folderKey>.
' Devuelve la ruta con prefijo "~/" lista para guardar en la BD, o "" si no
' se subio nada o hubo un error.
'==============================================================================
Function SaveUploadedFile(fieldName, folderKey, allowedExts, dbPrefix)
    Dim fso, upload, fileName, ext, safeName, finalName, targetDir, targetPath
    Dim bytes, stream, uniqueName

    SaveUploadedFile = ""

    If Request.Files.Count = 0 Then Exit Function

    Set upload = Request.Files(fieldName)
    If upload Is Nothing Then Exit Function
    If Len(CStr(upload.FileName & "")) = 0 Then Exit Function
    If upload.Size = 0 Then Exit Function

    ' Limite de tamano
    If upload.Size > UPLOAD_MAX_BYTES Then
        Err.Raise vbObjectError + 2001, "SaveUploadedFile", _
            "El archivo supera el limite de " & (UPLOAD_MAX_BYTES \ 1048576) & " MB."
    End If

    fileName = CStr(upload.FileName)
    ' Algunos navegadores mandan la ruta completa: nos quedamos con el nombre
    fileName = Replace(fileName, "/", "\")
    Dim slashPos
    slashPos = InStrRev(fileName, "\")
    If slashPos > 0 Then fileName = Mid(fileName, slashPos + 1)

    ext = GetExtension(fileName)

    ' Validacion de extension contra lista blanca
    If InStr(allowedExts, "|" & ext & "|") = 0 Then
        Err.Raise vbObjectError + 2002, "SaveUploadedFile", _
            "El tipo de archivo ." & ext & " no esta permitido."
    End If

    ' Nombre unico: base + marca de tiempo para no sobrescribir
    safeName = SanitizeFileName(GetBaseName(fileName))
    If Len(safeName) > 50 Then safeName = Left(safeName, 50)
    uniqueName = safeName & "_" & Hex(Timer() * 1000) & "_" & CLng(Rnd() * 1000) & "." & ext

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    targetDir = Server.MapPath(MEDIA_BASE_URL & folderKey)

    If Not fso.FolderExists(targetDir) Then
        fso.CreateFolder targetDir
    End If

    targetPath = fso.BuildPath(targetDir, uniqueName)

    ' Guarda el binario con ADODB.Stream
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write upload
    stream.SaveToFile targetPath, 2
    stream.Close
    Set stream = Nothing
    Set fso = Nothing

    SaveUploadedFile = dbPrefix & uniqueName
End Function
%>
