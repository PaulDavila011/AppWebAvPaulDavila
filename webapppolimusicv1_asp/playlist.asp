<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
'==============================================================================
' PoliMusic - Detalle de playlist
'==============================================================================
' Reglas de visibilidad:
'   * Playlist publica            -> visible para todos.
'   * Playlist privada            -> visible para su dueño y para el admin.
'==============================================================================
Dim playlistId, plRs, songsRs, queueJson, trackIndex
Dim isOwner, canSee, isPublic

playlistId = ParseId(Request.QueryString("id"))
If playlistId = 0 Then
    NotFound "No se indico que playlist quieres abrir."
End If

Set plRs = PlaylistById(playlistId)
If plRs.EOF Then
    plRs.Close : Set plRs = Nothing
    NotFound "La playlist que buscas no existe o fue eliminada."
End If

isPublic = (CLng(plRs("TYPE")) = PLAYLIST_PUBLIC)
isOwner  = (CLng(plRs("ID_USER")) = CurrentUserId())
canSee   = isPublic Or isOwner Or IsAdmin()

If Not canSee Then
    plRs.Close : Set plRs = Nothing
    DenyAccess "Esta playlist es privada. Solo su dueño puede verla."
End If

Set songsRs = SongsByPlaylist(playlistId)
queueJson = TracksJsonSafe(songsRs)

Dim cover, plTitle, ownerName, songCount, createdDate
cover       = MediaUrl(plRs("PHOTO"), MEDIA_BASE_URL & "playlistImages/default.jpg")
plTitle     = CStr(plRs("TITLE") & "")
ownerName   = CStr(plRs("USERNAME") & "")
songCount   = CLng(plRs("SONG_COUNT"))
createdDate = plRs("CREATION_DATE")

Dim actionsHtml, canManage
canManage = isOwner Or IsAdmin()
actionsHtml = ""
If canManage Then
    actionsHtml = "<a class=""pm-btn pm-btn-ghost pm-btn-sm"" href=""playlist_edit.asp?id=" & playlistId & """>" & _
                  Icon("edit") & "<span>Editar</span></a>"
End If

Response.Write "<!DOCTYPE html>" & vbCrLf
%>
<html lang="es">
<head>
<% RenderHead plTitle %>
</head>
<body class="pm-app">
<a class="pm-skip" href="#pmMain">Ir al contenido principal</a>

<% RenderSidebar "playlists" %>

<div class="pm-main">
<% RenderTopbar "Playlist", plTitle, actionsHtml %>

<main class="pm-content" id="pmMain">
<% RenderFlash %>

    <!-- ===== Cabecera de la playlist ===== -->
    <section class="pm-detail">
        <div class="pm-detail-art">
            <img src="<%= Ha(cover) %>" alt="Portada de <%= Ha(plTitle) %>">
        </div>

        <div class="pm-detail-body">
            <p class="pm-detail-kind">
                Playlist <% If isPublic Then %>publica<% Else %>privada<% End If %>
            </p>
            <h2 class="pm-detail-title"><%= H(plTitle) %></h2>

            <div class="pm-detail-meta">
                <span class="pm-chip">
                    <% Response.Write Icon(IIf(isPublic, "globe", "lock")) %>
                    <%= IIf(isPublic, "Publica", "Privada") %>
                </span>
                <span class="pm-chip">
                    <% Response.Write Icon("music") %>
                    <%= songCount %> <% If songCount = 1 Then %>cancion<% Else %>canciones<% End If %>
                </span>
                <span class="pm-chip">
                    <% Response.Write Icon("user") %>
                    <%= H(ownerName) %>
                </span>
                <span class="pm-chip">
                    <% Response.Write Icon("calendar") %>
                    <%= FmtDate(createdDate) %>
                </span>
            </div>

            <div class="pm-detail-actions">
                <% If songCount > 0 Then %>
                <button class="pm-btn pm-btn-primary" type="button"
                        onclick="PoliMusicPlayer.load(window.PM_QUEUE, 0)">
                    <% Response.Write Icon("play") %><span>Reproducir playlist</span>
                </button>
                <% End If %>
                <% If canManage Then %>
                <a class="pm-btn pm-btn-ghost" href="playlist_add.asp?id=<%= playlistId %>">
                    <% Response.Write Icon("plus") %><span>Agregar canciones</span>
                </a>
                <% End If %>
                <a class="pm-btn pm-btn-ghost" href="playlists.asp">
                    <% Response.Write Icon("arrow-left") %><span>Volver</span>
                </a>
            </div>
        </div>
    </section>

    <!-- ===== Canciones ===== -->
    <section class="pm-section">
        <div class="pm-section-head">
            <div>
                <h2 class="pm-section-title">Canciones de la playlist</h2>
                <p class="pm-section-sub">Haz clic en una fila para reproducirla</p>
            </div>
        </div>

        <% If songsRs.EOF Then %>
        <% RenderEmpty "music", "Esta playlist esta vacia", _
                       "Agrega canciones del catalogo para empezar a escucharla.", _
                       IIf(canManage, "playlist_add.asp?id=" & playlistId, ""), _
                       IIf(canManage, "Agregar canciones", "") %>
        <% Else %>
        <div class="pm-tracks">
            <% RenderTrackHead %>
            <%
            trackIndex = 0
            songsRs.MoveFirst
            Do While Not songsRs.EOF
            %>
            <div class="pm-track-row" role="row" data-play="<%= trackIndex %>" tabindex="0">
                <span class="pm-track-num">
                    <span class="pm-track-num-text"><%= trackIndex + 1 %></span>
                    <span class="pm-track-num-play"><% Response.Write Icon("play") %></span>
                </span>

                <span class="pm-track-main">
                    <span class="pm-track-cover">
                        <img src="<%= Ha(CoverUrl(songsRs("ALBUM_IMAGE_PATH"))) %>" alt="" loading="lazy">
                    </span>
                    <span class="pm-track-text">
                        <span class="pm-track-title"><%= H(songsRs("SONG_NAME")) %></span>
                        <span class="pm-track-sub"><%= H(songsRs("ARTIST_NAME") & "") %></span>
                    </span>
                </span>

                <span class="pm-track-artist">
                    <% If Not IsNull(songsRs("ID_ARTIST")) And CLng(songsRs("ID_ARTIST")) > 0 Then %>
                    <a href="artist.asp?id=<%= CLng(songsRs("ID_ARTIST")) %>"><%= H(songsRs("ARTIST_NAME") & "") %></a>
                    <% Else %>-<% End If %>
                </span>

                <span class="pm-track-album">
                    <% If Not IsNull(songsRs("ID_ALBUM")) And CLng(songsRs("ID_ALBUM")) > 0 Then %>
                    <a href="album.asp?id=<%= CLng(songsRs("ID_ALBUM")) %>"><%= H(songsRs("TITLE_ALBUM") & "") %></a>
                    <% Else %>-<% End If %>
                </span>

                <span class="pm-track-plays"><%= CLng(songsRs("PLAYS")) %></span>

                <span class="pm-track-actions">
                    <% If canManage Then %>
                    <form method="post" action="playlist_song.asp" class="pm-inline-form"
                          onsubmit="return pmConfirmDelete(event, '<%= JsArg(songsRs("SONG_NAME")) %>')">
                        <% Response.Write CsrfField() %>
                        <input type="hidden" name="playlist" value="<%= playlistId %>">
                        <input type="hidden" name="song" value="<%= CLng(songsRs("ID_SONG")) %>">
                        <input type="hidden" name="op" value="remove">
                        <button class="pm-btn pm-btn-danger pm-btn-sm" type="submit"
                                title="Quitar de la playlist">
                            <% Response.Write Icon("close") %><span>Quitar</span>
                        </button>
                    </form>
                    <% End If %>
                </span>
            </div>
            <%
                trackIndex = trackIndex + 1
                songsRs.MoveNext
            Loop
            %>
        </div>
        <% End If %>
    </section>

</main>
<% RenderFooter %>
</div>

<% RenderPlayer %>
<div class="pm-sidebar-backdrop" data-pm-sidebar-toggle aria-hidden="true"></div>
<% RenderDeleteModal %>

<script>
window.PM_QUEUE = <% Response.Write queueJson %>;
</script>
<script src="assets/js/player.js"></script>
</body>
</html>
<%
plRs.Close    : Set plRs = Nothing
songsRs.Close : Set songsRs = Nothing
%>
