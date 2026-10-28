<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
'==============================================================================
' PoliMusic - Detalle de album
'==============================================================================
Dim albumId, albumRs, songsRs, queueJson, trackIndex

albumId = ParseId(Request.QueryString("id"))
If albumId = 0 Then
    NotFound "No se indico que album quieres consultar."
End If

Set albumRs = AlbumById(albumId)
If albumRs.EOF Then
    albumRs.Close : Set albumRs = Nothing
    NotFound "El album que buscas no existe o fue eliminado."
End If

Set songsRs = SongsByAlbum(albumId)
queueJson = TracksJsonSafe(songsRs)

Dim cover, albumTitle, releaseYear, songCount
cover       = CoverUrl(albumRs("ALBUM_IMAGE_PATH"))
albumTitle  = CStr(albumRs("TITLE_ALBUM") & "")
releaseYear = albumRs("RELEASE_YEAR")
songCount   = CLng(albumRs("SONG_COUNT"))

Dim actionsHtml
actionsHtml = ""
If IsAdmin() Then
    actionsHtml = "<a class=""pm-btn pm-btn-ghost pm-btn-sm"" href=""admin/albums.asp?edit=" & albumId & """>" & _
                  Icon("edit") & "<span>Editar album</span></a>"
End If

Response.Write "<!DOCTYPE html>" & vbCrLf
%>
<html lang="es">
<head>
<% RenderHead albumTitle %>
</head>
<body class="pm-app">
<a class="pm-skip" href="#pmMain">Ir al contenido principal</a>

<% RenderSidebar "albums" %>

<div class="pm-main">
<% RenderTopbar "Album", albumTitle, actionsHtml %>

<main class="pm-content" id="pmMain">
<% RenderFlash %>

    <!-- ===== Cabecera del album ===== -->
    <section class="pm-detail">
        <div class="pm-detail-art">
            <img src="<%= Ha(cover) %>" alt="Portada de <%= Ha(albumTitle) %>">
        </div>

        <div class="pm-detail-body">
            <p class="pm-detail-kind">Album</p>
            <h2 class="pm-detail-title"><%= H(albumTitle) %></h2>

            <div class="pm-detail-meta">
                <span class="pm-chip">
                    <% Response.Write Icon("calendar") %>
                    <%= IIf(IsNull(releaseYear), "Sin ano", releaseYear) %>
                </span>
                <span class="pm-chip">
                    <% Response.Write Icon("music") %>
                    <%= songCount %> <% If songCount = 1 Then %>cancion<% Else %>canciones<% End If %>
                </span>
            </div>

            <div class="pm-detail-actions">
                <% If songCount > 0 Then %>
                <button class="pm-btn pm-btn-primary" type="button"
                        onclick="PoliMusicPlayer.load(window.PM_QUEUE, 0)">
                    <% Response.Write Icon("play") %><span>Reproducir album</span>
                </button>
                <% End If %>
                <a class="pm-btn pm-btn-ghost" href="albums.asp">
                    <% Response.Write Icon("arrow-left") %><span>Volver a albumes</span>
                </a>
            </div>
        </div>
    </section>

    <!-- ===== Listado de canciones ===== -->
    <section class="pm-section">
        <div class="pm-section-head">
            <div>
                <h2 class="pm-section-title">Canciones del album</h2>
                <p class="pm-section-sub">Haz clic en una fila para reproducirla</p>
            </div>
        </div>

        <% If songsRs.EOF Then %>
        <% RenderEmpty "music", "Este album no tiene canciones", _
                       "Todavia no se han asociado canciones a este album.", "", "" %>
        <% Else %>
        <div class="pm-tracks">
            <% RenderTrackHead %>
            <%
            trackIndex = 0
            songsRs.MoveFirst
            Do While Not songsRs.EOF
                RenderTrackRow songsRs, queueJson, trackIndex
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

<script>
window.PM_QUEUE = <% Response.Write queueJson %>;
</script>
<script src="assets/js/player.js"></script>
</body>
</html>
<%
albumRs.Close : Set albumRs = Nothing
songsRs.Close : Set songsRs = Nothing
%>
