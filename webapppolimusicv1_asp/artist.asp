<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
'==============================================================================
' PoliMusic - Detalle de artista
'==============================================================================
Dim artistId, artistRs, songsRs, albumsRs, queueJson, trackIndex

artistId = ParseId(Request.QueryString("id"))
If artistId = 0 Then
    NotFound "No se indico que artista quieres consultar."
End If

Set artistRs = ArtistById(artistId)
If artistRs.EOF Then
    artistRs.Close : Set artistRs = Nothing
    NotFound "El artista que buscas no existe o fue eliminado."
End If

Set songsRs = SongsByArtist(artistId)
queueJson = TracksJsonSafe(songsRs)

' Albumes en los que participa el artista
Set albumsRs = ExecuteQuery( _
    "SELECT DISTINCT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " & _
    "  (SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS SONG_COUNT " & _
    "FROM TBL_ALBUM al " & _
    "INNER JOIN TBL_TRACKLIST t ON t.ID_ALBUM = al.ID_ALBUM " & _
    "INNER JOIN TBL_PLAYER p ON p.ID_SONG = t.ID_SONG " & _
    "WHERE p.ID_ARTIST = ? ORDER BY al.TITLE_ALBUM", Array(artistId))

Dim img, fullName, songCount, nArtistPlays
img        = MediaUrl(artistRs("ARTIST_IMAGE"), MEDIA_BASE_URL & "artistImages/default.jpg")
fullName   = ArtistNameFull(artistRs)
songCount  = CLng(artistRs("SONG_COUNT"))
nArtistPlays = CLng(artistRs("TOTAL_PLAYS"))

Dim actionsHtml
actionsHtml = ""
If IsAdmin() Then
    actionsHtml = "<a class=""pm-btn pm-btn-ghost pm-btn-sm"" href=""admin/artists.asp?edit=" & artistId & """>" & _
                  Icon("edit") & "<span>Editar artista</span></a>"
End If

Response.Write "<!DOCTYPE html>" & vbCrLf
%>
<html lang="es">
<head>
<% RenderHead fullName %>
</head>
<body class="pm-app">
<a class="pm-skip" href="#pmMain">Ir al contenido principal</a>

<% RenderSidebar "artists" %>

<div class="pm-main">
<% RenderTopbar "Artista", fullName, actionsHtml %>

<main class="pm-content" id="pmMain">
<% RenderFlash %>

    <!-- ===== Cabecera del artista ===== -->
    <section class="pm-detail">
        <div class="pm-detail-art pm-detail-art-round">
            <img src="<%= Ha(img) %>" alt="<%= Ha(fullName) %>">
        </div>

        <div class="pm-detail-body">
            <p class="pm-detail-kind">Artista</p>
            <h2 class="pm-detail-title"><%= H(fullName) %></h2>

            <div class="pm-detail-meta">
                <span class="pm-chip">
                    <% Response.Write Icon("music") %>
                    <%= songCount %> <% If songCount = 1 Then %>cancion<% Else %>canciones<% End If %>
                </span>
                <span class="pm-chip">
                    <% Response.Write Icon("fire") %>
                    <%= nArtistPlays %> reproducciones
                </span>
            </div>

            <div class="pm-detail-actions">
                <% If songCount > 0 Then %>
                <button class="pm-btn pm-btn-primary" type="button"
                        onclick="PoliMusicPlayer.load(window.PM_QUEUE, 0)">
                    <% Response.Write Icon("play") %><span>Reproducir todo</span>
                </button>
                <% End If %>
                <a class="pm-btn pm-btn-ghost" href="artists.asp">
                    <% Response.Write Icon("arrow-left") %><span>Volver a artistas</span>
                </a>
            </div>
        </div>
    </section>

    <!-- ===== Canciones ===== -->
    <section class="pm-section">
        <div class="pm-section-head">
            <div>
                <h2 class="pm-section-title">Canciones</h2>
                <p class="pm-section-sub">Todas las canciones interpretadas por este artista</p>
            </div>
        </div>

        <% If songsRs.EOF Then %>
        <% RenderEmpty "music", "Sin canciones registradas", _
                       "Este artista todavia no tiene canciones asociadas en el catalogo.", "", "" %>
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

    <!-- ===== Albumes ===== -->
    <% If Not albumsRs.EOF Then %>
    <section class="pm-section">
        <div class="pm-section-head">
            <div>
                <h2 class="pm-section-title">Aparece en</h2>
                <p class="pm-section-sub">Albumes que incluyen canciones de este artista</p>
            </div>
        </div>

        <div class="pm-grid">
            <%
            albumsRs.MoveFirst
            Do While Not albumsRs.EOF
                RenderAlbumCard albumsRs
                albumsRs.MoveNext
            Loop
            %>
        </div>
    </section>
    <% End If %>

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
artistRs.Close : Set artistRs = Nothing
songsRs.Close  : Set songsRs = Nothing
albumsRs.Close : Set albumsRs = Nothing
%>
