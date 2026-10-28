<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
'==============================================================================
' PoliMusic - Pantalla principal
'==============================================================================
' Portada publica. Muestra el catalogo destacado, las playlists visibles para
' el usuario y las estadisticas del sistema.
'==============================================================================
Dim topSongs, recentAlbums, allArtists, visiblePlaylists
Dim nSongs, nAlbums, nArtists, nPlaylists, nPlays
Dim queueJson, trackIndex

Set topSongs = SongsTop(5)
queueJson = TracksJsonSafe(topSongs)

Set recentAlbums = AlbumsAll()
Set allArtists = ArtistsAll()
Set visiblePlaylists = PlaylistsVisible(CurrentUserId())

' Nota: los nombres locales no pueden coincidir con los de las funciones del
' repositorio (VBScript no distingue mayusculas): TotalPlays() ya existe, por
' eso aqui se usa nPlays.
nSongs     = CountRows("TBL_SONG", "")
nAlbums    = CountRows("TBL_ALBUM", "")
nArtists   = CountRows("TBL_ARTIST", "")
nPlaylists = CountRows("TBL_PLAYLIST", "")
nPlays     = TotalPlays()

Dim actionsHtml
actionsHtml = ""
If IsAdmin() Then
    actionsHtml = "<a class=""pm-btn pm-btn-primary pm-btn-sm"" href=""admin/default.asp"">" & _
                  Icon("dashboard") & "<span>Panel de administracion</span></a>"
End If

'=====================================================================
' Cabecera del documento
'=====================================================================
Response.Write "<!DOCTYPE html>" & vbCrLf
%>
<html lang="es">
<head>
<% RenderHead "Inicio" %>
</head>
<body class="pm-app">
<a class="pm-skip" href="#pmMain">Ir al contenido principal</a>

<% RenderSidebar "home" %>

<div class="pm-main">
<% RenderTopbar "Buenos dias, " & CurrentUserName(), "Aqui esta tu musica de hoy", actionsHtml %>

<main class="pm-content" id="pmMain">
<% RenderFlash %>

    <!-- ===== Portada ===== -->
    <section class="pm-hero">
        <span class="pm-hero-eyebrow">
            <% Response.Write Icon("sparkle") %>
            <span>Biblioteca de la Escuela Politecnica Nacional</span>
        </span>

        <h2 class="pm-hero-title">Toda la musica de la EPN, en un solo lugar.</h2>

        <p class="pm-hero-text">
            Explora el catalogo completo de canciones, albumes y artistas.
            Crea tus propias playlists y reproduce sin interrupciones mientras
            navegas por la aplicacion.
        </p>

        <div class="pm-hero-actions">
            <% If nSongs > 0 Then %>
            <button class="pm-btn pm-btn-primary pm-btn-lg" type="button"
                    onclick="PoliMusicPlayer.load(window.PM_QUEUE, 0)">
                <% Response.Write Icon("play") %><span>Reproducir todo</span>
            </button>
            <% End If %>
            <a class="pm-btn pm-btn-ghost pm-btn-lg" href="songs.asp">
                <% Response.Write Icon("music") %><span>Ver catalogo</span>
            </a>
        </div>

        <div class="pm-hero-stats">
            <div>
                <div class="pm-hero-stat-value"><%= nSongs %></div>
                <div class="pm-hero-stat-label">Canciones</div>
            </div>
            <div>
                <div class="pm-hero-stat-value"><%= nAlbums %></div>
                <div class="pm-hero-stat-label">Albumes</div>
            </div>
            <div>
                <div class="pm-hero-stat-value"><%= nArtists %></div>
                <div class="pm-hero-stat-label">Artistas</div>
            </div>
            <div>
                <div class="pm-hero-stat-value"><%= nPlaylists %></div>
                <div class="pm-hero-stat-label">Playlists</div>
            </div>
            <div>
                <div class="pm-hero-stat-value"><%= nPlays %></div>
                <div class="pm-hero-stat-label">Reproducciones</div>
            </div>
        </div>
    </section>

    <!-- ===== Canciones mas escuchadas ===== -->
    <section class="pm-section">
        <div class="pm-section-head">
            <div>
                <h2 class="pm-section-title">
                    <% Response.Write Icon("fire") %> Lo mas escuchado
                </h2>
                <p class="pm-section-sub">Las canciones con mas reproducciones del catalogo</p>
            </div>
            <a class="pm-section-link" href="songs.asp">
                Ver todas <% Response.Write Icon("chevron-right") %>
            </a>
        </div>

        <% If topSongs.EOF Then %>
        <% RenderEmpty "music", "Todavia no hay canciones", _
                       "Cuando se agreguen canciones al catalogo apareceran aqui.", _
                       "", "" %>
        <% Else %>
        <div class="pm-tracks">
            <% RenderTrackHead %>
            <%
            trackIndex = 0
            topSongs.MoveFirst
            Do While Not topSongs.EOF
                RenderTrackRow topSongs, queueJson, trackIndex
                trackIndex = trackIndex + 1
                topSongs.MoveNext
            Loop
            %>
        </div>
        <% End If %>
    </section>

    <!-- ===== Albumes ===== -->
    <section class="pm-section">
        <div class="pm-section-head">
            <div>
                <h2 class="pm-section-title"><% Response.Write Icon("album") %> Albumes</h2>
                <p class="pm-section-sub">Colecciones completas del catalogo</p>
            </div>
            <a class="pm-section-link" href="albums.asp">
                Ver todos <% Response.Write Icon("chevron-right") %>
            </a>
        </div>

        <% If recentAlbums.EOF Then %>
        <% RenderEmpty "album", "No hay albumes registrados", _
                       "El catalogo de albumes esta vacio por ahora.", "", "" %>
        <% Else %>
        <div class="pm-grid">
            <%
            Dim albumShown
            albumShown = 0
            recentAlbums.MoveFirst
            Do While Not recentAlbums.EOF And albumShown < 5
                RenderAlbumCard recentAlbums
                albumShown = albumShown + 1
                recentAlbums.MoveNext
            Loop
            %>
        </div>
        <% End If %>
    </section>

    <!-- ===== Artistas ===== -->
    <section class="pm-section">
        <div class="pm-section-head">
            <div>
                <h2 class="pm-section-title"><% Response.Write Icon("artist") %> Artistas</h2>
                <p class="pm-section-sub">Quienes interpretan la musica del catalogo</p>
            </div>
            <a class="pm-section-link" href="artists.asp">
                Ver todos <% Response.Write Icon("chevron-right") %>
            </a>
        </div>

        <% If allArtists.EOF Then %>
        <% RenderEmpty "artist", "No hay artistas registrados", _
                       "El catalogo de artistas esta vacio por ahora.", "", "" %>
        <% Else %>
        <div class="pm-grid">
            <%
            Dim artistShown
            artistShown = 0
            allArtists.MoveFirst
            Do While Not allArtists.EOF And artistShown < 5
                RenderArtistCard allArtists
                artistShown = artistShown + 1
                allArtists.MoveNext
            Loop
            %>
        </div>
        <% End If %>
    </section>

    <!-- ===== Playlists ===== -->
    <section class="pm-section">
        <div class="pm-section-head">
            <div>
                <h2 class="pm-section-title"><% Response.Write Icon("playlist") %> Playlists</h2>
                <p class="pm-section-sub">Listas publicas y las tuyas propias</p>
            </div>
            <a class="pm-section-link" href="playlists.asp">
                Ver todas <% Response.Write Icon("chevron-right") %>
            </a>
        </div>

        <% If visiblePlaylists.EOF Then %>
        <% RenderEmpty "playlist", "Todavia no hay playlists", _
                       "Crea tu primera playlist para organizar tus canciones favoritas.", _
                       "playlist_edit.asp", "Crear una playlist" %>
        <% Else %>
        <div class="pm-grid">
            <%
            Dim plShown
            plShown = 0
            visiblePlaylists.MoveFirst
            Do While Not visiblePlaylists.EOF And plShown < 5
                RenderPlaylistCard visiblePlaylists
                plShown = plShown + 1
                visiblePlaylists.MoveNext
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
topSongs.Close     : Set topSongs = Nothing
recentAlbums.Close : Set recentAlbums = Nothing
allArtists.Close   : Set allArtists = Nothing
visiblePlaylists.Close : Set visiblePlaylists = Nothing
%>
