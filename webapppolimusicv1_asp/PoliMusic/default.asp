<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireLogin
pageTitle = "Inicio"
activeNav = "home"

Dim topSongs, albums, artists, playlists, p, a, ar, stats
topSongs = QList(Replace(SongSelectSql(), "SELECT s.ID_SONG", "SELECT TOP 6 s.ID_SONG") & " ORDER BY ISNULL(s.PLAYS,0) DESC, s.SONG_NAME", Array())
albums = QList("SELECT TOP 10 al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, (SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS SONGS FROM TBL_ALBUM al ORDER BY al.RELEASE_YEAR DESC, al.TITLE_ALBUM", Array())
artists = QList("SELECT TOP 10 a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, (SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS SONGS FROM TBL_ARTIST a ORDER BY SONGS DESC, a.ARTIST_NAME", Array())
If IsAdmin() Then
    playlists = QList("SELECT TOP 6 pl.ID_PLAYLIST, pl.TITLE, pl.PHOTO, pl.TYPE FROM TBL_PLAYLIST pl ORDER BY pl.CREATION_DATE DESC", Array())
    Set stats = QRow("SELECT (SELECT COUNT(*) FROM TBL_USER) AS USERS, (SELECT COUNT(*) FROM TBL_SONG) AS SONGS, (SELECT COUNT(*) FROM TBL_ALBUM) AS ALBUMS, (SELECT COUNT(*) FROM TBL_ARTIST) AS ARTISTS", Array())
Else
    playlists = QList("SELECT TOP 6 pl.ID_PLAYLIST, pl.TITLE, pl.PHOTO, pl.TYPE FROM TBL_PLAYLIST pl WHERE pl.TYPE = ? OR pl.ID_USER = ? ORDER BY pl.CREATION_DATE DESC", Array(PLAYLIST_PUBLIC, CurrentUserId()))
End If
%>
<!--#include file="includes/layout_top.asp"-->
<section class="hero">
  <div class="hero-text">
    <p class="eyebrow"><%= Icon("headphones") %> <%= FmtDate(Date()) %></p>
    <h1><%= Greeting() %>, <span class="grad"><%= H(Session("uname")) %></span></h1>
    <p class="muted">Esto es lo que suena hoy en PoliMusic.</p>
    <div class="hero-actions">
      <% If UBound(topSongs) >= 0 Then %><button class="btn btn-primary btn-lg" type="button" data-play-list="#top-songs"><%= Icon("play") %>Reproducir lo más escuchado</button><% End If %>
      <a class="btn btn-ghost btn-lg" href="albums.asp">Explorar álbumes</a>
    </div>
  </div>
  <div class="hero-art" aria-hidden="true"><div class="bars"><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i></div></div>
</section>

<% If IsAdmin() Then %>
<section class="admin-strip">
  <div class="admin-strip-text"><%= Icon("shield") %><div><strong>Panel de administración</strong><small><%= Plural(stats("USERS"), "usuario", "usuarios") %> · <%= Plural(stats("SONGS"), "canción", "canciones") %> · <%= Plural(stats("ALBUMS"), "álbum", "álbumes") %> · <%= Plural(stats("ARTISTS"), "artista", "artistas") %></small></div></div>
  <a class="btn btn-primary" href="admin.asp">Ir al panel</a>
</section>
<% End If %>

<% If UBound(playlists) >= 0 Then %>
<section class="section">
  <div class="section-head"><h2>Playlists</h2><a class="link-more" href="playlists.asp">Ver todas</a></div>
  <div class="quick-grid">
    <% For Each p In playlists %>
    <a class="quick" href="playlist.asp?id=<%= p("ID_PLAYLIST") %>">
      <%= Img(p("PHOTO"), "cover", "", "") %><span><%= H(p("TITLE")) %></span>
    </a>
    <% Next %>
  </div>
</section>
<% End If %>

<section class="section">
  <div class="section-head"><h2>Lo más escuchado</h2><a class="link-more" href="songs.asp">Todas las canciones</a></div>
  <div id="top-songs"><% RenderTrackList topSongs, True, "" %></div>
</section>

<section class="section">
  <div class="section-head"><h2>Álbumes</h2><a class="link-more" href="albums.asp">Ver todos</a></div>
  <div class="card-row">
    <% For Each a In albums
         RenderCard "album.asp?id=" & a("ID_ALBUM"), a("ALBUM_IMAGE_PATH"), "cover", a("TITLE_ALBUM"), Nz(a("RELEASE_YEAR"), "—") & " · " & Plural(a("SONGS"), "canción", "canciones")
       Next %>
  </div>
</section>

<section class="section">
  <div class="section-head"><h2>Artistas</h2><a class="link-more" href="artists.asp">Ver todos</a></div>
  <div class="card-row">
    <% For Each ar In artists
         RenderCard "artist.asp?id=" & ar("ID_ARTIST"), ar("ARTIST_IMAGE"), "artist", ArtistFullName(ar), "Artista · " & Plural(ar("SONGS"), "canción", "canciones")
       Next %>
  </div>
</section>
<!--#include file="includes/layout_bottom.asp"-->
