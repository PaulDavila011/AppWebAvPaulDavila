<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireLogin
activeNav = "artists"

Dim id, artist, songs, albums, a, totalPlays, s
id = ToInt(QS("id"), 0)
Set artist = QRow("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(id))
If artist Is Nothing Then
    Response.Status = "404 Not Found"
    pageTitle = "Artista no encontrado"
Else
    pageTitle = ArtistFullName(artist)
    songs = QList(SongSelectSql() & " WHERE s.ID_SONG IN (SELECT ID_SONG FROM TBL_PLAYER WHERE ID_ARTIST = ?) ORDER BY ISNULL(s.PLAYS,0) DESC, s.SONG_NAME", Array(id))
    albums = QList("SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH FROM TBL_ALBUM al WHERE al.ID_ALBUM IN (SELECT t.ID_ALBUM FROM TBL_TRACKLIST t JOIN TBL_PLAYER p ON p.ID_SONG = t.ID_SONG WHERE p.ID_ARTIST = ?) ORDER BY al.RELEASE_YEAR DESC", Array(id))
    totalPlays = 0
    For Each s In songs
        totalPlays = totalPlays + CLng(s("PLAYS"))
    Next
End If
%>
<!--#include file="includes/layout_top.asp"-->
<% If artist Is Nothing Then %>
  <div class="empty empty-lg"><%= Icon("mic") %><h2>Artista no encontrado</h2><p>Puede que haya sido eliminado.</p><a class="btn btn-primary" href="artists.asp">Ver artistas</a></div>
<% Else %>
<section class="detail-hero detail-hero-artist">
  <%= Img(artist("ARTIST_IMAGE"), "artist", ArtistFullName(artist), "detail-cover round") %>
  <div class="detail-info">
    <p class="eyebrow"><%= Icon("check") %> Artista</p>
    <h1><%= H(ArtistFullName(artist)) %></h1>
    <p class="detail-meta"><span><%= Plural(UBound(songs) + 1, "canción", "canciones") %></span><span><%= Plural(UBound(albums) + 1, "álbum", "álbumes") %></span><span><%= Plural(totalPlays, "reproducción", "reproducciones") %></span></p>
  </div>
</section>
<div class="action-bar">
  <% If UBound(songs) >= 0 Then %><button class="btn btn-primary btn-play-big" type="button" data-play-list="#artist-tracks" aria-label="Reproducir artista"><%= Icon("play") %></button><% End If %>
  <% If IsAdmin() Then %><a class="btn btn-ghost" href="admin_artist_form.asp?id=<%= id %>"><%= Icon("edit") %>Editar artista</a><% End If %>
</div>
<section class="section">
  <div class="section-head"><h2>Populares</h2></div>
  <div id="artist-tracks"><% RenderTrackList songs, True, "" %></div>
</section>
<% If UBound(albums) >= 0 Then %>
<section class="section">
  <div class="section-head"><h2>Aparece en</h2></div>
  <div class="card-row">
    <% For Each a In albums
         RenderCard "album.asp?id=" & a("ID_ALBUM"), a("ALBUM_IMAGE_PATH"), "cover", a("TITLE_ALBUM"), Nz(a("RELEASE_YEAR"), "—") & " · Álbum"
       Next %>
  </div>
</section>
<% End If %>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
