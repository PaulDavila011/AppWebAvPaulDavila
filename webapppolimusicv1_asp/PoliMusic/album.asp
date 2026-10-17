<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireLogin
activeNav = "albums"

Dim id, album, songs, artists, ar, totalPlays, s
id = ToInt(QS("id"), 0)
Set album = QRow("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(id))
If album Is Nothing Then
    Response.Status = "404 Not Found"
    pageTitle = "Álbum no encontrado"
Else
    pageTitle = album("TITLE_ALBUM")
    songs = QList(SongSelectSql() & " WHERE s.ID_SONG IN (SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = ?) ORDER BY s.ID_SONG", Array(id))
    artists = QList("SELECT DISTINCT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE FROM TBL_ARTIST a JOIN TBL_PLAYER p ON p.ID_ARTIST = a.ID_ARTIST JOIN TBL_TRACKLIST t ON t.ID_SONG = p.ID_SONG WHERE t.ID_ALBUM = ? ORDER BY a.ARTIST_NAME", Array(id))
    totalPlays = 0
    For Each s In songs
        totalPlays = totalPlays + CLng(s("PLAYS"))
    Next
End If
%>
<!--#include file="includes/layout_top.asp"-->
<% If album Is Nothing Then %>
  <div class="empty empty-lg"><%= Icon("disc") %><h2>Álbum no encontrado</h2><p>Puede que haya sido eliminado.</p><a class="btn btn-primary" href="albums.asp">Ver álbumes</a></div>
<% Else %>
<section class="detail-hero">
  <%= Img(album("ALBUM_IMAGE_PATH"), "cover", album("TITLE_ALBUM"), "detail-cover") %>
  <div class="detail-info">
    <p class="eyebrow">Álbum</p>
    <h1><%= H(album("TITLE_ALBUM")) %></h1>
    <p class="detail-meta">
      <% For Each ar In artists %><a href="artist.asp?id=<%= ar("ID_ARTIST") %>" class="artist-chip"><%= Img(ar("ARTIST_IMAGE"), "artist", "", "") %><%= H(ArtistFullName(ar)) %></a><% Next %>
      <span><%= Nz(album("RELEASE_YEAR"), "—") %></span><span><%= Plural(UBound(songs) + 1, "canción", "canciones") %></span><span><%= Plural(totalPlays, "reproducción", "reproducciones") %></span>
    </p>
  </div>
</section>
<div class="action-bar">
  <% If UBound(songs) >= 0 Then %><button class="btn btn-primary btn-play-big" type="button" data-play-list="#album-tracks" aria-label="Reproducir álbum"><%= Icon("play") %></button><% End If %>
  <% If IsAdmin() Then %><a class="btn btn-ghost" href="admin_album_form.asp?id=<%= id %>"><%= Icon("edit") %>Editar álbum</a><% End If %>
</div>
<div id="album-tracks"><% RenderTrackList songs, False, Nz(album("ALBUM_IMAGE_PATH"), "") %></div>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
