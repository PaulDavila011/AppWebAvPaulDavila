<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireLogin
activeNav = "playlists"

Dim id, pl, songs, canView
id = ToInt(QS("id"), 0)
Set pl = QRow("SELECT pl.ID_PLAYLIST, pl.ID_USER, pl.TITLE, pl.PHOTO, pl.TYPE, pl.CREATION_DATE, u.USERNAME, u.USER_PHOTO FROM TBL_PLAYLIST pl JOIN TBL_USER u ON u.ID_USER = pl.ID_USER WHERE pl.ID_PLAYLIST = ?", Array(id))
canView = False
If Not pl Is Nothing Then
    ' Una playlist privada solo la ve su dueño o un administrador
    canView = IsAdmin() Or CLng(pl("TYPE")) = PLAYLIST_PUBLIC Or CLng(pl("ID_USER")) = CurrentUserId()
End If
If Not canView Then
    Response.Status = "404 Not Found"
    pageTitle = "Playlist no disponible"
Else
    pageTitle = pl("TITLE")
    songs = QList(SongSelectSql() & " WHERE s.ID_SONG IN (SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?) ORDER BY s.SONG_NAME", Array(id))
End If
%>
<!--#include file="includes/layout_top.asp"-->
<% If Not canView Then %>
  <div class="empty empty-lg"><%= Icon("lock") %><h2>Playlist no disponible</h2><p>No existe o es privada.</p><a class="btn btn-primary" href="playlists.asp">Ver playlists</a></div>
<% Else %>
<section class="detail-hero">
  <%= Img(pl("PHOTO"), "cover", pl("TITLE"), "detail-cover") %>
  <div class="detail-info">
    <p class="eyebrow">Playlist <%= PlaylistTypeBadge(pl("TYPE")) %></p>
    <h1><%= H(pl("TITLE")) %></h1>
    <p class="detail-meta">
      <span class="artist-chip"><%= Img(pl("USER_PHOTO"), "user", "", "") %><%= H(pl("USERNAME")) %></span>
      <span>Creada el <%= FmtDate(pl("CREATION_DATE")) %></span><span><%= Plural(UBound(songs) + 1, "canción", "canciones") %></span>
    </p>
  </div>
</section>
<div class="action-bar">
  <% If UBound(songs) >= 0 Then %><button class="btn btn-primary btn-play-big" type="button" data-play-list="#playlist-tracks" aria-label="Reproducir playlist"><%= Icon("play") %></button><% End If %>
  <% If IsAdmin() Then %><a class="btn btn-ghost" href="admin_playlist_form.asp?id=<%= id %>"><%= Icon("edit") %>Editar playlist</a><% End If %>
</div>
<div id="playlist-tracks"><% RenderTrackList songs, True, "" %></div>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
