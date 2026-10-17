<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireAdmin
pageTitle = "Panel de administración"
activeNav = "admin"

Dim st, topSongs, users, u, s, maxPlays, missing, allSongs
Set st = QRow("SELECT (SELECT COUNT(*) FROM TBL_USER) AS USERS, (SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = 1) AS ADMINS, " & _
              "(SELECT COUNT(*) FROM TBL_SONG) AS SONGS, (SELECT COUNT(*) FROM TBL_ALBUM) AS ALBUMS, (SELECT COUNT(*) FROM TBL_ARTIST) AS ARTISTS, " & _
              "(SELECT COUNT(*) FROM TBL_PLAYLIST) AS PLAYLISTS, (SELECT ISNULL(SUM(PLAYS),0) FROM TBL_SONG) AS PLAYS", Array())
topSongs = QList(Replace(SongSelectSql(), "SELECT s.ID_SONG", "SELECT TOP 5 s.ID_SONG") & " ORDER BY ISNULL(s.PLAYS,0) DESC", Array())
users = QList("SELECT TOP 6 ID_USER, USERNAME, EMAIL, USER_TYPE, USER_PHOTO FROM TBL_USER ORDER BY ID_USER DESC", Array())
maxPlays = 1
For Each s In topSongs
    If CLng(s("PLAYS")) > maxPlays Then maxPlays = CLng(s("PLAYS"))
Next
' Canciones cuyo archivo de audio no está en el servidor
missing = 0
allSongs = QList("SELECT SONG_PATH FROM TBL_SONG", Array())
For Each s In allSongs
    If Not MediaExists(s("SONG_PATH")) Then missing = missing + 1
Next

Sub StatCard(href, iconName, label, value, note)
    Response.Write "<a class=""stat"" href=""" & href & """><span class=""stat-ico"">" & Icon(iconName) & "</span><span class=""stat-body""><small>" & label & "</small><strong>" & FmtNum(value) & "</strong>"
    If note <> "" Then Response.Write "<em>" & note & "</em>"
    Response.Write "</span></a>"
End Sub
%>
<!--#include file="includes/layout_top.asp"-->
<% RenderAdminHeader "Panel de administración", "Resumen del catálogo y accesos rápidos de gestión.", "", "" %>

<div class="stat-grid">
  <% StatCard "admin_users.asp", "users", "Usuarios", st("USERS"), Plural(st("ADMINS"), "administrador", "administradores") %>
  <% StatCard "admin_songs.asp", "music", "Canciones", st("SONGS"), IIfStr(missing > 0, Plural(missing, "sin archivo de audio", "sin archivo de audio"), "Todas con audio") %>
  <% StatCard "admin_albums.asp", "disc", "Álbumes", st("ALBUMS"), "" %>
  <% StatCard "admin_artists.asp", "mic", "Artistas", st("ARTISTS"), "" %>
  <% StatCard "admin_playlists.asp", "list", "Playlists", st("PLAYLISTS"), "" %>
  <% StatCard "songs.asp?sort=plays", "headphones", "Reproducciones", st("PLAYS"), "Total acumulado" %>
</div>

<% If missing > 0 Then %>
<div class="alert alert-warning"><%= Icon("alert") %><span><strong><%= Plural(missing, "canción no tiene", "canciones no tienen") %></strong> su archivo de audio en el servidor. Súbelo desde <a href="admin_songs.asp">Gestión de canciones</a> o copia los .mp3 a la carpeta <code>songFiles</code>.</span></div>
<% End If %>

<div class="grid-2">
  <section class="panel">
    <div class="panel-head"><h2>Top canciones</h2><a class="link-more" href="admin_songs.asp">Gestionar</a></div>
    <% If UBound(topSongs) < 0 Then RenderEmpty "music", "Sin canciones." %>
    <ol class="bar-list">
      <% For Each s In topSongs %>
      <li>
        <div class="bar-label"><strong><%= H(s("SONG_NAME")) %></strong><small><%= H(Nz(s("ARTISTS"), "—")) %></small></div>
        <div class="bar-track"><span style="width:<%= Int(CLng(s("PLAYS")) * 100 / maxPlays) %>%"></span></div>
        <span class="bar-value"><%= FmtNum(s("PLAYS")) %></span>
      </li>
      <% Next %>
    </ol>
  </section>
  <section class="panel">
    <div class="panel-head"><h2>Usuarios recientes</h2><a class="link-more" href="admin_users.asp">Gestionar</a></div>
    <ul class="mini-list">
      <% For Each u In users %>
      <li>
        <span class="avatar"><%= Img(u("USER_PHOTO"), "user", "", "") %></span>
        <div><strong><%= H(u("USERNAME")) %></strong><small><%= H(u("EMAIL")) %></small></div>
        <%= RoleBadge(u("USER_TYPE")) %>
      </li>
      <% Next %>
    </ul>
  </section>
</div>

<section class="section">
  <div class="section-head"><h2>Accesos rápidos</h2></div>
  <div class="quick-actions">
    <a class="qa" href="admin_user_form.asp"><%= Icon("users") %><span>Nuevo usuario</span></a>
    <a class="qa" href="admin_song_form.asp"><%= Icon("upload") %><span>Subir canción</span></a>
    <a class="qa" href="admin_artist_form.asp"><%= Icon("mic") %><span>Nuevo artista</span></a>
    <a class="qa" href="admin_album_form.asp"><%= Icon("disc") %><span>Nuevo álbum</span></a>
    <a class="qa" href="admin_playlist_form.asp"><%= Icon("list") %><span>Nueva playlist</span></a>
  </div>
</section>
<!--#include file="includes/layout_bottom.asp"-->
