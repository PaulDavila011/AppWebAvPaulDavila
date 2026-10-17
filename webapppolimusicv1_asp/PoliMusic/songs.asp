<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
' =====================================================================
'  Página de canciones (Song List)
'  - Ambos roles: ver y reproducir.  - Solo ADMIN: crear, editar, eliminar.
' =====================================================================
RequireLogin
pageTitle = "Song List"
activeNav = "songs"

Dim q, sortKey, orderBy, songs, s, i, errMsg, userName, hasFile, cover, artists

' --- Eliminación (CRUD: Delete). El rol se valida en el servidor, no solo ocultando botones.
If IsPost() Then
    If Not IsAdmin() Then RenderForbidden
    RequireCsrf
    If FormValue("action") = "delete" Then
        errMsg = DeleteSong(ToInt(FormValue("id"), 0))
        If errMsg = "" Then SetFlash "success", "Canción eliminada." Else SetFlash "error", errMsg
    End If
    RedirectTo "songs.asp"
End If

' Nombre leído de la variable de sesión guardada al iniciar sesión en login.asp
userName = Session("uname")

q = QS("q")
sortKey = QS("sort")
Select Case sortKey
    Case "plays" : orderBy = "ISNULL(s.PLAYS,0) DESC, s.SONG_NAME"
    Case "new"   : orderBy = "s.ID_SONG DESC"
    Case Else    : sortKey = "name" : orderBy = "s.SONG_NAME"
End Select

If q <> "" Then
    songs = QList("SELECT * FROM (" & SongSelectSql() & ") s WHERE s.SONG_NAME LIKE ? OR s.ARTISTS LIKE ? OR s.ALBUM_TITLE LIKE ? ORDER BY " & orderBy, Array("%" & q & "%", "%" & q & "%", "%" & q & "%"))
Else
    songs = QList(SongSelectSql() & " ORDER BY " & orderBy, Array())
End If

Function SortLink(key, label)
    SortLink = "<a class=""chip" & IIfStr(key = sortKey, " active", "") & """ href=""songs.asp?sort=" & key & IIfStr(q <> "", "&amp;q=" & Server.URLEncode(q), "") & """>" & label & "</a>"
End Function
%>
<!--#include file="includes/layout_top.asp"-->

<section class="welcome" aria-label="Bienvenida">
  <span class="avatar avatar-lg"><%= Img(Session("photo"), "user", "", "") %></span>
  <div class="welcome-text">
    <h2>¡Bienvenido, <span><%= H(userName) %></span>!</h2>
    <p>Sesión iniciada como <%= RoleBadge(Session("role")) %>
      <% If IsAdmin() Then %>· Puedes crear, editar y eliminar canciones.<% Else %>· Puedes escuchar todas las canciones del catálogo.<% End If %></p>
  </div>
  <form method="post" action="logout.asp" data-no-pjax>
    <%= CsrfField() %>
    <button type="submit" class="btn btn-ghost"><%= Icon("logout") %>Cerrar sesión</button>
  </form>
</section>

<header class="page-head">
  <div>
    <p class="eyebrow"><%= Icon("music") %> Biblioteca</p>
    <h1>Song List</h1>
    <p class="muted">A continuación se presenta el listado de canciones disponibles (<%= Plural(UBound(songs) + 1, "canción", "canciones") %><% If q <> "" Then %> para “<%= H(q) %>”<% End If %>).</p>
  </div>
  <div class="head-actions">
    <% If IsAdmin() Then %><a class="btn btn-ghost" href="admin_song_form.asp?r=songs"><%= Icon("plus") %>Nueva canción</a><% End If %>
    <% If UBound(songs) >= 0 Then %><button class="btn btn-primary btn-play-big" type="button" data-play-list="#song-table" aria-label="Reproducir todo"><%= Icon("play") %></button><% End If %>
  </div>
</header>

<div class="toolbar-row">
  <% RenderSearchBar "Filtrar por título, artista o álbum", q %>
  <div class="chips" role="group" aria-label="Ordenar"><%= SortLink("name", "A–Z") %><%= SortLink("plays", "Más escuchadas") %><%= SortLink("new", "Recientes") %></div>
</div>

<% If UBound(songs) < 0 Then
     RenderEmpty "music", "No hay canciones para mostrar."
   Else %>
<div class="table-wrap">
  <table class="table song-table" id="song-table" data-tracklist>
    <caption class="sr-only">Song List: listado de canciones</caption>
    <thead>
      <tr>
        <th class="col-idx" scope="col">#</th>
        <th scope="col">Canción</th>
        <th scope="col" class="hide-sm">Álbum</th>
        <th scope="col" class="num">Reproducciones</th>
        <% If IsAdmin() Then %><th scope="col" class="actions">Acciones</th><% End If %>
      </tr>
    </thead>
    <tbody>
    <% i = 0
       For Each s In songs
         i = i + 1
         cover = Nz(s("COVER"), "")
         artists = Nz(s("ARTISTS"), "")
         If artists = "" Then artists = "Artista desconocido"
         hasFile = MediaExists(s("SONG_PATH")) %>
      <tr data-track data-id="<%= s("ID_SONG") %>" data-src="<%= H(MediaUrl(s("SONG_PATH"))) %>" data-title="<%= H(s("SONG_NAME")) %>" data-artist="<%= H(artists) %>" data-cover="<%= H(MediaUrl(cover)) %>">
        <td class="col-idx">
          <span class="tl-idx">
            <span class="tl-n"><%= i %></span>
            <button type="button" class="tl-play" data-play-track aria-label="Reproducir <%= H(s("SONG_NAME")) %>"><%= Icon("play") %></button>
            <span class="eq" aria-hidden="true"><i></i><i></i><i></i></span>
          </span>
        </td>
        <td>
          <div class="cell-main">
            <%= Img(cover, "cover", "", "thumb") %>
            <div><strong class="song-title"><%= H(s("SONG_NAME")) %></strong><small><%= H(artists) %></small>
              <% If Not hasFile Then %><small class="no-audio"><%= Icon("alert") %>Audio no disponible</small><% End If %></div>
          </div>
        </td>
        <td class="hide-sm"><% If Nz(s("ID_ALBUM"), "") <> "" Then %><a href="album.asp?id=<%= s("ID_ALBUM") %>"><%= H(s("ALBUM_TITLE")) %></a><% Else %><span class="muted">Sencillo</span><% End If %></td>
        <td class="num tl-plays"><%= FmtNum(s("PLAYS")) %></td>
        <% If IsAdmin() Then %>
        <td class="actions">
          <%= EditButton("admin_song_form.asp?id=" & s("ID_SONG") & "&amp;r=songs", s("SONG_NAME")) %>
          <%= DeleteButton(s("ID_SONG"), "la canción «" & s("SONG_NAME") & "»") %>
        </td>
        <% End If %>
      </tr>
    <% Next %>
    </tbody>
  </table>
</div>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
