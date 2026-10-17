<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireAdmin
pageTitle = "Gestión de canciones"
activeNav = "admin_songs"

Dim q, songs, s, id, errMsg, dummy, hasFile, oldMedia
If IsPost() Then
    RequireCsrf
    If FormValue("action") = "delete" Then
        id = ToInt(FormValue("id"), 0)
        errMsg = DeleteSong(id)
        If errMsg = "" Then SetFlash "success", "Canción eliminada del catálogo." Else SetFlash "error", errMsg
    End If
    RedirectTo "admin_songs.asp"
End If

q = QS("q")
If q <> "" Then
    songs = QList("SELECT x.*, (SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_SONG = x.ID_SONG) AS ALBUMS, (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_SONG = x.ID_SONG) AS PLAYLISTS FROM (" & SongSelectSql() & ") x WHERE x.SONG_NAME LIKE ? OR x.ARTISTS LIKE ? ORDER BY x.SONG_NAME", Array("%" & q & "%", "%" & q & "%"))
Else
    songs = QList("SELECT x.*, (SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_SONG = x.ID_SONG) AS ALBUMS, (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_SONG = x.ID_SONG) AS PLAYLISTS FROM (" & SongSelectSql() & ") x ORDER BY x.SONG_NAME", Array())
End If
%>
<!--#include file="includes/layout_top.asp"-->
<% RenderAdminHeader "Canciones", "Sube archivos de audio y relaciona cada canción con sus artistas y álbumes.", "admin_song_form.asp", "Nueva canción" %>
<div class="toolbar-row"><% RenderSearchBar "Buscar por título o artista", q %></div>
<% If UBound(songs) < 0 Then
     RenderEmpty "music", "No hay canciones que coincidan."
   Else %>
<div class="table-wrap">
<table class="table">
  <thead><tr><th>Canción</th><th class="hide-md">Archivo</th><th class="num hide-sm">Álbumes</th><th class="num hide-sm">Playlists</th><th class="num">Reprod.</th><th class="actions"><span class="sr-only">Acciones</span></th></tr></thead>
  <tbody>
  <% For Each s In songs
       hasFile = MediaExists(s("SONG_PATH")) %>
    <tr>
      <td><div class="cell-main"><%= Img(s("COVER"), "cover", "", "thumb") %><div><strong><%= H(s("SONG_NAME")) %></strong><small><%= H(Nz(s("ARTISTS"), "Sin artista")) %></small></div></div></td>
      <td class="hide-md"><% If hasFile Then %><span class="badge badge-success"><%= Icon("check") %>Disponible</span><% Else %><span class="badge badge-warning" title="<%= H(s("SONG_PATH")) %>"><%= Icon("alert") %>Falta archivo</span><% End If %><small class="path"><%= H(s("SONG_PATH")) %></small></td>
      <td class="num hide-sm"><%= s("ALBUMS") %></td>
      <td class="num hide-sm"><%= s("PLAYLISTS") %></td>
      <td class="num"><%= FmtNum(s("PLAYS")) %></td>
      <td class="actions">
        <%= EditButton("admin_song_form.asp?id=" & s("ID_SONG"), s("SONG_NAME")) %>
        <%= DeleteButton(s("ID_SONG"), "la canción «" & s("SONG_NAME") & "»") %>
      </td>
    </tr>
  <% Next %>
  </tbody>
</table>
</div>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
