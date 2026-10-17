<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireAdmin
pageTitle = "Gestión de playlists"
activeNav = "admin_playlists"

Dim q, playlists, p, id, errMsg, dummy, oldMedia
If IsPost() Then
    RequireCsrf
    If FormValue("action") = "delete" Then
        id = ToInt(FormValue("id"), 0)
        oldMedia = QScalar("SELECT PHOTO FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(id))
        errMsg = TryBatch("SET @out = ?; DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = @out; DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = @out;", Array(id), dummy)
        If errMsg = "" Then DeleteUploadedMedia oldMedia
        If errMsg = "" Then SetFlash "success", "Playlist eliminada." Else SetFlash "error", errMsg
    End If
    RedirectTo "admin_playlists.asp"
End If

q = QS("q")
playlists = QList("SELECT pl.ID_PLAYLIST, pl.TITLE, pl.PHOTO, pl.TYPE, pl.CREATION_DATE, u.USERNAME, (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = pl.ID_PLAYLIST) AS SONGS " & _
                  "FROM TBL_PLAYLIST pl JOIN TBL_USER u ON u.ID_USER = pl.ID_USER WHERE (? = '' OR pl.TITLE LIKE ? OR u.USERNAME LIKE ?) ORDER BY pl.CREATION_DATE DESC", _
                  Array(q, "%" & q & "%", "%" & q & "%"))
%>
<!--#include file="includes/layout_top.asp"-->
<% RenderAdminHeader "Playlists", "Crea playlists, asígnalas a un usuario y define su visibilidad.", "admin_playlist_form.asp", "Nueva playlist" %>
<div class="toolbar-row"><% RenderSearchBar "Buscar por título o propietario", q %></div>
<% If UBound(playlists) < 0 Then
     RenderEmpty "list", "No hay playlists que coincidan."
   Else %>
<div class="table-wrap">
<table class="table">
  <thead><tr><th>Playlist</th><th class="hide-sm">Propietario</th><th>Visibilidad</th><th class="hide-md">Creación</th><th class="num">Canciones</th><th class="actions"><span class="sr-only">Acciones</span></th></tr></thead>
  <tbody>
  <% For Each p In playlists %>
    <tr>
      <td><div class="cell-main"><%= Img(p("PHOTO"), "cover", "", "thumb") %><div><strong><a href="playlist.asp?id=<%= p("ID_PLAYLIST") %>"><%= H(p("TITLE")) %></a></strong><small class="show-sm">de <%= H(p("USERNAME")) %></small></div></div></td>
      <td class="hide-sm"><%= H(p("USERNAME")) %></td>
      <td><%= PlaylistTypeBadge(p("TYPE")) %></td>
      <td class="hide-md"><%= FmtDate(p("CREATION_DATE")) %></td>
      <td class="num"><%= p("SONGS") %></td>
      <td class="actions">
        <%= EditButton("admin_playlist_form.asp?id=" & p("ID_PLAYLIST"), p("TITLE")) %>
        <%= DeleteButton(p("ID_PLAYLIST"), "la playlist «" & p("TITLE") & "»") %>
      </td>
    </tr>
  <% Next %>
  </tbody>
</table>
</div>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
