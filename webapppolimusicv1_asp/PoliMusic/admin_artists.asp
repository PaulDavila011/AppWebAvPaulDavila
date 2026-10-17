<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireAdmin
pageTitle = "Gestión de artistas"
activeNav = "admin_artists"

Dim q, artists, ar, id, errMsg, dummy, oldMedia
If IsPost() Then
    RequireCsrf
    If FormValue("action") = "delete" Then
        id = ToInt(FormValue("id"), 0)
        ' Las canciones se conservan; solo se elimina la relación artista-canción
        oldMedia = QScalar("SELECT ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(id))
        errMsg = TryBatch("SET @out = ?; DELETE FROM TBL_PLAYER WHERE ID_ARTIST = @out; DELETE FROM TBL_ARTIST WHERE ID_ARTIST = @out;", Array(id), dummy)
        If errMsg = "" Then DeleteUploadedMedia oldMedia
        If errMsg = "" Then SetFlash "success", "Artista eliminado." Else SetFlash "error", errMsg
    End If
    RedirectTo "admin_artists.asp"
End If

q = QS("q")
artists = QList("SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, (SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS SONGS " & _
                "FROM TBL_ARTIST a WHERE (? = '' OR (a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME,'')) LIKE ?) ORDER BY a.ARTIST_NAME, a.ARTIST_LASTNAME", Array(q, "%" & q & "%"))
%>
<!--#include file="includes/layout_top.asp"-->
<% RenderAdminHeader "Artistas", "Gestiona los artistas del catálogo y su imagen.", "admin_artist_form.asp", "Nuevo artista" %>
<div class="toolbar-row"><% RenderSearchBar "Buscar artista", q %></div>
<% If UBound(artists) < 0 Then
     RenderEmpty "mic", "No hay artistas que coincidan."
   Else %>
<div class="table-wrap">
<table class="table">
  <thead><tr><th>Artista</th><th class="hide-sm">Nombre</th><th class="hide-sm">Apellido</th><th class="num">Canciones</th><th class="actions"><span class="sr-only">Acciones</span></th></tr></thead>
  <tbody>
  <% For Each ar In artists %>
    <tr>
      <td><div class="cell-main"><span class="avatar avatar-lg"><%= Img(ar("ARTIST_IMAGE"), "artist", "", "") %></span><div><strong><a href="artist.asp?id=<%= ar("ID_ARTIST") %>"><%= H(ArtistFullName(ar)) %></a></strong><small>ID <%= ar("ID_ARTIST") %></small></div></div></td>
      <td class="hide-sm"><%= H(ar("ARTIST_NAME")) %></td>
      <td class="hide-sm"><%= H(Nz(ar("ARTIST_LASTNAME"), "—")) %></td>
      <td class="num"><%= ar("SONGS") %></td>
      <td class="actions">
        <%= EditButton("admin_artist_form.asp?id=" & ar("ID_ARTIST"), ArtistFullName(ar)) %>
        <%= DeleteButton(ar("ID_ARTIST"), "a " & ArtistFullName(ar) & IIfStr(CLng(ar("SONGS")) > 0, " (sus canciones se conservarán sin este artista)", "")) %>
      </td>
    </tr>
  <% Next %>
  </tbody>
</table>
</div>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
