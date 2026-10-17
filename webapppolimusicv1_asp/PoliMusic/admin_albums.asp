<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireAdmin
pageTitle = "Gestión de álbumes"
activeNav = "admin_albums"

Dim q, albums, a, id, errMsg, dummy, oldMedia
If IsPost() Then
    RequireCsrf
    If FormValue("action") = "delete" Then
        id = ToInt(FormValue("id"), 0)
        ' Las canciones se conservan; solo se elimina el tracklist del álbum
        oldMedia = QScalar("SELECT ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(id))
        errMsg = TryBatch("SET @out = ?; DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = @out; DELETE FROM TBL_ALBUM WHERE ID_ALBUM = @out;", Array(id), dummy)
        If errMsg = "" Then DeleteUploadedMedia oldMedia
        If errMsg = "" Then SetFlash "success", "Álbum eliminado." Else SetFlash "error", errMsg
    End If
    RedirectTo "admin_albums.asp"
End If

q = QS("q")
albums = QList("SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, (SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS SONGS " & _
               "FROM TBL_ALBUM al WHERE (? = '' OR al.TITLE_ALBUM LIKE ?) ORDER BY al.TITLE_ALBUM", Array(q, "%" & q & "%"))
%>
<!--#include file="includes/layout_top.asp"-->
<% RenderAdminHeader "Álbumes", "Gestiona álbumes, portadas y su lista de canciones.", "admin_album_form.asp", "Nuevo álbum" %>
<div class="toolbar-row"><% RenderSearchBar "Buscar álbum", q %></div>
<% If UBound(albums) < 0 Then
     RenderEmpty "disc", "No hay álbumes que coincidan."
   Else %>
<div class="table-wrap">
<table class="table">
  <thead><tr><th>Álbum</th><th class="num">Año</th><th class="num">Canciones</th><th class="actions"><span class="sr-only">Acciones</span></th></tr></thead>
  <tbody>
  <% For Each a In albums %>
    <tr>
      <td><div class="cell-main"><%= Img(a("ALBUM_IMAGE_PATH"), "cover", "", "thumb") %><div><strong><a href="album.asp?id=<%= a("ID_ALBUM") %>"><%= H(a("TITLE_ALBUM")) %></a></strong><small>ID <%= a("ID_ALBUM") %></small></div></div></td>
      <td class="num"><%= Nz(a("RELEASE_YEAR"), "—") %></td>
      <td class="num"><%= a("SONGS") %></td>
      <td class="actions">
        <%= EditButton("admin_album_form.asp?id=" & a("ID_ALBUM"), a("TITLE_ALBUM")) %>
        <%= DeleteButton(a("ID_ALBUM"), "el álbum «" & a("TITLE_ALBUM") & "» (las canciones se conservan)") %>
      </td>
    </tr>
  <% Next %>
  </tbody>
</table>
</div>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
