<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireLogin
pageTitle = "Álbumes"
activeNav = "albums"

Dim q, albums, a, sqlA
q = QS("q")
sqlA = "SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, (SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS SONGS FROM TBL_ALBUM al "
If q <> "" Then
    albums = QList(sqlA & "WHERE al.TITLE_ALBUM LIKE ? OR CAST(al.RELEASE_YEAR AS varchar(10)) = ? ORDER BY al.TITLE_ALBUM", Array("%" & q & "%", q))
Else
    albums = QList(sqlA & "ORDER BY al.TITLE_ALBUM", Array())
End If
%>
<!--#include file="includes/layout_top.asp"-->
<header class="page-head">
  <div>
    <p class="eyebrow"><%= Icon("disc") %> Biblioteca</p>
    <h1>Álbumes</h1>
    <p class="muted"><%= Plural(UBound(albums) + 1, "álbum", "álbumes") %><% If q <> "" Then %> para “<%= H(q) %>”<% End If %></p>
  </div>
</header>
<div class="toolbar-row"><% RenderSearchBar "Buscar álbum por título o año", q %></div>
<% If UBound(albums) < 0 Then
     RenderEmpty "disc", "No se encontraron álbumes."
   Else %>
<div class="card-grid">
  <% For Each a In albums
       RenderCard "album.asp?id=" & a("ID_ALBUM"), a("ALBUM_IMAGE_PATH"), "cover", a("TITLE_ALBUM"), Nz(a("RELEASE_YEAR"), "—") & " · " & Plural(a("SONGS"), "canción", "canciones")
     Next %>
</div>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
