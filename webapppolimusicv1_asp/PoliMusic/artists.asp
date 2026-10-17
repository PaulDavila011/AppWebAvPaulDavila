<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireLogin
pageTitle = "Artistas"
activeNav = "artists"

Dim q, artists, ar, sqlA
q = QS("q")
sqlA = "SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, (SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS SONGS FROM TBL_ARTIST a "
If q <> "" Then
    artists = QList(sqlA & "WHERE (a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME,'')) LIKE ? ORDER BY a.ARTIST_NAME, a.ARTIST_LASTNAME", Array("%" & q & "%"))
Else
    artists = QList(sqlA & "ORDER BY a.ARTIST_NAME, a.ARTIST_LASTNAME", Array())
End If
%>
<!--#include file="includes/layout_top.asp"-->
<header class="page-head">
  <div>
    <p class="eyebrow"><%= Icon("mic") %> Biblioteca</p>
    <h1>Artistas</h1>
    <p class="muted"><%= Plural(UBound(artists) + 1, "artista", "artistas") %><% If q <> "" Then %> para “<%= H(q) %>”<% End If %></p>
  </div>
</header>
<div class="toolbar-row"><% RenderSearchBar "Buscar artista", q %></div>
<% If UBound(artists) < 0 Then
     RenderEmpty "mic", "No se encontraron artistas."
   Else %>
<div class="card-grid">
  <% For Each ar In artists
       RenderCard "artist.asp?id=" & ar("ID_ARTIST"), ar("ARTIST_IMAGE"), "artist", ArtistFullName(ar), Plural(ar("SONGS"), "canción", "canciones")
     Next %>
</div>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
