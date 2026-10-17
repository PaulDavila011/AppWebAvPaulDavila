<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireLogin
pageTitle = "Buscar"
activeNav = "search"

Dim q, pat, songs, albums, artists, a, ar, total
q = QS("q")
If Len(q) > 50 Then q = Left(q, 50)
total = 0
If q <> "" Then
    pageTitle = "Buscar: " & q
    pat = "%" & q & "%"
    songs = QList("SELECT TOP 20 * FROM (" & SongSelectSql() & ") s WHERE s.SONG_NAME LIKE ? OR s.ARTISTS LIKE ? ORDER BY ISNULL(s.PLAYS,0) DESC", Array(pat, pat))
    albums = QList("SELECT TOP 12 al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH FROM TBL_ALBUM al WHERE al.TITLE_ALBUM LIKE ? ORDER BY al.TITLE_ALBUM", Array(pat))
    artists = QList("SELECT TOP 12 a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE FROM TBL_ARTIST a WHERE (a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME,'')) LIKE ? ORDER BY a.ARTIST_NAME", Array(pat))
    total = UBound(songs) + UBound(albums) + UBound(artists) + 3
End If
%>
<!--#include file="includes/layout_top.asp"-->
<header class="page-head">
  <div>
    <p class="eyebrow"><%= Icon("search") %> Buscar</p>
    <h1><% If q = "" Then %>¿Qué quieres escuchar?<% Else %>Resultados para “<%= H(q) %>”<% End If %></h1>
    <% If q <> "" Then %><p class="muted"><%= Plural(total, "resultado", "resultados") %></p><% End If %>
  </div>
</header>
<div class="toolbar-row"><% RenderSearchBar "Canciones, álbumes o artistas", q %></div>

<% If q = "" Then %>
  <div class="browse-grid">
    <a class="browse-tile t1" href="songs.asp?sort=plays"><span>Lo más escuchado</span><%= Icon("headphones") %></a>
    <a class="browse-tile t2" href="songs.asp?sort=new"><span>Novedades</span><%= Icon("music") %></a>
    <a class="browse-tile t3" href="albums.asp"><span>Álbumes</span><%= Icon("disc") %></a>
    <a class="browse-tile t4" href="artists.asp"><span>Artistas</span><%= Icon("mic") %></a>
    <a class="browse-tile t5" href="playlists.asp"><span>Playlists</span><%= Icon("list") %></a>
  </div>
<% ElseIf total = 0 Then
     RenderEmpty "search", "No encontramos nada para esa búsqueda. Prueba con otras palabras."
   Else %>
  <% If UBound(songs) >= 0 Then %>
  <section class="section"><div class="section-head"><h2>Canciones</h2></div><div id="search-songs"><% RenderTrackList songs, True, "" %></div></section>
  <% End If %>
  <% If UBound(artists) >= 0 Then %>
  <section class="section"><div class="section-head"><h2>Artistas</h2></div><div class="card-row">
    <% For Each ar In artists : RenderCard "artist.asp?id=" & ar("ID_ARTIST"), ar("ARTIST_IMAGE"), "artist", ArtistFullName(ar), "Artista" : Next %>
  </div></section>
  <% End If %>
  <% If UBound(albums) >= 0 Then %>
  <section class="section"><div class="section-head"><h2>Álbumes</h2></div><div class="card-row">
    <% For Each a In albums : RenderCard "album.asp?id=" & a("ID_ALBUM"), a("ALBUM_IMAGE_PATH"), "cover", a("TITLE_ALBUM"), Nz(a("RELEASE_YEAR"), "—") & " · Álbum" : Next %>
  </div></section>
  <% End If %>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
