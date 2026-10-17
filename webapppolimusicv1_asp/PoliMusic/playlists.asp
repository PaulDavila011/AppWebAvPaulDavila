<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireLogin
pageTitle = "Playlists"
activeNav = "playlists"

Dim playlists, p, sqlP
sqlP = "SELECT pl.ID_PLAYLIST, pl.TITLE, pl.PHOTO, pl.TYPE, pl.CREATION_DATE, u.USERNAME, (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = pl.ID_PLAYLIST) AS SONGS " & _
       "FROM TBL_PLAYLIST pl JOIN TBL_USER u ON u.ID_USER = pl.ID_USER "
' Usuarios normales: playlists públicas y las propias. Administradores: todas.
If IsAdmin() Then
    playlists = QList(sqlP & "ORDER BY pl.CREATION_DATE DESC", Array())
Else
    playlists = QList(sqlP & "WHERE pl.TYPE = ? OR pl.ID_USER = ? ORDER BY pl.CREATION_DATE DESC", Array(PLAYLIST_PUBLIC, CurrentUserId()))
End If
%>
<!--#include file="includes/layout_top.asp"-->
<header class="page-head">
  <div>
    <p class="eyebrow"><%= Icon("list") %> Biblioteca</p>
    <h1>Playlists</h1>
    <p class="muted"><% If IsAdmin() Then %>Todas las playlists del sistema<% Else %>Playlists públicas y las tuyas<% End If %> · <%= Plural(UBound(playlists) + 1, "playlist", "playlists") %></p>
  </div>
  <% If IsAdmin() Then %><a class="btn btn-primary" href="admin_playlist_form.asp"><%= Icon("plus") %>Nueva playlist</a><% End If %>
</header>
<% If UBound(playlists) < 0 Then
     RenderEmpty "list", "Aún no hay playlists disponibles."
   Else %>
<div class="card-grid">
  <% For Each p In playlists
       RenderCard "playlist.asp?id=" & p("ID_PLAYLIST"), p("PHOTO"), "cover", p("TITLE"), "De " & p("USERNAME") & " · " & Plural(p("SONGS"), "canción", "canciones") & IIfStr(CLng(p("TYPE")) = PLAYLIST_PRIVATE, " · Privada", "")
     Next %>
</div>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
