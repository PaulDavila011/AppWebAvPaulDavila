<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
'==============================================================================
' PoliMusic - Catalogo de albumes
'==============================================================================
Dim albums, totalAlbums, shown

Set albums = AlbumsAll()
totalAlbums = CountRows("TBL_ALBUM", "")

Dim actionsHtml
actionsHtml = ""
If IsAdmin() Then
    actionsHtml = "<a class=""pm-btn pm-btn-primary pm-btn-sm"" href=""admin/albums.asp?action=new"">" & _
                  Icon("plus") & "<span>Nuevo album</span></a>"
End If

Response.Write "<!DOCTYPE html>" & vbCrLf
%>
<html lang="es">
<head>
<% RenderHead "Albumes" %>
</head>
<body class="pm-app">
<a class="pm-skip" href="#pmMain">Ir al contenido principal</a>

<% RenderSidebar "albums" %>

<div class="pm-main">
<% RenderTopbar "Albumes", CStr(totalAlbums) & " " & IIf(totalAlbums = 1, "album", "albumes") & " en el catalogo", actionsHtml %>

<main class="pm-content" id="pmMain">
<% RenderFlash %>

    <% If albums.EOF Then %>
    <% RenderEmpty "album", "No hay albumes registrados", _
                   "Todavia no se ha creado ningun album en el catalogo.", _
                   IIf(IsAdmin(), "admin/albums.asp?action=new", "default.asp"), _
                   IIf(IsAdmin(), "Crear el primer album", "Volver al inicio") %>
    <% Else %>

    <div class="pm-section-head">
        <div>
            <h2 class="pm-section-title">Todos los albumes</h2>
            <p class="pm-section-sub">Explora las colecciones completas del catalogo</p>
        </div>
    </div>

    <div class="pm-grid">
        <%
        albums.MoveFirst
        Do While Not albums.EOF
            RenderAlbumCard albums
            albums.MoveNext
        Loop
        %>
    </div>
    <% End If %>

</main>
<% RenderFooter %>
</div>

<% RenderPlayer %>
<div class="pm-sidebar-backdrop" data-pm-sidebar-toggle aria-hidden="true"></div>

<script src="assets/js/player.js"></script>
</body>
</html>
<%
albums.Close : Set albums = Nothing
%>
