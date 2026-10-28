<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
'==============================================================================
' PoliMusic - Catalogo de artistas
'==============================================================================
Dim artists, totalArtists

Set artists = ArtistsAll()
totalArtists = CountRows("TBL_ARTIST", "")

Dim actionsHtml
actionsHtml = ""
If IsAdmin() Then
    actionsHtml = "<a class=""pm-btn pm-btn-primary pm-btn-sm"" href=""admin/artists.asp?action=new"">" & _
                  Icon("plus") & "<span>Nuevo artista</span></a>"
End If

Response.Write "<!DOCTYPE html>" & vbCrLf
%>
<html lang="es">
<head>
<% RenderHead "Artistas" %>
</head>
<body class="pm-app">
<a class="pm-skip" href="#pmMain">Ir al contenido principal</a>

<% RenderSidebar "artists" %>

<div class="pm-main">
<% RenderTopbar "Artistas", CStr(totalArtists) & " " & IIf(totalArtists = 1, "artista", "artistas") & " en el catalogo", actionsHtml %>

<main class="pm-content" id="pmMain">
<% RenderFlash %>

    <% If artists.EOF Then %>
    <% RenderEmpty "artist", "No hay artistas registrados", _
                   "Todavia no se ha registrado ningun artista en el catalogo.", _
                   IIf(IsAdmin(), "admin/artists.asp?action=new", "default.asp"), _
                   IIf(IsAdmin(), "Registrar el primer artista", "Volver al inicio") %>
    <% Else %>

    <div class="pm-section-head">
        <div>
            <h2 class="pm-section-title">Todos los artistas</h2>
            <p class="pm-section-sub">Quienes interpretan la musica del catalogo</p>
        </div>
    </div>

    <div class="pm-grid">
        <%
        artists.MoveFirst
        Do While Not artists.EOF
            RenderArtistCard artists
            artists.MoveNext
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
artists.Close : Set artists = Nothing
%>
