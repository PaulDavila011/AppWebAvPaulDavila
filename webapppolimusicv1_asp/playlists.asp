<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
'==============================================================================
' PoliMusic - Listado de playlists
'==============================================================================
' El usuario normal ve las playlists publicas mas las propias.
' El administrador ve absolutamente todas.
'==============================================================================
Dim playlists, totalVisible

If IsAdmin() Then
    Set playlists = PlaylistsAll()
Else
    Set playlists = PlaylistsVisible(CurrentUserId())
End If

Dim actionsHtml
actionsHtml = "<a class=""pm-btn pm-btn-primary pm-btn-sm"" href=""playlist_edit.asp"">" & _
              Icon("plus") & "<span>Nueva playlist</span></a>"

Response.Write "<!DOCTYPE html>" & vbCrLf
%>
<html lang="es">
<head>
<% RenderHead "Playlists" %>
</head>
<body class="pm-app">
<a class="pm-skip" href="#pmMain">Ir al contenido principal</a>

<% RenderSidebar "playlists" %>

<div class="pm-main">
<% RenderTopbar "Playlists", "Tus listas y las publicas de la comunidad", actionsHtml %>

<main class="pm-content" id="pmMain">
<% RenderFlash %>

    <% If playlists.EOF Then %>
    <% RenderEmpty "playlist", "Todavia no hay playlists", _
                   "Crea tu primera playlist para agrupar las canciones que mas te gustan.", _
                   "playlist_edit.asp", "Crear una playlist" %>
    <% Else %>

    <div class="pm-section-head">
        <div>
            <h2 class="pm-section-title">
                <% If IsAdmin() Then %>Todas las playlists<% Else %>Playlists disponibles<% End If %>
            </h2>
            <p class="pm-section-sub">
                <% If IsAdmin() Then %>
                Como administrador ves todas las playlists del sistema.
                <% Else %>
                Incluye las playlists publicas y las que son tuyas.
                <% End If %>
            </p>
        </div>
    </div>

    <div class="pm-grid">
        <%
        playlists.MoveFirst
        Do While Not playlists.EOF
            RenderPlaylistCard playlists
            playlists.MoveNext
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
playlists.Close : Set playlists = Nothing
%>
