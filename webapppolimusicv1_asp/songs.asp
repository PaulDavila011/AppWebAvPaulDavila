<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
'==============================================================================
' PoliMusic - Catalogo de canciones
'==============================================================================
Dim songs, queueJson, trackIndex, totalSongs

Set songs = SongsAll(0)
queueJson = TracksJsonSafe(songs)
totalSongs = CountRows("TBL_SONG", "")

Dim actionsHtml
actionsHtml = ""
If IsAdmin() Then
    actionsHtml = "<a class=""pm-btn pm-btn-primary pm-btn-sm"" href=""admin/songs.asp?action=new"">" & _
                  Icon("plus") & "<span>Nueva cancion</span></a>"
End If

Response.Write "<!DOCTYPE html>" & vbCrLf
%>
<html lang="es">
<head>
<% RenderHead "Canciones" %>
</head>
<body class="pm-app">
<a class="pm-skip" href="#pmMain">Ir al contenido principal</a>

<% RenderSidebar "songs" %>

<div class="pm-main">
<% RenderTopbar "Canciones", CStr(totalSongs) & " " & IIf(totalSongs = 1, "cancion en el catalogo", "canciones en el catalogo"), actionsHtml %>

<main class="pm-content" id="pmMain">
<% RenderFlash %>

    <% If songs.EOF Then %>
    <% RenderEmpty "music", "El catalogo esta vacio", _
                   "Todavia no hay canciones registradas. Un administrador puede agregarlas.", _
                   IIf(IsAdmin(), "admin/songs.asp?action=new", "default.asp"), _
                   IIf(IsAdmin(), "Agregar la primera cancion", "Volver al inicio") %>
    <% Else %>

    <div class="pm-section-head">
        <div>
            <h2 class="pm-section-title">Todas las canciones</h2>
            <p class="pm-section-sub">
                <%= totalSongs %> <% If totalSongs = 1 Then %>cancion<% Else %>canciones<% End If %>
                en el catalogo. Haz clic en cualquier fila para reproducirla.
            </p>
        </div>
        <button class="pm-btn pm-btn-primary pm-btn-sm" type="button"
                onclick="PoliMusicPlayer.load(window.PM_QUEUE, 0)">
            <% Response.Write Icon("play") %><span>Reproducir todo</span>
        </button>
    </div>

    <div class="pm-tracks">
        <% RenderTrackHead %>
        <%
        trackIndex = 0
        songs.MoveFirst
        Do While Not songs.EOF
            RenderTrackRow songs, queueJson, trackIndex
            trackIndex = trackIndex + 1
            songs.MoveNext
        Loop
        %>
    </div>
    <% End If %>

</main>
<% RenderFooter %>
</div>

<% RenderPlayer %>
<div class="pm-sidebar-backdrop" data-pm-sidebar-toggle aria-hidden="true"></div>
<% RenderDeleteModal %>

<script>
window.PM_QUEUE = <% Response.Write queueJson %>;
</script>
<script src="assets/js/player.js"></script>
</body>
</html>
<%
songs.Close : Set songs = Nothing
%>
