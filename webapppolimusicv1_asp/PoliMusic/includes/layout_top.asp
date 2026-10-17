<%
Function NavLink(key, href, iconName, label)
    Dim cls
    cls = "nav-link"
    If key = activeNav Then cls = cls & " active"
    NavLink = "<a class=""" & cls & """ href=""" & href & """ data-nav=""" & key & """>" & Icon(iconName) & "<span>" & label & "</span></a>"
End Function
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="dark">
<meta name="theme-color" content="#0d0817">
<meta name="csrf-token" content="<%= H(CsrfToken()) %>">
<title><%= H(pageTitle) %> · <%= APP_NAME %></title>
<link rel="icon" href="assets/img/logo.svg" type="image/svg+xml">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="assets/css/app.css?v=<%= APP_VERSION %>">
<script>
/* Respaldo de imágenes faltantes (se registra antes de que carguen) */
document.addEventListener('error', function (e) {
  var t = e.target;
  if (t && t.tagName === 'IMG' && t.dataset.fallback && t.src.indexOf(t.dataset.fallback) === -1) { t.src = t.dataset.fallback; }
}, true);
</script>
</head>
<body class="<%= pageClass %>">
<!--#include file="icons.asp"-->
<div class="app" id="app">
  <aside class="sidebar" id="sidebar" aria-label="Navegación principal">
    <a class="brand" href="default.asp" data-nav="home">
      <img src="assets/img/logo.svg" alt="" width="36" height="36">
      <span>Poli<b>Music</b></span>
    </a>
    <nav class="nav">
      <p class="nav-label">Explorar</p>
      <%= NavLink("home", "default.asp", "home", "Inicio") %>
      <%= NavLink("search", "search.asp", "search", "Buscar") %>
      <%= NavLink("songs", "songs.asp", "music", "Canciones") %>
      <%= NavLink("albums", "albums.asp", "disc", "Álbumes") %>
      <%= NavLink("artists", "artists.asp", "mic", "Artistas") %>
      <%= NavLink("playlists", "playlists.asp", "list", "Playlists") %>
      <% If IsAdmin() Then %>
      <p class="nav-label">Administración</p>
      <%= NavLink("admin", "admin.asp", "dashboard", "Panel") %>
      <%= NavLink("admin_users", "admin_users.asp", "users", "Usuarios") %>
      <%= NavLink("admin_songs", "admin_songs.asp", "music", "Canciones") %>
      <%= NavLink("admin_artists", "admin_artists.asp", "mic", "Artistas") %>
      <%= NavLink("admin_albums", "admin_albums.asp", "disc", "Álbumes") %>
      <%= NavLink("admin_playlists", "admin_playlists.asp", "list", "Playlists") %>
      <% End If %>
    </nav>
    <div class="sidebar-foot">
      <div class="role-card">
        <% If IsAdmin() Then %>
        <%= Icon("shield") %><div><strong>Modo administrador</strong><small>Gestión completa del catálogo</small></div>
        <% Else %>
        <%= Icon("headphones") %><div><strong>Cuenta de oyente</strong><small>Escucha y explora el catálogo</small></div>
        <% End If %>
      </div>
    </div>
  </aside>
  <div class="scrim" data-close-menu></div>

  <div class="main-col">
    <header class="topbar">
      <button class="icon-btn menu-btn" type="button" data-open-menu aria-label="Abrir menú"><%= Icon("menu") %></button>
      <form class="top-search" action="search.asp" method="get" role="search" data-pjax-form>
        <%= Icon("search") %>
        <input type="search" name="q" placeholder="¿Qué quieres escuchar?" aria-label="Buscar canciones, álbumes o artistas" value="<% If activeNav = "search" Then Response.Write H(QS("q")) %>">
      </form>
      <details class="user-menu">
        <summary aria-label="Menú de usuario">
          <span class="avatar"><%= Img(Session("photo"), "user", "", "") %></span>
          <span class="user-meta"><strong><%= H(Session("uname")) %></strong><small><% If IsAdmin() Then Response.Write "Administrador" Else Response.Write "Usuario" %></small></span>
          <%= Icon("chevron") %>
        </summary>
        <div class="menu-pop">
          <div class="menu-head">
            <strong><%= H(Session("uname")) %></strong>
            <small><%= H(Session("email")) %></small>
            <%= RoleBadge(Session("role")) %>
          </div>
          <% If IsAdmin() Then %><a class="menu-item" href="admin.asp"><%= Icon("dashboard") %>Panel de administración</a><% End If %>
          <form method="post" action="logout.asp" data-no-pjax>
            <%= CsrfField() %>
            <button type="submit" class="menu-item danger"><%= Icon("logout") %>Cerrar sesión</button>
          </form>
        </div>
      </details>
    </header>

    <main id="content" class="content" data-nav="<%= activeNav %>" tabindex="-1">
<% RenderFlash %>
