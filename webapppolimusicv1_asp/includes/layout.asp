<%
'==============================================================================
' PoliMusic - Componentes de interfaz (layout, iconos, tarjetas)
'==============================================================================
' Todo el marcado HTML compartido vive aqui para que las paginas solo Declaren
' su contenido. Los iconos son SVG en linea: no hay peticiones extra ni
' dependencias de fuentes externas.
'==============================================================================

'==============================================================================
' Icon() - Devuelve el SVG de un icono con el tamano por defecto.
'==============================================================================
Function Icon(name)
    Icon = IconSized(name, 0)
End Function

'==============================================================================
' IconSized() - SVG en linea. Tamanos: 0 = 24px (por defecto).
'==============================================================================
Function IconSized(name, size)
    Dim s, st, common

    If size <= 0 Then size = 24
    st = "width=""" & size & """ height=""" & size & """"
    common = " xmlns=""http://www.w3.org/2000/svg"" viewBox=""0 0 24 24"" " & _
             "fill=""none"" stroke=""currentColor"" stroke-width=""1.8"" " & _
             "stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"" class=""pm-ico"""

    Select Case LCase(CStr(name))

    Case "home"
        IconSized = "<svg" & common & " " & st & "><path d=""M3 10.5 12 3l9 7.5""/><path d=""M5 9.5V21h14V9.5""/><path d=""M9.5 21v-6h5v6""/></svg>"
    Case "music", "song"
        IconSized = "<svg" & common & " " & st & "><path d=""M9 18V5l11-2v13""/><circle cx=""6"" cy=""18"" r=""3""/><circle cx=""17"" cy=""16"" r=""3""/></svg>"
    Case "album"
        IconSized = "<svg" & common & " " & st & "><circle cx=""12"" cy=""12"" r=""9""/><circle cx=""12"" cy=""12"" r=""3""/><path d=""M12 3v6""/></svg>"
    Case "artist", "mic"
        IconSized = "<svg" & common & " " & st & "><path d=""M12 15a4 4 0 0 0 4-4V7a4 4 0 0 0-8 0v4a4 4 0 0 0 4 4Z""/><path d=""M5 11a7 7 0 0 0 14 0""/><path d=""M12 18v3""/><path d=""M8.5 21h7""/></svg>"
    Case "playlist", "list"
        IconSized = "<svg" & common & " " & st & "><path d=""M4 6h11""/><path d=""M4 11h11""/><path d=""M4 16h7""/><circle cx=""17"" cy=""17"" r=""3""/><path d=""M20 17V9l-2 1""/></svg>"
    Case "play"
        IconSized = "<svg" & common & " " & st & " stroke=""none"" fill=""currentColor""><path d=""M8 5.5v13a.6.6 0 0 0 .93.5l10-6.5a.6.6 0 0 0 0-1L8.93 5a.6.6 0 0 0-.93.5Z""/></svg>"
    Case "pause"
        IconSized = "<svg" & common & " " & st & " stroke=""none"" fill=""currentColor""><rect x=""7"" y=""5"" width=""3.6"" height=""14"" rx=""1.1""/><rect x=""13.4"" y=""5"" width=""3.6"" height=""14"" rx=""1.1""/></svg>"
    Case "next"
        IconSized = "<svg" & common & " " & st & " stroke=""none"" fill=""currentColor""><path d=""M6 5.6v12.8a.6.6 0 0 0 .92.5l8.7-6.4a.6.6 0 0 0 0-1L6.92 5.1A.6.6 0 0 0 6 5.6Z""/><rect x=""17"" y=""5.4"" width=""2.4"" height=""13.2"" rx=""1""/></svg>"
    Case "prev"
        IconSized = "<svg" & common & " " & st & " stroke=""none"" fill=""currentColor""><path d=""M18 5.6v12.8a.6.6 0 0 1-.92.5l-8.7-6.4a.6.6 0 0 1 0-1l8.7-6.4a.6.6 0 0 1 .92.5Z""/><rect x=""4.6"" y=""5.4"" width=""2.4"" height=""13.2"" rx=""1""/></svg>"
    Case "shuffle"
        IconSized = "<svg" & common & " " & st & "><path d=""M16 4h4v4""/><path d=""M4 20 20 4""/><path d=""M16 20h4v-4""/><path d=""M4 4l5 5""/><path d=""M15 15l5 5""/></svg>"
    Case "repeat"
        IconSized = "<svg" & common & " " & st & "><path d=""M17 3l3 3-3 3""/><path d=""M20 6H8a4 4 0 0 0-4 4v1""/><path d=""M7 21l-3-3 3-3""/><path d=""M4 18h12a4 4 0 0 0 4-4v-1""/></svg>"
    Case "volume"
        IconSized = "<svg" & common & " " & st & "><path d=""M4 9.5h3.2L12 5.4v13.2L7.2 14.5H4Z""/><path d=""M16 9a4.5 4.5 0 0 1 0 6""/><path d=""M18.6 6.4a8 8 0 0 1 0 11.2""/></svg>"
    Case "mute"
        IconSized = "<svg" & common & " " & st & "><path d=""M4 9.5h3.2L12 5.4v13.2L7.2 14.5H4Z""/><path d=""M16.5 9.5l5 5""/><path d=""M21.5 9.5l-5 5""/></svg>"
    Case "search"
        IconSized = "<svg" & common & " " & st & "><circle cx=""11"" cy=""11"" r=""7""/><path d=""M20 20l-3.6-3.6""/></svg>"
    Case "user"
        IconSized = "<svg" & common & " " & st & "><circle cx=""12"" cy=""8"" r=""4""/><path d=""M4.5 21a7.5 7.5 0 0 1 15 0""/></svg>"
    Case "users"
        IconSized = "<svg" & common & " " & st & "><circle cx=""9"" cy=""8"" r=""3.5""/><path d=""M2.5 20a6.5 6.5 0 0 1 13 0""/><path d=""M16 5.2a3.5 3.5 0 0 1 0 6.6""/><path d=""M17.5 14.2A6.5 6.5 0 0 1 21.5 20""/></svg>"
    Case "settings", "gear"
        IconSized = "<svg" & common & " " & st & "><circle cx=""12"" cy=""12"" r=""3""/><path d=""M12 3v2.2""/><path d=""M12 18.8V21""/><path d=""M3 12h2.2""/><path d=""M18.8 12H21""/><path d=""M5.6 5.6l1.6 1.6""/><path d=""M16.8 16.8l1.6 1.6""/><path d=""M18.4 5.6l-1.6 1.6""/><path d=""M7.2 16.8l-1.6 1.6""/></svg>"
    Case "plus", "add"
        IconSized = "<svg" & common & " " & st & "><path d=""M12 5v14""/><path d=""M5 12h14""/></svg>"
    Case "edit", "pencil"
        IconSized = "<svg" & common & " " & st & "><path d=""M4 20h4l10-10a2.5 2.5 0 0 0-3.5-3.5L4.5 16.5V20Z""/><path d=""M13.5 7l3.5 3.5""/></svg>"
    Case "trash", "delete"
        IconSized = "<svg" & common & " " & st & "><path d=""M4 7h16""/><path d=""M9 7V4.8h6V7""/><path d=""M6 7l1 13h10l1-13""/><path d=""M10 11v5""/><path d=""M14 11v5""/></svg>"
    Case "logout", "exit"
        IconSized = "<svg" & common & " " & st & "><path d=""M15 4h3a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2h-3""/><path d=""M10 16l4-4-4-4""/><path d=""M14 12H3""/></svg>"
    Case "login"
        IconSized = "<svg" & common & " " & st & "><path d=""M9 4H6a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h3""/><path d=""M14 8l4 4-4 4""/><path d=""M18 12H7""/></svg>"
    Case "heart"
        IconSized = "<svg" & common & " " & st & "><path d=""M12 20s-7.5-4.6-7.5-9.6A4.4 4.4 0 0 1 12 7.6a4.4 4.4 0 0 1 7.5 2.8C19.5 15.4 12 20 12 20Z""/></svg>"
    Case "sparkle", "star"
        IconSized = "<svg" & common & " " & st & "><path d=""M12 3.5l2.2 5.1 5.3.5-4 3.6 1.2 5.3L12 15.4l-4.7 2.6 1.2-5.3-4-3.6 5.3-.5Z""/></svg>"
    Case "chart", "stats"
        IconSized = "<svg" & common & " " & st & "><path d=""M4 20h16""/><path d=""M7 20V11""/><path d=""M12 20V5""/><path d=""M17 20v-6""/></svg>"
    Case "shield"
        IconSized = "<svg" & common & " " & st & "><path d=""M12 3l7 3v6c0 4.4-3 7.7-7 9-4-1.3-7-4.6-7-9V6Z""/><path d=""M9.2 12l2 2 3.6-4""/></svg>"
    Case "alert", "warning"
        IconSized = "<svg" & common & " " & st & "><path d=""M12 4.5 21 19.5H3Z""/><path d=""M12 10v4""/><path d=""M12 17h.01""/></svg>"
    Case "info"
        IconSized = "<svg" & common & " " & st & "><circle cx=""12"" cy=""12"" r=""9""/><path d=""M12 11v5""/><path d=""M12 8h.01""/></svg>"
    Case "check"
        IconSized = "<svg" & common & " " & st & "><path d=""M5 12.5l4.5 4.5L19 7""/></svg>"
    Case "close", "x"
        IconSized = "<svg" & common & " " & st & "><path d=""M6 6l12 12""/><path d=""M18 6 6 18""/></svg>"
    Case "upload"
        IconSized = "<svg" & common & " " & st & "><path d=""M12 16V4""/><path d=""M8 8l4-4 4 4""/><path d=""M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3""/></svg>"
    Case "image", "photo"
        IconSized = "<svg" & common & " " & st & "><rect x=""3.5"" y=""5"" width=""17"" height=""14"" rx=""2.5""/><circle cx=""9"" cy=""10.5"" r=""1.8""/><path d=""M4.5 17l5-5 5 5 2.5-2.5 3 3""/></svg>"
    Case "clock"
        IconSized = "<svg" & common & " " & st & "><circle cx=""12"" cy=""12"" r=""9""/><path d=""M12 7.5V12l3 2""/></svg>"
    Case "calendar"
        IconSized = "<svg" & common & " " & st & "><rect x=""3.5"" y=""5"" width=""17"" height=""15"" rx=""2.5""/><path d=""M8 3v4""/><path d=""M16 3v4""/><path d=""M3.5 10h17""/></svg>"
    Case "menu", "bars"
        IconSized = "<svg" & common & " " & st & "><path d=""M4 7h16""/><path d=""M4 12h16""/><path d=""M4 17h16""/></svg>"
    Case "chevron-left"
        IconSized = "<svg" & common & " " & st & "><path d=""M15 5l-7 7 7 7""/></svg>"
    Case "chevron-right"
        IconSized = "<svg" & common & " " & st & "><path d=""M9 5l7 7-7 7""/></svg>"
    Case "arrow-left"
        IconSized = "<svg" & common & " " & st & "><path d=""M19 12H5""/><path d=""M11 6l-6 6 6 6""/></svg>"
    Case "lock"
        IconSized = "<svg" & common & " " & st & "><rect x=""5"" y=""10.5"" width=""14"" height=""9.5"" rx=""2""/><path d=""M8.5 10.5V8a3.5 3.5 0 0 1 7 0v2.5""/></svg>"
    Case "globe", "public"
        IconSized = "<svg" & common & " " & st & "><circle cx=""12"" cy=""12"" r=""9""/><path d=""M3.5 12h17""/><path d=""M12 3c2.4 2.5 3.6 5.5 3.6 9s-1.2 6.5-3.6 9c-2.4-2.5-3.6-5.5-3.6-9S9.6 5.5 12 3Z""/></svg>"
    Case "headphones"
        IconSized = "<svg" & common & " " & st & "><path d=""M4 14v-2a8 8 0 0 1 16 0v2""/><rect x=""3"" y=""13.5"" width=""4"" height=""7"" rx=""2""/><rect x=""17"" y=""13.5"" width=""4"" height=""7"" rx=""2""/></svg>"
    Case "disc"
        IconSized = "<svg" & common & " " & st & "><circle cx=""12"" cy=""12"" r=""9""/><circle cx=""12"" cy=""12"" r=""1.6"" fill=""currentColor"" stroke=""none""/></svg>"
    Case "fire", "trending"
        IconSized = "<svg" & common & " " & st & "><path d=""M12 3c.5 3 3 4 3 7a3 3 0 0 1-6 0c0-1.2.5-2 .5-2S6 10.5 6 14a6 6 0 0 0 12 0c0-5-4-6.5-6-11Z""/></svg>"
    Case "dashboard"
        IconSized = "<svg" & common & " " & st & "><rect x=""3.5"" y=""3.5"" width=""7"" height=""7"" rx=""1.8""/><rect x=""13.5"" y=""3.5"" width=""7"" height=""7"" rx=""1.8""/><rect x=""3.5"" y=""13.5"" width=""7"" height=""7"" rx=""1.8""/><rect x=""13.5"" y=""13.5"" width=""7"" height=""7"" rx=""1.8""/></svg>"
    Case "save"
        IconSized = "<svg" & common & " " & st & "><path d=""M5 4h11l3 3v13H5Z""/><path d=""M9 4v5h6V4""/><rect x=""8"" y=""13"" width=""8"" height=""7""/></svg>"
    Case "filter"
        IconSized = "<svg" & common & " " & st & "><path d=""M4 6h16""/><path d=""M7 12h10""/><path d=""M10 18h4""/></svg>"
    Case "download"
        IconSized = "<svg" & common & " " & st & "><path d=""M12 4v11""/><path d=""M8 11l4 4 4-4""/><path d=""M5 20h14""/></svg>"

    Case Else
        ' Icono desconocido: un pequeño cuadrado para que se note en desarrollo
        IconSized = "<svg" & common & " " & st & "><rect x=""5"" y=""5"" width=""14"" height=""14"" rx=""2""/></svg>"

    End Select
End Function

'==============================================================================
' BrandMark() - Logotipo de la aplicacion.
'==============================================================================
Function BrandMark()
    BrandMark = _
        "<svg class=""pm-brand-mark"" width=""38"" height=""38"" viewBox=""0 0 48 48"" " & _
        "fill=""none"" xmlns=""http://www.w3.org/2000/svg"" aria-hidden=""true"">" & _
        "<rect width=""48"" height=""48"" rx=""14"" fill=""url(#pmBrandGrad)""/>" & _
        "<path d=""M19 32V17l14-3v13"" stroke=""#fff"" stroke-width=""3"" " & _
        "stroke-linecap=""round"" stroke-linejoin=""round"" fill=""none""/>" & _
        "<circle cx=""16"" cy=""32"" r=""4"" fill=""#fff""/>" & _
        "<circle cx=""30"" cy=""29"" r=""4"" fill=""#fff""/>" & _
        "<defs><linearGradient id=""pmBrandGrad"" x1=""0"" y1=""0"" x2=""48"" y2=""48"">" & _
        "<stop offset=""0"" stop-color=""#A78BFA""/><stop offset=""1"" stop-color=""#7C3AED""/>" & _
        "</linearGradient></defs></svg>"
End Function

'==============================================================================
' RenderHead() - Cabecera <head> comun a las paginas de la aplicacion.
'==============================================================================
Sub RenderHead(pageTitle)
%>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#12101C">
<title><% Response.Write H(pageTitle) %> | PoliMusic</title>
<link rel="stylesheet" href="assets/css/polimusic.css">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 48 48'%3E%3Crect width='48' height='48' rx='14' fill='%237C3AED'/%3E%3Cpath d='M19 32V17l14-3v13' stroke='white' stroke-width='3' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3Ccircle cx='16' cy='32' r='4' fill='white'/%3E%3Ccircle cx='30' cy='29' r='4' fill='white'/%3E%3C/svg%3E">
<%
End Sub

'==============================================================================
' NavItem() - Un enlace de la barra lateral.
'==============================================================================
Function NavItem(currentKey, itemKey, href, iconName, label)
    Dim cls, active
    active = ""
    If LCase(CStr(currentKey)) = LCase(CStr(itemKey)) Then active = " is-active"

    NavItem = _
        "<a class=""pm-nav-item" & active & """ href=""" & Ha(href) & """" & _
        " aria-current=""" & IIf(Len(active) > 0, "page", "false") & """>" & _
        "<span class=""pm-nav-ico"">" & Icon(iconName) & "</span>" & _
        "<span class=""pm-nav-label"">" & H(label) & "</span></a>"
End Function

'==============================================================================
' RenderSidebar() - Barra lateral de navegacion principal.
'==============================================================================
Sub RenderSidebar(currentKey)
%>
<aside class="pm-sidebar" id="pmSidebar">
    <div class="pm-sidebar-brand">
        <a href="default.asp" class="pm-brand">
            <% Response.Write BrandMark() %>
            <span class="pm-brand-text">
                <span class="pm-brand-name">PoliMusic</span>
                <span class="pm-brand-sub">Escuela Politecnica Nacional</span>
            </span>
        </a>
    </div>

    <nav class="pm-nav" aria-label="Navegacion principal">
        <p class="pm-nav-title">Explorar</p>
        <% Response.Write NavItem(currentKey, "home", "default.asp", "home", "Inicio") %>
        <% Response.Write NavItem(currentKey, "songs", "songs.asp", "music", "Canciones") %>
        <% Response.Write NavItem(currentKey, "albums", "albums.asp", "album", "Albumes") %>
        <% Response.Write NavItem(currentKey, "artists", "artists.asp", "artist", "Artistas") %>
        <% Response.Write NavItem(currentKey, "playlists", "playlists.asp", "playlist", "Playlists") %>

        <% If IsAdmin() Then %>
        <p class="pm-nav-title">Administracion</p>
        <% Response.Write NavItem(currentKey, "admin", "admin/default.asp", "dashboard", "Panel") %>
        <% Response.Write NavItem(currentKey, "admin-songs", "admin/songs.asp", "music", "Gestionar canciones") %>
        <% Response.Write NavItem(currentKey, "admin-albums", "admin/albums.asp", "album", "Gestionar albumes") %>
        <% Response.Write NavItem(currentKey, "admin-artists", "admin/artists.asp", "artist", "Gestionar artistas") %>
        <% Response.Write NavItem(currentKey, "admin-playlists", "admin/playlists.asp", "playlist", "Gestionar playlists") %>
        <% Response.Write NavItem(currentKey, "admin-users", "admin/users.asp", "users", "Gestionar usuarios") %>
        <% End If %>
    </nav>

    <div class="pm-sidebar-foot">
        <div class="pm-sidebar-user">
            <span class="pm-avatar pm-avatar-sm">
                <% If Len(Trim(CurrentUserPhoto())) > 0 Then %>
                <img src="<% Response.Write Ha(UserPhotoUrl(CurrentUserPhoto())) %>" alt="">
                <% Else %>
                <span class="pm-avatar-initials"><% Response.Write H(UserInitials(CurrentUserName())) %></span>
                <% End If %>
            </span>
            <span class="pm-sidebar-user-text">
                <span class="pm-sidebar-user-name"><% Response.Write H(CurrentUserName()) %></span>
                <span class="pm-sidebar-user-role"><% Response.Write H(RoleLabel()) %></span>
            </span>
        </div>
        <a class="pm-btn pm-btn-ghost pm-btn-sm pm-btn-block" href="logout.asp">
            <% Response.Write Icon("logout") %><span>Salir</span>
        </a>
    </div>
</aside>
<%
End Sub

'==============================================================================
' RenderTopbar() - Barra superior con buscador y accesos de cuenta.
'==============================================================================
Sub RenderTopbar(pageTitle, pageSubtitle, actionsHtml)
%>
<header class="pm-topbar">
    <button class="pm-icon-btn pm-sidebar-toggle" type="button"
            data-pm-sidebar-toggle aria-label="Mostrar u ocultar el menu">
        <% Response.Write Icon("menu") %>
    </button>

    <div class="pm-topbar-text">
        <h1 class="pm-page-title"><% Response.Write H(pageTitle) %></h1>
        <% If Len(pageSubtitle) > 0 Then %>
        <p class="pm-page-sub"><% Response.Write H(pageSubtitle) %></p>
        <% End If %>
    </div>

    <form class="pm-search" method="get" action="search.asp" role="search">
        <span class="pm-search-ico"><% Response.Write Icon("search") %></span>
        <input class="pm-search-input" type="search" name="q" placeholder="Buscar canciones, albumes o artistas"
               aria-label="Buscar en el catalogo" maxlength="60">
    </form>

    <div class="pm-topbar-actions">
        <% If Len(actionsHtml) > 0 Then Response.Write actionsHtml %>
        <a class="pm-icon-btn" href="profile.asp" title="Mi perfil" aria-label="Mi perfil">
            <% Response.Write Icon("user") %>
        </a>
    </div>
</header>
<%
End Sub

'==============================================================================
' RenderFlash() - Mensajes de exito o error tras una redireccion.
'==============================================================================
Sub RenderFlash()
    Dim kind, msg
    kind = TakeFlashKind()
    msg = TakeFlashMsg()

    If Len(msg) = 0 Then Exit Sub

    Dim cls, icoName
    If LCase(kind) = "error" Then
        cls = "pm-flash-error"
        icoName = "alert"
    ElseIf LCase(kind) = "info" Then
        cls = "pm-flash-info"
        icoName = "info"
    Else
        cls = "pm-flash-success"
        icoName = "check"
    End If

    Response.Write "<div class=""pm-flash " & cls & """ role=""status"" data-pm-autohide>" & _
                   Icon(icoName) & "<span>" & H(msg) & "</span>" & _
                   "<button type=""button"" class=""pm-flash-close"" aria-label=""Cerrar"" " & _
                   "onclick=""this.parentNode.remove()"">" & Icon("close") & "</button></div>"
End Sub

'==============================================================================
' RenderPlayer() - Reproductor fijo en la parte inferior.
'==============================================================================
Sub RenderPlayer()
%>
<div class="pm-player" id="pmPlayer" aria-label="Reproductor de musica">
    <audio id="pmAudio" preload="metadata"></audio>

    <div class="pm-player-left">
        <span class="pm-player-cover" id="pmPlayerCover">
            <span class="pm-player-cover-fallback"><% Response.Write Icon("disc") %></span>
        </span>
        <span class="pm-player-meta">
            <span class="pm-player-title" id="pmPlayerTitle">Nada en reproduccion</span>
            <span class="pm-player-artist" id="pmPlayerArtist">Elige una cancion del catalogo</span>
        </span>
    </div>

    <div class="pm-player-center">
        <div class="pm-player-buttons">
            <button class="pm-icon-btn" type="button" id="pmShuffle"
                    title="Aleatorio" aria-label="Aleatorio">
                <% Response.Write Icon("shuffle") %>
            </button>
            <button class="pm-icon-btn" type="button" id="pmPrev"
                    title="Anterior" aria-label="Cancion anterior">
                <% Response.Write Icon("prev") %>
            </button>
            <button class="pm-icon-btn pm-icon-btn-main" type="button" id="pmToggle"
                    title="Reproducir o pausar" aria-label="Reproducir o pausar">
                <span id="pmToggleIco"><% Response.Write Icon("play") %></span>
            </button>
            <button class="pm-icon-btn" type="button" id="pmNext"
                    title="Siguiente" aria-label="Cancion siguiente">
                <% Response.Write Icon("next") %>
            </button>
            <button class="pm-icon-btn" type="button" id="pmRepeat"
                    title="Repetir" aria-label="Repetir">
                <% Response.Write Icon("repeat") %>
            </button>
        </div>

        <div class="pm-player-progress">
            <span class="pm-player-time" id="pmTimeNow">0:00</span>
            <input class="pm-seek" type="range" id="pmSeek" min="0" max="100" value="0" step="0.1"
                   aria-label="Posicion de la cancion">
            <span class="pm-player-time" id="pmTimeTotal">0:00</span>
        </div>
    </div>

    <div class="pm-player-right">
        <button class="pm-icon-btn" type="button" id="pmMute"
                title="Silenciar" aria-label="Silenciar">
            <span id="pmVolumeIco"><% Response.Write Icon("volume") %></span>
        </button>
        <input class="pm-volume" type="range" id="pmVolume" min="0" max="1" step="0.01" value="0.8"
               aria-label="Volumen">
    </div>
</div>
<%
End Sub

'==============================================================================
' RenderFooter() - Pie de pagina.
'==============================================================================
Sub RenderFooter()
%>
<footer class="pm-footer">
    <p>PoliMusic &middot; Proyecto academico &middot; Escuela Politecnica Nacional</p>
    <p class="pm-footer-dim">Aplicacion web desarrollada en ASP clasico con SQL Server.</p>
</footer>
<%
End Sub

'==============================================================================
' RenderEmpty() - Estado vacio reutilizable.
'==============================================================================
Sub RenderEmpty(iconName, title, message, actionHref, actionLabel)
    Response.Write "<div class=""pm-empty"">" & _
        "<span class=""pm-empty-ico"">" & Icon(iconName) & "</span>" & _
        "<h3>" & H(title) & "</h3>" & _
        "<p>" & H(message) & "</p>"
    If Len(actionHref) > 0 Then
        Response.Write "<a class=""pm-btn pm-btn-primary"" href=""" & Ha(actionHref) & """>" & _
                       H(actionLabel) & "</a>"
    End If
    Response.Write "</div>"
End Sub

'==============================================================================
' IIf() - Operador ternario (VBScript no lo trae de fabrica).
'==============================================================================
Function IIf(condition, ifTrue, ifFalse)
    If CBool(condition) Then
        IIf = ifTrue
    Else
        IIf = ifFalse
    End If
End Function

'==============================================================================
' RenderDeleteForm() - Formulario POST minimo para borrar con confirmacion.
'==============================================================================
Function RenderDeleteForm(action, idField, entityId, label)
    RenderDeleteForm = _
        "<form method=""post"" action=""" & Ha(action) & """ class=""pm-inline-form"" " & _
        "onsubmit=""return pmConfirmDelete(event, '" & JsArg(label) & "')"">" & _
        CsrfField() & _
        "<input type=""hidden"" name=""" & Ha(idField) & """ value=""" & Ha(entityId) & """>" & _
        "<button class=""pm-btn pm-btn-danger pm-btn-sm"" type=""submit"" title=""Eliminar"">" & _
        Icon("trash") & "<span>Eliminar</span></button></form>"
End Function
%>
