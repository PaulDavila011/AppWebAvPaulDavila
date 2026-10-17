<%
' =====================================================================
'  Componentes de interfaz reutilizables
' =====================================================================

' Lista de canciones reproducible. coverOverride: portada fija (p. ej. la del álbum) o ""
Sub RenderTrackList(songs, showAlbum, coverOverride)
    Dim s, i, cover, artists
    If UBound(songs) < 0 Then
        Response.Write "<div class=""empty"">" & Icon("music") & "<p>No hay canciones para mostrar.</p></div>"
        Exit Sub
    End If
    Response.Write "<div class=""tracklist" & IIfStr(showAlbum, "", " no-album") & """ data-tracklist>"
    Response.Write "<div class=""tl-head""><span class=""tl-idx"">#</span><span>Título</span>"
    If showAlbum Then Response.Write "<span class=""tl-album"">Álbum</span>"
    Response.Write "<span class=""tl-plays"" title=""Reproducciones"">" & Icon("headphones") & "</span></div>"
    i = 0
    For Each s In songs
        i = i + 1
        cover = coverOverride
        If cover = "" Then cover = Nz(s("COVER"), "")
        artists = Nz(s("ARTISTS"), "")
        If artists = "" Then artists = "Artista desconocido"
        Response.Write "<div class=""tl-row"" data-track data-id=""" & s("ID_SONG") & """ data-src=""" & H(MediaUrl(s("SONG_PATH"))) & """ data-title=""" & H(s("SONG_NAME")) & """ data-artist=""" & H(artists) & """ data-cover=""" & H(MediaUrl(cover)) & """>"
        Response.Write "<span class=""tl-idx""><span class=""tl-n"">" & i & "</span>"
        Response.Write "<button type=""button"" class=""tl-play"" data-play-track aria-label=""Reproducir " & H(s("SONG_NAME")) & """>" & Icon("play") & "</button>"
        Response.Write "<span class=""eq"" aria-hidden=""true""><i></i><i></i><i></i></span></span>"
        Response.Write "<span class=""tl-main"">" & Img(cover, "cover", "", "tl-cover") & "<span class=""tl-text""><strong>" & H(s("SONG_NAME")) & "</strong><small>" & H(artists) & "</small></span></span>"
        If showAlbum Then
            If Nz(s("ID_ALBUM"), "") <> "" Then
                Response.Write "<span class=""tl-album""><a href=""album.asp?id=" & s("ID_ALBUM") & """>" & H(s("ALBUM_TITLE")) & "</a></span>"
            Else
                Response.Write "<span class=""tl-album muted"">Sencillo</span>"
            End If
        End If
        Response.Write "<span class=""tl-plays"">" & FmtNum(s("PLAYS")) & "</span></div>"
    Next
    Response.Write "</div>"
End Sub

Sub RenderCard(href, imgPath, kind, title, subtitle)
    Response.Write "<a class=""card" & IIfStr(kind = "artist", " card-round", "") & """ href=""" & H(href) & """>"
    Response.Write "<div class=""card-media"">" & Img(imgPath, kind, title, "") & "<span class=""card-fab"" aria-hidden=""true"">" & Icon("play") & "</span></div>"
    Response.Write "<div class=""card-title"">" & H(title) & "</div><div class=""card-sub"">" & H(subtitle) & "</div></a>"
End Sub

Sub RenderEmpty(iconName, msg)
    Response.Write "<div class=""empty"">" & Icon(iconName) & "<p>" & H(msg) & "</p></div>"
End Sub

' Lista de casillas con filtro, para relaciones N:M en formularios de administración.
' rows deben exponer columnas ID, LABEL, SUB.
Sub RenderChecklist(name, rows, selected, emptyMsg)
    Dim r, chk
    If UBound(rows) < 0 Then
        Response.Write "<p class=""muted"">" & H(emptyMsg) & "</p>"
        Exit Sub
    End If
    Response.Write "<div class=""checklist"" data-checklist>"
    Response.Write "<div class=""checklist-tools""><input type=""search"" class=""input input-sm"" placeholder=""Filtrar…"" data-checklist-filter aria-label=""Filtrar lista""><span class=""muted small"" data-checklist-count></span></div>"
    Response.Write "<div class=""checklist-items"">"
    For Each r In rows
        chk = ""
        If selected.Exists(CStr(r("ID"))) Then chk = " checked"
        Response.Write "<label class=""check-item""><input type=""checkbox"" name=""" & name & """ value=""" & r("ID") & """" & chk & "><span><strong>" & H(r("LABEL")) & "</strong>"
        If Nz(r("SUB"), "") <> "" Then Response.Write "<small>" & H(r("SUB")) & "</small>"
        Response.Write "</span></label>"
    Next
    Response.Write "</div></div>"
End Sub

Function IdSet(rows, field)
    Dim d, r
    Set d = Server.CreateObject("Scripting.Dictionary")
    For Each r In rows
        d(CStr(r(field))) = True
    Next
    Set IdSet = d
End Function

Function IIfStr(cond, a, b)
    If cond Then
        IIfStr = a
    Else
        IIfStr = b
    End If
End Function

Function RoleBadge(role)
    If CLng(role) = ROLE_ADMIN Then
        RoleBadge = "<span class=""badge badge-primary"">" & Icon("shield") & "Admin</span>"
    Else
        RoleBadge = "<span class=""badge"">Usuario</span>"
    End If
End Function

Function PlaylistTypeBadge(t)
    If CLng(t) = PLAYLIST_PRIVATE Then
        PlaylistTypeBadge = "<span class=""badge"">" & Icon("lock") & "Privada</span>"
    Else
        PlaylistTypeBadge = "<span class=""badge badge-success"">" & Icon("globe") & "Pública</span>"
    End If
End Function

' Cabecera estándar de páginas de administración
Sub RenderAdminHeader(title, subtitle, newHref, newLabel)
    Response.Write "<header class=""page-head""><div><p class=""eyebrow"">" & Icon("shield") & " Administración</p><h1>" & H(title) & "</h1>"
    If subtitle <> "" Then Response.Write "<p class=""muted"">" & H(subtitle) & "</p>"
    Response.Write "</div>"
    If newHref <> "" Then Response.Write "<a class=""btn btn-primary"" href=""" & newHref & """>" & Icon("plus") & H(newLabel) & "</a>"
    Response.Write "</header>"
End Sub

Sub RenderSearchBar(placeholder, q)
    Response.Write "<form class=""toolbar"" method=""get"" data-pjax-form><div class=""input-icon"">" & Icon("search") & "<input class=""input"" type=""search"" name=""q"" value=""" & H(q) & """ placeholder=""" & H(placeholder) & """ aria-label=""Buscar""></div><button class=""btn btn-ghost"" type=""submit"">Buscar</button></form>"
End Sub

' Formulario de eliminación (POST + CSRF + confirmación)
Function DeleteButton(id, what)
    DeleteButton = "<form method=""post"" class=""inline"" data-confirm=""¿Eliminar " & H(what) & "? Esta acción no se puede deshacer."">" & CsrfField() & _
        "<input type=""hidden"" name=""action"" value=""delete""><input type=""hidden"" name=""id"" value=""" & id & """>" & _
        "<button class=""icon-btn danger"" type=""submit"" title=""Eliminar"" aria-label=""Eliminar " & H(what) & """>" & Icon("trash") & "</button></form>"
End Function

Function EditButton(href, what)
    EditButton = "<a class=""icon-btn"" href=""" & href & """ title=""Editar"" aria-label=""Editar " & H(what) & """>" & Icon("edit") & "</a>"
End Function
%>
