<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
RequireAdmin
pageTitle = "Usuarios"
activeNav = "admin_users"

Dim q, users, u, id, errMsg, dummy, roleFilter, oldMedia
If IsPost() Then
    RequireCsrf
    If FormValue("action") = "delete" Then
        id = ToInt(FormValue("id"), 0)
        If id = CurrentUserId() Then
            SetFlash "error", "No puedes eliminar tu propia cuenta mientras estás conectado."
        Else
            ' Elimina primero sus playlists (y canciones asociadas) para respetar las FK
            oldMedia = QScalar("SELECT USER_PHOTO FROM TBL_USER WHERE ID_USER = ?", Array(id))
            errMsg = TryBatch("DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN (SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = ?); " & _
                              "DELETE FROM TBL_PLAYLIST WHERE ID_USER = ?; DELETE FROM TBL_USER WHERE ID_USER = ?;", Array(id, id, id), dummy)
            If errMsg = "" Then DeleteUploadedMedia oldMedia
            If errMsg = "" Then SetFlash "success", "Usuario eliminado." Else SetFlash "error", errMsg
        End If
    End If
    RedirectTo "admin_users.asp"
End If

q = QS("q")
roleFilter = ToInt(QS("role"), -1)
users = QList("SELECT u.ID_USER, u.USERNAME, u.EMAIL, u.BIRTHDAY, u.USER_PHOTO, u.USER_TYPE, (SELECT COUNT(*) FROM TBL_PLAYLIST p WHERE p.ID_USER = u.ID_USER) AS PLAYLISTS " & _
              "FROM TBL_USER u WHERE (? = '' OR u.USERNAME LIKE ? OR u.EMAIL LIKE ?) AND (? = -1 OR u.USER_TYPE = ?) ORDER BY u.USER_TYPE DESC, u.USERNAME", _
              Array(q, "%" & q & "%", "%" & q & "%", roleFilter, roleFilter))

Function RoleChip(v, label)
    RoleChip = "<a class=""chip" & IIfStr(roleFilter = v, " active", "") & """ href=""admin_users.asp?role=" & v & IIfStr(q <> "", "&amp;q=" & Server.URLEncode(q), "") & """>" & label & "</a>"
End Function
%>
<!--#include file="includes/layout_top.asp"-->
<% RenderAdminHeader "Usuarios", "Crea, edita y elimina cuentas. Asigna el rol de administrador con cuidado.", "admin_user_form.asp", "Nuevo usuario" %>
<div class="toolbar-row">
  <% RenderSearchBar "Buscar por usuario o correo", q %>
  <div class="chips"><%= RoleChip(-1, "Todos") %><%= RoleChip(ROLE_ADMIN, "Administradores") %><%= RoleChip(ROLE_USER, "Usuarios") %></div>
</div>
<% If UBound(users) < 0 Then
     RenderEmpty "users", "No hay usuarios que coincidan."
   Else %>
<div class="table-wrap">
<table class="table">
  <thead><tr><th>Usuario</th><th class="hide-sm">Correo</th><th class="hide-md">Nacimiento</th><th>Rol</th><th class="num hide-sm">Playlists</th><th class="actions"><span class="sr-only">Acciones</span></th></tr></thead>
  <tbody>
  <% For Each u In users %>
    <tr>
      <td><div class="cell-main"><span class="avatar"><%= Img(u("USER_PHOTO"), "user", "", "") %></span><div><strong><%= H(u("USERNAME")) %></strong><% If CLng(u("ID_USER")) = CurrentUserId() Then %> <span class="badge badge-outline">Tú</span><% End If %><small class="show-sm"><%= H(u("EMAIL")) %></small></div></div></td>
      <td class="hide-sm"><%= H(u("EMAIL")) %></td>
      <td class="hide-md"><%= FmtDate(u("BIRTHDAY")) %></td>
      <td><%= RoleBadge(u("USER_TYPE")) %></td>
      <td class="num hide-sm"><%= u("PLAYLISTS") %></td>
      <td class="actions">
        <%= EditButton("admin_user_form.asp?id=" & u("ID_USER"), u("USERNAME")) %>
        <% If CLng(u("ID_USER")) <> CurrentUserId() Then Response.Write DeleteButton(u("ID_USER"), "al usuario " & u("USERNAME") & IIfStr(CLng(u("PLAYLISTS")) > 0, " y sus " & u("PLAYLISTS") & " playlists", "")) %>
      </td>
    </tr>
  <% Next %>
  </tbody>
</table>
</div>
<% End If %>
<!--#include file="includes/layout_bottom.asp"-->
