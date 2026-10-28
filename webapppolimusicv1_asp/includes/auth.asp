<%
'==============================================================================
' PoliMusic - Autenticacion, sesiones y control de acceso basado en roles
'==============================================================================
' Las contrasenas se almacenan y comparan en texto plano, tal como estan en la
' base de datos heredada (TBL_USER.PASSWORD varchar(50)). Es una decision
' explicita del proyecto para mantener compatibilidad con los datos existentes.
'==============================================================================

Const SESSION_USER_ID    = "pm_uid"
Const SESSION_USER_NAME  = "pm_uname"
Const SESSION_USER_TYPE  = "pm_utype"
Const SESSION_USER_MAIL  = "pm_umail"
Const SESSION_USER_PIC   = "pm_upic"
Const SESSION_LAST_SEEN  = "pm_seen"

'==============================================================================
' IsLoggedIn() - True si hay una sesion de usuario activa.
'==============================================================================
Function IsLoggedIn()
    Dim uid
    uid = Session(SESSION_USER_ID)
    If IsNull(uid) Or IsEmpty(uid) Then
        IsLoggedIn = False
    Else
        IsLoggedIn = (CLng(uid) > 0)
    End If
End Function

'==============================================================================
' Accessors de la sesion actual
'==============================================================================
Function CurrentUserId()
    If Not IsLoggedIn() Then
        CurrentUserId = 0
    Else
        CurrentUserId = CLng(Session(SESSION_USER_ID))
    End If
End Function

Function CurrentUserName()
    If Not IsLoggedIn() Then
        CurrentUserName = ""
    Else
        CurrentUserName = CStr(Session(SESSION_USER_NAME) & "")
    End If
End Function

Function CurrentUserEmail()
    If Not IsLoggedIn() Then
        CurrentUserEmail = ""
    Else
        CurrentUserEmail = CStr(Session(SESSION_USER_MAIL) & "")
    End If
End Function

Function CurrentUserPhoto()
    If Not IsLoggedIn() Then
        CurrentUserPhoto = ""
    Else
        CurrentUserPhoto = CStr(Session(SESSION_USER_PIC) & "")
    End If
End Function

'==============================================================================
' CurrentUserType() - ROLE_USER (0) o ROLE_ADMIN (1).
'==============================================================================
Function CurrentUserType()
    If Not IsLoggedIn() Then
        CurrentUserType = ROLE_USER
    Else
        CurrentUserType = CLng(Session(SESSION_USER_TYPE))
    End If
End Function

'==============================================================================
' IsAdmin() - True si el usuario de la sesion tiene rol de administrador.
'==============================================================================
Function IsAdmin()
    IsAdmin = (IsLoggedIn() And CurrentUserType() = ROLE_ADMIN)
End Function

'==============================================================================
' RoleLabel() - Texto legible del rol.
'==============================================================================
Function RoleLabel()
    If IsAdmin() Then
        RoleLabel = "Administrador"
    Else
        RoleLabel = "Usuario"
    End If
End Function

'==============================================================================
' UserInitials() - Iniciales para el avatar de respaldo.
'==============================================================================
Function UserInitials(userName)
    Dim s, parts, i, out
    s = Trim(CStr(userName & ""))
    If Len(s) = 0 Then
        UserInitials = "?"
        Exit Function
    End If

    parts = Split(s, ".")
    out = ""
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then out = out & UCase(Left(parts(i), 1))
        If Len(out) >= 2 Then Exit For
    Next

    If Len(out) = 0 Then out = UCase(Left(s, 1))
    UserInitials = out
End Function

'==============================================================================
' PasswordMatches() - Compara contrasena en texto plano contra lo almacenado.
' Se usa comparacion exacta (binaria) para no depender del collation.
'==============================================================================
Function PasswordMatches(plain, stored)
    Dim s, p
    s = CStr(stored & "")
    p = CStr(plain & "")

    PasswordMatches = False
    If Len(s) = 0 Or Len(p) = 0 Then Exit Function

    If StrComp(s, p, 0) = 0 Then
        PasswordMatches = True
    End If
End Function

'==============================================================================
' Authenticate() - Valida credenciales contra TBL_USER.
' Devuelve el ID_USER si son correctas, 0 en caso contrario.
'==============================================================================
Function Authenticate(userName, plainPassword)
    Dim rs, uid
    Authenticate = 0

    Set rs = ExecuteQuery( _
        "SELECT ID_USER, USERNAME, PASSWORD, EMAIL, USER_TYPE, USER_PHOTO " & _
        "FROM TBL_USER WHERE USERNAME = ?", _
        Array(CStr(userName)))

    If Not rs.EOF Then
        If PasswordMatches(plainPassword, rs("PASSWORD")) Then
            uid = CLng(rs("ID_USER"))

            ' Establece la sesion
            Session(SESSION_USER_ID)   = uid
            Session(SESSION_USER_NAME) = CStr(rs("USERNAME") & "")
            Session(SESSION_USER_TYPE) = CLng(rs("USER_TYPE"))
            Session(SESSION_USER_MAIL) = CStr(rs("EMAIL") & "")
            Session(SESSION_USER_PIC)  = CStr(rs("USER_PHOTO") & "")
            Session(SESSION_LAST_SEEN) = Now()

            ' Renueva el token CSRF al iniciar sesion (evita fijacion de sesion)
            Session("pm_csrf") = NewToken()

            Authenticate = uid
        End If
    End If

    rs.Close
    Set rs = Nothing
End Function

'==============================================================================
' Logout() - Destruye la sesion por completo.
'==============================================================================
Sub Logout()
    Session.Abandon
End Sub

'==============================================================================
' RequireLogin() - Exige sesion activa; si no la hay, redirige a login.asp
' guardando el destino para volver despues de autenticarse.
'==============================================================================
Sub RequireLogin()
    If Not IsLoggedIn() Then
        Dim target
        target = CStr(Request.ServerVariables("URL") & "")
        If Len(Request.ServerVariables("QUERY_STRING") & "") > 0 Then
            target = target & "?" & Request.ServerVariables("QUERY_STRING")
        End If
        Session("pm_return_to") = target
        Response.Redirect "login.asp"
        Response.End
    End If
End Sub

'==============================================================================
' RequireAdmin() - Exige rol de administrador. Si hay sesion pero no permiso,
' devuelve 403; si no hay sesion, manda al login.
'==============================================================================
Sub RequireAdmin()
    If Not IsLoggedIn() Then
        RequireLogin()
    End If
    If Not IsAdmin() Then
        DenyAccess "Esta seccion esta reservada a los administradores. " & _
                   "Tu cuenta no tiene permisos para gestionar el catalogo."
    End If
End Sub

'==============================================================================
' TouchSession() - Refresca la marca de actividad (control de expiracion).
'==============================================================================
Sub TouchSession()
    If IsLoggedIn() Then
        Session(SESSION_LAST_SEEN) = Now()
    End If
End Sub

'==============================================================================
' RefreshSessionUser() - Vuelve a leer el usuario desde la BD para reflejar
' cambios de perfil (nombre, correo, foto) sin cerrar sesion.
'==============================================================================
Sub RefreshSessionUser()
    Dim rs
    If Not IsLoggedIn() Then Exit Sub

    Set rs = ExecuteQuery( _
        "SELECT USERNAME, EMAIL, USER_TYPE, USER_PHOTO FROM TBL_USER WHERE ID_USER = ?", _
        Array(CurrentUserId()))

    If Not rs.EOF Then
        Session(SESSION_USER_NAME) = CStr(rs("USERNAME") & "")
        Session(SESSION_USER_TYPE) = CLng(rs("USER_TYPE"))
        Session(SESSION_USER_MAIL) = CStr(rs("EMAIL") & "")
        Session(SESSION_USER_PIC)  = CStr(rs("USER_PHOTO") & "")
    End If

    rs.Close
    Set rs = Nothing
End Sub
%>
