<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Response.ContentType = "text/plain"
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
Dim errors, i
errors = 0

Response.Write "=== PoliMusic autotest ===" & vbCrLf

' --- 1. Conexion -------------------------------------------------------------
On Error Resume Next
Dim cn
Set cn = OpenConnection()
If Err.Number <> 0 Then
    Response.Write "FAIL conexion: " & Err.Description & vbCrLf
    errors = errors + 1
Else
    Response.Write "OK   conexion a " & DB_NAME & vbCrLf
End If
Err.Clear
On Error GoTo 0

' --- 2. Consulta simple ------------------------------------------------------
On Error Resume Next
Dim rs
Set rs = ExecuteQuery("SELECT COUNT(*) AS N FROM TBL_SONG", Empty)
If Err.Number <> 0 Then
    Response.Write "FAIL ExecuteQuery: " & Err.Description & vbCrLf
    errors = errors + 1
Else
    Response.Write "OK   ExecuteQuery COUNT = " & rs("N") & vbCrLf
End If
Err.Clear
On Error GoTo 0

' --- 3. Consulta con parametro ----------------------------------------------
On Error Resume Next
Set rs = ExecuteQuery("SELECT USERNAME, PASSWORD, USER_TYPE FROM TBL_USER WHERE USERNAME = ?", Array("userAdmin"))
If Err.Number <> 0 Then
    Response.Write "FAIL parametro: " & Err.Description & vbCrLf
    errors = errors + 1
Else
    If rs.EOF Then
        Response.Write "FAIL userAdmin no encontrado" & vbCrLf
        errors = errors + 1
    Else
        Response.Write "OK   userAdmin tipo=" & rs("USER_TYPE") & " passLen=" & Len(rs("PASSWORD")) & vbCrLf
    End If
End If
Err.Clear
On Error GoTo 0

' --- 4. Authenticate ---------------------------------------------------------
On Error Resume Next
Dim uid
uid = Authenticate("userAdmin", "12345678")
If Err.Number <> 0 Then
    Response.Write "FAIL Authenticate: " & Err.Description & vbCrLf
    errors = errors + 1
Else
    Response.Write "OK   Authenticate userAdmin -> uid=" & uid & vbCrLf
End If
Err.Clear
On Error GoTo 0

' --- 5. Funciones del repositorio -------------------------------------------
On Error Resume Next
Set rs = SongsTop(5)
If Err.Number <> 0 Then
    Response.Write "FAIL SongsTop: " & Err.Description & vbCrLf
    errors = errors + 1
Else
    Dim n
    n = 0
    Do While Not rs.EOF
        n = n + 1
        rs.MoveNext
    Loop
    Response.Write "OK   SongsTop devolvio " & n & " canciones" & vbCrLf
End If
Err.Clear
On Error GoTo 0

On Error Resume Next
Set rs = AlbumsAll()
If Err.Number <> 0 Then
    Response.Write "FAIL AlbumsAll: " & Err.Description & vbCrLf
    errors = errors + 1
Else
    n = 0
    Do While Not rs.EOF
        n = n + 1
        rs.MoveNext
    Loop
    Response.Write "OK   AlbumsAll devolvio " & n & " albumes" & vbCrLf
End If
Err.Clear
On Error GoTo 0

On Error Resume Next
Set rs = ArtistsAll()
If Err.Number <> 0 Then
    Response.Write "FAIL ArtistsAll: " & Err.Description & vbCrLf
    errors = errors + 1
Else
    n = 0
    Do While Not rs.EOF
        n = n + 1
        rs.MoveNext
    Loop
    Response.Write "OK   ArtistsAll devolvio " & n & " artistas" & vbCrLf
End If
Err.Clear
On Error GoTo 0

On Error Resume Next
Set rs = PlaylistsVisible(23)
If Err.Number <> 0 Then
    Response.Write "FAIL PlaylistsVisible: " & Err.Description & vbCrLf
    errors = errors + 1
Else
    n = 0
    Do While Not rs.EOF
        n = n + 1
        rs.MoveNext
    Loop
    Response.Write "OK   PlaylistsVisible devolvio " & n & " playlists" & vbCrLf
End If
Err.Clear
On Error GoTo 0

' --- 6. Helpers de presentacion ---------------------------------------------
On Error Resume Next
Dim t
t = MediaUrl("~/songFiles/bensound-adventure.mp3", "")
Response.Write "OK   MediaUrl = " & t & vbCrLf
t = UserPhotoUrl("~/userImages/root.png")
Response.Write "OK   UserPhotoUrl = " & t & vbCrLf
t = CoverUrl("~/albumImages/album1.png")
Response.Write "OK   CoverUrl = " & t & vbCrLf
t = UserInitials("paul.davila")
Response.Write "OK   UserInitials = " & t & vbCrLf
t = H("<b>x</b>")
Response.Write "OK   H() = " & t & vbCrLf
If Err.Number <> 0 Then
    Response.Write "FAIL helpers: " & Err.Description & vbCrLf
    errors = errors + 1
End If
Err.Clear
On Error GoTo 0

' --- 7. Iconos ---------------------------------------------------------------
On Error Resume Next
Dim ic
ic = Icon("home")
If Err.Number <> 0 Then
    Response.Write "FAIL Icon: " & Err.Description & vbCrLf
    errors = errors + 1
Else
    Response.Write "OK   Icon(home) len=" & Len(ic) & vbCrLf
End If
ic = Icon("play")
ic = Icon("album")
ic = Icon("artist")
ic = Icon("playlist")
ic = Icon("user")
ic = Icon("dashboard")
If Err.Number <> 0 Then
    Response.Write "FAIL Icon varios: " & Err.Description & vbCrLf
    errors = errors + 1
Else
    Response.Write "OK   Iconos varios generados" & vbCrLf
End If
Err.Clear
On Error GoTo 0

Response.Write vbCrLf & "=== RESUMEN: " & errors & " fallos ===" & vbCrLf
%>
