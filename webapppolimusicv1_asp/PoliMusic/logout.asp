<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
' Cierre de sesión solo por POST con token CSRF (evita cierres forzados vía enlaces externos)
If IsPost() Then
    If CsrfValid() Then
        SignOut
        RedirectTo "login.asp?bye=1"
    End If
End If
If IsLoggedIn() Then
    RedirectTo "default.asp"
Else
    RedirectTo "login.asp"
End If
%>
