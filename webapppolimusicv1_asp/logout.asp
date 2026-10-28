<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
'==============================================================================
' PoliMusic - Cierre de sesion
'==============================================================================
' Destruye la sesion por completo y vuelve al login con un mensaje.
'==============================================================================
Logout()
Response.Redirect "login.asp"
Response.End
%>
