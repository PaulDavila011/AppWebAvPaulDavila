<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Response.ContentType = "text/plain"
Response.Write "inicio" & vbCrLf
Response.Flush
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
Response.Write "bootstrap cargado" & vbCrLf
Response.Flush
On Error Resume Next
Dim cn
Set cn = OpenConnection()
Response.Write "conn err: " & Err.Number & " " & Err.Description & vbCrLf
Err.Clear
Dim rs
Set rs = ExecuteQuery("SELECT COUNT(*) AS N FROM TBL_SONG", Empty)
Response.Write "query err: " & Err.Number & " " & Err.Description & vbCrLf
If Err.Number = 0 And Not rs.EOF Then Response.Write "N=" & rs("N") & vbCrLf
Err.Clear
On Error GoTo 0
Response.Write "fin" & vbCrLf
%>
