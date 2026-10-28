<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Response.ContentType = "text/plain"
Response.Buffer = True
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
On Error Resume Next
Dim rs
Response.Write "[1] consulta directa" & vbCrLf
Set rs = ExecuteQuery("SELECT USERNAME, PASSWORD, USER_TYPE FROM TBL_USER WHERE USERNAME = ?", Array("userAdmin"))
Response.Write "    err=" & Err.Number & " " & Err.Description & vbCrLf
If Err.Number = 0 Then
  If rs.EOF Then
    Response.Write "    EOF" & vbCrLf
  Else
    Response.Write "    user=" & rs("USERNAME") & " type=" & rs("USER_TYPE") & vbCrLf
  End If
End If
Err.Clear

Response.Write "[2] Authenticate" & vbCrLf
Dim uid
uid = Authenticate("userAdmin", "12345678")
Response.Write "    err=" & Err.Number & " " & Err.Description & vbCrLf
Response.Write "    uid=" & uid & vbCrLf
Err.Clear
On Error GoTo 0
Response.Write "[fin]" & vbCrLf
%>
