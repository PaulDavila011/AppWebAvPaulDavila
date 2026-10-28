<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Response.ContentType = "text/plain"
Response.Buffer = True
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
' Llamada a la funcion interna tal cual
Dim rs
Response.Write "A) ExecuteQueryTyped con Empty como 3er arg" & vbCrLf
Set rs = ExecuteQueryTyped("SELECT USERNAME FROM TBL_USER WHERE USERNAME = ?", Array("userAdmin"), Empty)
Response.Write "   TypeName(rs)=" & TypeName(rs) & " IsObject=" & IsObject(rs) & vbCrLf
If IsObject(rs) Then
  If rs Is Nothing Then Response.Write "   rs is Nothing" & vbCrLf Else Response.Write "   rs no es Nothing" & vbCrLf
End If

Response.Write vbCrLf & "B) ExecuteQuery" & vbCrLf
On Error Resume Next
Set rs = ExecuteQuery("SELECT USERNAME FROM TBL_USER WHERE USERNAME = ?", Array("userAdmin"))
Response.Write "   err=" & Err.Number & " " & Err.Description & vbCrLf
Response.Write "   TypeName=" & TypeName(rs) & " IsObject=" & IsObject(rs) & vbCrLf
Err.Clear
On Error GoTo 0
%>
