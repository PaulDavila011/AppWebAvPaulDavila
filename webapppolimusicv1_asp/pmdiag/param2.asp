<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Response.ContentType = "text/plain"
Response.Buffer = True
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
Dim arr, i
arr = Array("userAdmin")
Response.Write "IsArray(arr)=" & IsArray(arr) & vbCrLf
Response.Write "UBound=" & UBound(arr) & vbCrLf
For i = 0 To UBound(arr)
  Response.Write "  arr(" & i & ") vartype=" & VarType(arr(i)) & " val=[" & arr(i) & "]" & vbCrLf
  Response.Write "  ResolveParamType=" & ResolveParamType(arr(i)) & vbCrLf
  Response.Write "  ResolveParamSize=" & ResolveParamSize(arr(i)) & vbCrLf
Next

Response.Write vbCrLf & "Ahora ExecuteQuery con array:" & vbCrLf
Dim rs
Set rs = ExecuteQuery("SELECT USERNAME FROM TBL_USER WHERE USERNAME = ?", arr)
If rs.EOF Then Response.Write "  EOF" & vbCrLf Else Response.Write "  OK " & rs("USERNAME") & vbCrLf
Response.Write "  rs.State=" & rs.State & vbCrLf

Response.Write vbCrLf & "ExecuteQuery con Array() inline:" & vbCrLf
Set rs = ExecuteQuery("SELECT USERNAME FROM TBL_USER WHERE USERNAME = ?", Array("userAdmin"))
If rs.EOF Then Response.Write "  EOF" & vbCrLf Else Response.Write "  OK " & rs("USERNAME") & vbCrLf
%>
