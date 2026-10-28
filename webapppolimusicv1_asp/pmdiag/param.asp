<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Response.ContentType = "text/plain"
Response.Buffer = True
%>
<!--#include virtual="/includes/bootstrap.asp" -->
<%
Dim cn, cmd, rs, p
Response.Write "provider probado:" & vbCrLf
Response.Write "  MSOLEDBSQL -> " & ConnectionString() & vbCrLf & vbCrLf

Set cn = OpenConnection()
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = cn
cmd.CommandType = adCmdText
cmd.CommandText = "SELECT ID_USER, USERNAME, PASSWORD, USER_TYPE FROM TBL_USER WHERE USERNAME = ?"

Set p = cmd.CreateParameter("u", adVarWChar, adParamInput, 50)
p.Value = "userAdmin"
cmd.Parameters.Append p

Response.Write "parametros: " & cmd.Parameters.Count & vbCrLf
Response.Write "p.Value=[" & cmd.Parameters(0).Value & "] type=" & cmd.Parameters(0).Type & " size=" & cmd.Parameters(0).Size & vbCrLf

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = adUseClient
On Error Resume Next
rs.Open cmd, , adOpenStatic, adLockReadOnly
Response.Write "open err=" & Err.Number & " " & Err.Description & vbCrLf
If Err.Number = 0 Then
  Response.Write "recordcount check..." & vbCrLf
  If rs.EOF Then
    Response.Write "EOF -> ninguna fila" & vbCrLf
  Else
    Response.Write "fila: " & rs("USERNAME") & " tipo=" & rs("USER_TYPE") & vbCrLf
  End If
End If
Err.Clear
On Error GoTo 0

' Prueba 2: LIKE de control, sin parametro
Response.Write vbCrLf & "control sin parametro:" & vbCrLf
Set rs = ExecuteQuery("SELECT COUNT(*) AS N FROM TBL_USER WHERE USERNAME = 'userAdmin'", Empty)
If Not rs.EOF Then Response.Write "  N=" & rs("N") & vbCrLf

' Prueba 3: parametro entero
Set rs = ExecuteQuery("SELECT USERNAME FROM TBL_USER WHERE ID_USER = ?", Array(2))
If Not rs.EOF Then Response.Write "  int param OK: " & rs("USERNAME") & vbCrLf Else Response.Write "  int param EOF" & vbCrLf

CloseConnection cn
%>
