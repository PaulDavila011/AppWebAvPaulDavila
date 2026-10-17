<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/core.asp"-->
<%
' Endpoint AJAX: registra una reproducción (TBL_SONG.PLAYS + 1)
RequireLoginApi
Response.ContentType = "application/json"

Dim id, plays
If Not IsPost() Or Not CsrfValid() Then
    Response.Status = "400 Bad Request"
    Response.Write "{""ok"":false,""error"":""bad_request""}"
Else
    id = ToInt(FormValue("id"), 0)
    plays = QScalar("SET NOCOUNT ON; UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS,0) + 1 WHERE ID_SONG = ?; SELECT PLAYS FROM TBL_SONG WHERE ID_SONG = ?;", Array(id, id))
    If IsNull(plays) Then
        Response.Status = "404 Not Found"
        Response.Write "{""ok"":false,""error"":""not_found""}"
    Else
        Response.Write "{""ok"":true,""id"":" & id & ",""plays"":" & CLng(plays) & "}"
    End If
End If
DbClose
%>
