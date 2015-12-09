<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>
<!--#include file="conn.asp" -->

<%

Set rsQuery = Con.Execute("UPDATE tblSkiData SET [18aar] = '" & request("18aar") & "', [15aar] = '" & request("15aar") & "'")

Set Con = Nothing
Set rsQuery   = Nothing

response.redirect "oversigter.asp"


 %>