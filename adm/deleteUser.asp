<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>
<!--#include file="conn.asp" -->

<%

if request("delete") = "delete" then

Set rsQuery = Con.Execute("DELETE FROM tblSkiBruger WHERE email like '" & request("email") & "' and pw like '" & request("pw") & "'")

Set rsQuery = Con.Execute("DELETE FROM tblSkiBrugerTider WHERE email like '" & request("email") & "' and pw like '" & request("pw") & "'")

Set Con = Nothing
Set rsQuery   = Nothing

end if

response.redirect "list.asp"


 %>