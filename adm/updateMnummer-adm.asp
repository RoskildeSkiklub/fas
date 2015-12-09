<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>
 <!--#include file="conn.asp" -->




<%
Set rsQuery = Con.Execute("UPDATE tblSkiBruger SET mnummer='" & request("mnummer") & "', bnummer='" & request("bnummer") & "', vente='" & request("vente") & "'  WHERE email like '" & request("email") & "' and pw like '" & request("pw") & "'")


Set Con = Nothing
Set rsQuery   = Nothing

'session("tak") = 1

response.redirect "list.asp"


 %>
