
<!--#include file="conn.asp" -->

<%
Set rsQuery = Con.Execute("SELECT * FROM tblSkiBruger WHERE email like '" & request("email") & "' and pw like " & request("pw") & "")

if request("pw") = "45777530" then
session("adm") = "1"
end if

if rsQuery.eof = false then
session("pw") = rsQuery("pw")
session("email") = rsQuery("email")

if request("tid") = 3 then
response.redirect "result.asp"
else
response.redirect "profil.asp"
end if

else
session("succes") = 4
response.redirect "index.asp"
end if


Set Con = Nothing
Set rsQuery   = Nothing

'56408110

 %>