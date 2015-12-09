<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>
<!--#include file="conn.asp" -->


<%

if request("tidStatus") <> "" then

Set rsQuery1 = Con.Execute("UPDATE tblSkiTid SET tid = '" & request("tidStatus") & "' WHERE id = 1" )
Set rsQuery   = Nothing

end if



response.redirect "oversigter.asp"

 %>