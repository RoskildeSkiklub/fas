<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>
<!--#include file="conn.asp" -->


<%
strID = request("id")
strAct = request("action")

if request("status") <> "" then
strStatus = ""
else
strStatus = 1
end if


Set rsQuery1 = Con.Execute("UPDATE tblSkiBrugerTider SET [" & strAct & "] = '" & strStatus & "' WHERE id = " & request("id") )
'response.write "UPDATE tblBrugerTider SET [" & strAct & "] = '" & strStatus & "' WHERE id = " & request("id") 

Set rsQuery   = Nothing

response.redirect "listByVagtleder.asp"

 %>