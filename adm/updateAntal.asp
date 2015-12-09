<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>
<!--#include file="conn.asp" -->

<%

'Set rsQuery = Con.Execute("SELECT Count(tblPerioder.id) AS AntalOfid FROM tblPerioder")
'strAntal = rsQuery("AntalOfID")

'response.write strAntal

Set rsQuery   = Nothing

'response.write "UPDATE tblPerioder SET [max] = " & request("3") & " WHERE id = 3"
'"UPDATE tblPerioder SET [value] = " & request("1") & " WHERE id = 1" 

for i = 1 to 50 ' skal blot være mere end id for sidste post i Perioder

if request(i) = "" then
strI = 1
else
strI = request(i)
end if

sql = "UPDATE tblSkiPerioder SET [max] = " & strI & " WHERE id = " & i 

Set rsQuery = Con.Execute(sql)

'response.write (" (UPDATE tblPerioder SET [max] = " & request(i) & " WHERE id = " & i & ") ")

next



Set Con = Nothing
Set rsQuery   = Nothing

response.redirect "listall.asp"


 %>