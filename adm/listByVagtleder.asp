<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>


<!--#include file="conn.asp" -->



<HTML xml:lang="dk" xmlns="http://www.w3.org/1999/xhtml">

<html>
<head>
	<title></title>
<LINK HREF="adm.css" REL="stylesheet" TYPE="text/css">
</head>

<body OnLoad="setVariables();checkLocation()">
<!--#include file="menu.asp" -->
<br>
<br>
<br>
<br>
Ved at køre musen hen over et navn, vises personen data.<br>
Ved klik på navn linkes til personen profil, som så herefter kan revideres.
<br>
<br>
<%


Set rsQuery = Con.Execute("SELECT tblSkiPerioder.dato FROM tblSkiPerioder GROUP BY tblSkiPerioder.dato ORDER BY tblSkiPerioder.dato")
while not rsQuery.eof = true

' datoer puttes i array
strArr = strArr & rsQuery("dato") & ";"

rsQuery.movenext
wend
'response.write strArr

'Array splittes for at kunne lave loop
arrDates = split(strArr,";")

'response.write ubound(arrDates)

strEnd = ubound(arrDates) - 1
'strStart = cdate(arrDates(0))

'response.write strStart
'response.write strEnd

for i = 0 to strEnd

'response.write arrDates(i)

'sql = "SELECT tblBruger.fdato, tblBruger.email, tblBruger.pw, tblBruger.navn, tblBruger.inact, tblBruger.Efternavn, tblBruger.vagtleder, tblPerioder.dato, tblPerioder.tid, tblPerioder.sortorder, tblBrugerTider.vagtled, tblBrugerTider.id FROM (tblBruger INNER JOIN tblBrugerTider ON tblBruger.pw = tblBrugerTider.pw) INNER JOIN tblPerioder ON tblBrugerTider.PeriodeID = tblPerioder.id WHERE inact <> '1' or inact Is Null  ORDER BY tblPerioder.sortorder asc,  tblBrugerTider.vagtled desc, tblBruger.vagtleder asc, navn asc"

sql = "SELECT tblSkiBruger.fdato, tblSkiBruger.email, tblSkiBruger.pw, tblSkiBruger.navn, tblSkiBruger.inact, tblSkiBruger.Efternavn, tblSkiBruger.vagtleder, tblSkiBrugerTider.id, tblSkiBrugerTider.vagtled, tblSkiPerioder.sortorder, tblSkiPerioder.dato, tblSkiPerioder.tid FROM (tblSkiBrugerTider LEFT JOIN tblSkiBruger ON tblSkiBrugerTider.pw = tblSkiBruger.pw) RIGHT JOIN tblSkiPerioder ON tblSkiBrugerTider.PeriodeID = tblSkiPerioder.id WHERE (((tblSkiBruger.inact)<>'1' Or (tblSkiBruger.inact) Is Null) AND ((tblSkiPerioder.dato) like CONVERT(DATETIME, '" & right(arrDates(i),4) & "-" & mid(arrDates(i),4,2) & "-" & left(arrDates(i),2) & " 00:00:00', 102))) ORDER BY tblSkiPerioder.sortorder asc, tblSkiBrugerTider.vagtled desc, tblSkiBruger.vagtleder asc, tblSkiBruger.navn asc "


'SELECT qryUserPeriode.id, qryUserPeriode.minmax, qryUserPeriode.dato, qryUserPeriode.tid, tblBruger.fdato, tblBruger.email, tblBruger.pw, tblBruger.navn, tblBruger.inact, tblBruger.Efternavn, tblBruger.vagtleder, tblBrugerTider.vagtled, qryUserPeriode.sortorder, qryUserPeriode.max FROM tblBrugerTider RIGHT JOIN (qryUserPeriode LEFT JOIN tblBruger ON qryUserPeriode.pw = tblBruger.pw) ON tblBrugerTider.pw = tblBruger.pw WHERE (((qryUserPeriode.dato) Like '" & dateAdd("d",strStart,i) & "') AND ((tblBruger.inact)<>'1' Or (tblBruger.inact) Is Null)) ORDER BY qryUserPeriode.sortorder, tblBruger.navn"

'SELECT qryUserPeriode.id, qryUserPeriode.minmax, qryUserPeriode.dato, qryUserPeriode.tid, tblBruger.fdato, tblBruger.email, tblBruger.pw, tblBruger.navn, tblBruger.inact, tblBruger.Efternavn, tblBruger.vagtleder, tblBrugerTider.vagtled, qryUserPeriode.sortorder, qryUserPeriode.max FROM qryUserPeriode LEFT JOIN tblBruger ON qryUserPeriode.pw = tblBruger.pw WHERE qryUserPeriode.dato like '" & dateAdd("d",strStart,i) & "' and ( inact <> '1' or inact Is Null ) ORDER BY qryUserPeriode.sortorder, tblBruger.navn asc"

Set rsQuery = Con.Execute(sql)
%>
<br>
<br>
<strong><%=weekdayname(weekday(rsQuery("dato"))) %>&nbsp;<%=rsQuery("dato") %></strong>
<table cellpadding="5" border="1">

		<%'første række %>
<tr>
<%
while not rsQuery.eof = true
%>
<%
if strTid <> rsQuery("tid") then
%>
<td valign="top" style="white-space:nowrap" width="250px">
<strong><%=rsQuery("tid") %></strong>

</td>
<%
end if 
 %>
<%
strTid = rsQuery("tid")
rsQuery.movenext
wend
rsQuery.movefirst
%>
 </tr>


<%
' række med lister over deltagere 
%>
<tr>


<td valign="top">

<%
strTid = ""
while not rsQuery.eof = true
if strTid <> rsQuery("tid") and strTid <> "" then
%>
</td><td valign="top" style="white-space:nowrap">
<%
end if 

strTid = rsQuery("tid")
%>

<%
' hvis der er angivet id
if rsQuery("id") <> "" then 
 %>

<%if rsQuery("vagtled") = "1" then%>
<strong>Leder</strong>
<%else %>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%end if %>
<a href="updateVagtleder.asp?id=<%=rsQuery("id") %>&action=vagtled&status=<%=rsQuery("vagtled") %>" target="_self"><%=rsQuery("navn") %>&nbsp;<%=rsQuery("efternavn") %></a>
<%if rsQuery("vagtleder") = "Ja" then %>
<span style="color:red;font-weight:bold">:-)</span>
<%end if %>
&nbsp;<%response.write datediff("yyyy",rsQuery("fdato"),date()) %>&nbsp;år
&nbsp;<a href="profil-adm.asp?pw=<%=rsQuery("pw") %>&email=<%=rsQuery("email") %>">Profil</a>
<br>
		<br>

<%end if %>
		
<%
rsQuery.movenext
k = k + 1
wend


 %>
</td>
 </tr>
 </table>
<%
next 
%>
</form>
<br>
<br>


</body>
</html>
<%
Set Con = Nothing
Set rsQuery   = Nothing
 %>
 
