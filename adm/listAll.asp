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

<form action="updateAntal.asp" method="get">
<script type="text/javascript">
function submenu(e) {
e = document.getElementById(e)
e.style.display = (e.style.display=="block")?"none":"block"
}
</script>

<input type="submit" value="Opdater maxantal"><br>
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
'strStart = arrDates(0)
'strStart = right(arrDates(0),4) & "-" & mid(arrDates(0),4,2) & "-" & left(arrDates(0),2)

'response.write strStart
'response.write strEnd

for i = 0 to strEnd

'response.write arrDates(i)

sql = "SELECT Vw_SkiUserPeriode.id, Vw_SkiUserPeriode.minmax, Vw_SkiUserPeriode.dato, Vw_SkiUserPeriode.tid, tblSkiBruger.id AS Uid, tblSkiBruger.navn, tblSkiBruger.Efternavn, tblSkiBruger.adresse, tblSkiBruger.postnr, tblSkiBruger.bye, tblSkiBruger.email, tblSkiBruger.telefon, tblSkiBruger.vagtleder, tblSkiBruger.pw, tblSkiBruger.email, tblSkiBruger.tshirtStr, Vw_SkiUserPeriode.sortorder, Vw_SkiUserPeriode.max FROM Vw_SkiUserPeriode LEFT JOIN tblSkiBruger ON Vw_SkiUserPeriode.pw = tblSkiBruger.pw and 	Vw_SkiUserPeriode.email = tblSkiBruger.email WHERE Vw_SkiUserPeriode.dato like CONVERT(DATETIME, '" & right(arrDates(i),4) & "-" & mid(arrDates(i),4,2) & "-" & left(arrDates(i),2) & " 00:00:00', 102) and ( inact <> '1' or inact Is Null ) ORDER BY Vw_SkiUserPeriode.sortorder, tblSkiBruger.id asc"

'response.write sql

'sql = "SELECT Vw_SkiUserPeriode.id, Vw_SkiUserPeriode.minmax, Vw_SkiUserPeriode.dato, Vw_SkiUserPeriode.tid, tblSkiBruger.id AS Uid, tblSkiBruger.navn, tblSkiBruger.Efternavn, tblSkiBruger.adresse, tblSkiBruger.postnr, tblSkiBruger.bye, tblSkiBruger.email, tblSkiBruger.telefon, tblSkiBruger.vagtleder, tblSkiBruger.pw, tblSkiBruger.tshirtStr, Vw_SkiUserPeriode.sortorder, Vw_SkiUserPeriode.max FROM Vw_SkiUserPeriode LEFT JOIN tblSkiBruger ON Vw_SkiUserPeriode.pw = tblSkiBruger.pw WHERE Vw_SkiUserPeriode.dato like '" & dateAdd("d",cdate(strStart),i) & "' and ( inact <> '1' or inact Is Null ) ORDER BY Vw_SkiUserPeriode.sortorder, tblSkiBruger.id asc"

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
<br>
Åbne=<input type="text" size="2" value="<%=rsquery("max")%>" name="<%=rsQuery("id") %>"><br>
Min/Max:&nbsp;<%=rsQuery("minmax") %>
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

<%if rsQuery("pw") <> "" then %>
<a href="profil-adm.asp?pw=<%=rsQuery("pw") %>&email=<%=rsQuery("email") %>" onmouseover="submenu('desct<%=k %>')"  onmouseout="submenu('desct<%=k %>')">
<%end if %>

<%if rsQuery("efternavn") <> "" then %>
<%=rsQuery("navn") %>&nbsp;<%=rsQuery("efternavn") %>
<%end if %>
</a>
<div id="desct<%=k %>" class="descr" style="display:none">
<%=rsQuery("navn") %>&nbsp;<%=rsQuery("efternavn") %><br>
<%=rsQuery("Adresse") %><br>
 <%=rsQuery("postnr") %>&nbsp;<%=rsQuery("bye") %><br>
 
		<%=rsQuery("telefon") %>
	<br>
			<%=rsQuery("email") %>
		<br>
		Vagtleder: <%=rsQuery("vagtleder") %>
		</div>
		<br>

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
 
