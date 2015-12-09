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
Ved klik på navn linkes til personen profil, som så herefter kan revideres.
<br>
Personer under 18 markeres med grøn baggrund, alle andre med blå baggruund.
<br>
Det samlede maxantal for vagten vises med lilla/lyserød baggrund.
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

sql = "SELECT Vw_SkiUserPeriode.id, Vw_SkiUserPeriode.minmax, Vw_SkiUserPeriode.dato, Vw_SkiUserPeriode.tid, tblSkiBruger.id AS Uid, tblSkiBruger.navn, tblSkiBruger.Efternavn, tblSkiBruger.fdato, tblSkiBruger.adresse, tblSkiBruger.postnr, tblSkiBruger.bye, tblSkiBruger.email, tblSkiBruger.telefon, tblSkiBruger.vagtleder, tblSkiBruger.pw, tblSkiBruger.tshirtStr, Vw_SkiUserPeriode.sortorder, Vw_SkiUserPeriode.max FROM Vw_SkiUserPeriode LEFT JOIN tblSkiBruger ON Vw_SkiUserPeriode.pw = tblSkiBruger.pw WHERE Vw_SkiUserPeriode.dato like CONVERT(DATETIME, '" & right(arrDates(i),4) & "-" & mid(arrDates(i),4,2) & "-" & left(arrDates(i),2) & " 00:00:00', 102) and ( inact <> '1' or inact Is Null ) ORDER BY Vw_SkiUserPeriode.sortorder, tblSkiBruger.id asc"

'sql = "SELECT qryUserPeriode.id, qryUserPeriode.minmax, qryUserPeriode.dato, qryUserPeriode.tid, tblBruger.navn, tblBruger.fdato, tblBruger.Efternavn, tblBruger.adresse, tblBruger.postnr, tblBruger.bye, tblBruger.email, tblBruger.telefon, tblBruger.vagtleder, tblBruger.pw, tblBruger.[4x6], tblBruger.[3x8], tblBruger.[2x12], tblBruger.tshirtStr, qryUserPeriode.sortorder, qryUserPeriode.max FROM qryUserPeriode LEFT JOIN tblBruger ON qryUserPeriode.pw = tblBruger.pw WHERE qryUserPeriode.dato like '" & dateAdd("d",strStart,i) & "' and ( inact <> '1' or inact Is Null ) ORDER BY qryUserPeriode.sortorder asc, tblBruger.fdato asc"

'response.write sql

Set rsQuery = Con.Execute(sql)
%>


<%=rsQuery("Dato") %>
<table style="" border="0" cellpadding="0" cellspacing="1">
<tr>
	<td style="display:none">

	<%while not rsQuery.eof = true%>

	<%if strTid <> rsQuery("tid") then%>
		</td>
		<td style="font-size:10px;padding:3px" valign="bottom" align="center">
	<%end if %>
	
	
	
	<%if strTid <> rsQuery("tid") then%>
		<strong><%=rsQuery("tid") %></strong>
		<br>
		M=<%=rsQuery("max") %>
		<%end if %>

<%
'	strPeriode = rsQuery("dato")
'	strDag = rsQuery("dato")
	strTid = rsQuery("tid")
	rsQuery.movenext
	wend
	rsQuery.movefirst
%>

</tr>
<tr>



<%
while not rsQuery.eof = true
%>


<%if strTid <> rsQuery("tid") then%>
</div>
	</td>
	<td style="width:40px;" valign="top" align="center">


		<div style="height: <%=rsQuery("max")*12 %>; background-color: #ddaaee; z-index: auto; ">

<%end if %>



<%if rsQuery("fdato") <> "" or isnull(rsQuery("fdato")) = false then %>
	<div style="z-index:auto;vertical-align:bottom;height:12px;background-color:	<%if datediff("yyyy",rsQuery("fdato"),"01-07-2008") < 18 then %>	#C0F6C0;<%else %>#B0C4DE;<%end if %>font-size:9px;">
	<a href="profil-adm.asp?pw=<%=rsQuery("pw") %>&email=<%=rsQuery("email") %>" style="text-decoration:none">
	<%=datediff("yyyy",rsQuery("fdato"),"01-07-2008" )%></a>
	</div>
<%end if %>

<%
	strTid = rsQuery("tid")
	rsQuery.movenext
%>
<%	wend%>


</td>
</tr>
</table>

<br>
<br>

<%
next 
%>



<!-- ------------------------------------------------------------------------------------------ -->




</body>
</html>
<%
Set Con = Nothing
Set rsQuery   = Nothing
 %>
 
