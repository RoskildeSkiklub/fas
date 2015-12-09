<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>

<!--#include file="conn.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
	<title></title>
<LINK HREF="adm.css" REL="stylesheet" TYPE="text/css">
</head>

<body>
<%

Set rsQuery = Con.Execute("SELECT tblSkiPerioder.dato, tblSkiPerioder.tid, tblSkiBruger.navn, tblSkiBruger.arbopg, tblSkiBruger.telefon, tblSkiBruger.Efternavn, tblSkiPerioder.id, tblSkiBrugerTider.vagtled FROM (tblSkiBrugerTider RIGHT JOIN tblSkiPerioder ON tblSkiBrugerTider.PeriodeID = tblSkiPerioder.id) LEFT JOIN tblSkiBruger ON tblSkiBrugerTider.pw = tblSkiBruger.pw AND tblSkiBrugerTider.email = tblSkiBruger.email WHERE (((tblSkiPerioder.id)=" & request("periode")  & ")) ORDER BY tblSkiPerioder.sortorder asc, tblSkiBrugerTider.vagtled desc, tblSkiBruger.navn asc")

'Set rsQuery = Con.Execute("SELECT tblSkiPerioder.dato, tblSkiPerioder.tid, tblSkiBruger.navn, tblSkiBruger.Efternavn, tblSkiPerioder.id FROM (tblSkiBrugerTider RIGHT JOIN tblSkiPerioder ON tblSkiBrugerTider.PeriodeID = tblSkiPerioder.id) LEFT JOIN tblSkiBruger ON tblSkiBrugerTider.pw = tblSkiBruger.pw WHERE (((tblSkiPerioder.id)=" & request("periode")  & ")) ORDER BY tblSkiPerioder.sortorder, tblSkiBruger.navn ")
 %>

<h1>
<%=weekdayname(weekday(rsQuery("dato"))) %>&nbsp;<%=rsQuery("dato") %>: <%=rsQuery("tid") %>
</h1>

<table cellpadding="5" style="background-color:#000000" border="1" cellspacing="1" width="100%">
<tr>
<td style="background-color:#ffffff" width="1%">Navn</td>
<td style="background-color:#ffffff" width="1%" align="center">Telefon</td>
<td align="center" style="background-color:#ffffff" width="1%">Arbejdsopgave</td>
<td align="center" style="background-color:#ffffff">Kommentarer</td>
</tr>
<%
 while rsQuery.eof = false
%>
<tr>
<td style="font-size:10pt;background-color:#ffffff;white-space:nowrap">
<%=rsQuery("navn") %>&nbsp;<%=rsQuery("efternavn") %>
<%if rsQuery("vagtled") = "1" then %>
<br>
<span style="color:red;font-weight:bold">Vagtleder</span>
<%end if %>
</td>
<td style="background-color:#ffffff">
<%=rsQuery("telefon") %>
</td>
<td style="background-color:#ffffff">
<%=rsQuery("arbopg") %>
</td>
<td style="background-color:#ffffff">&nbsp;</td>
</tr>

<%
rsquery.movenext
wend
 %>

</table>
<br>
<br>
<a href="oversigter.asp">Tilbage</a>
</form>
</body>
</html>
