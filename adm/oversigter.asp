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

<body OnLoad="setVariables();checkLocation()">
<!--#include file="menu.asp" -->
<br>
<br>
<br>
<br>
<table cellpadding="5" cellspacing="1" style="background-color:#aaaaaa" border="0">
<tr>
	<td bgcolor="#ffffff">
	<strong>T-shirts</strong>
	<br>
	<%
	Set rsQuery = Con.Execute("SELECT tblSkiBruger.tshirtStr, Count(tblSkiBruger.id) AS AntalTS FROM tblSkiBruger INNER JOIN Vw_SkiAntalTimer ON (tblSkiBruger.pw = Vw_SkiAntalTimer.pw) AND (tblSkiBruger.email = Vw_SkiAntalTimer.email) GROUP BY tblSkiBruger.tshirtStr")
 	%>
	<%
 	while rsQuery.eof = false
  	%>
 	<%=rsQuery("tshirtstr") %>:&nbsp;<%=rsQuery("AntalTS") %><br>
	<%
	rsquery.movenext
	wend
	%>
	</td>
	
	<td bgcolor="#ffffff" valign="top">
<%
Set rsQuery = Con.Execute("SELECT tblSkiBruger.camping, Count(tblSkiBruger.id) AS AntalTelt FROM tblSkiBruger GROUP BY tblSkiBruger.camping HAVING (((tblSkiBruger.camping)='Ja'));")
 %>
 <%
 while rsQuery.eof = false
  %>
 
	<strong>Antal tilmeldt camping</strong>:&nbsp;<%=rsQuery("AntalTelt") %><br>

<%
rsquery.movenext
wend
 %>
	
	
	</td>
</tr>
</table>
<br>
<hr>
<br>
<strong>For at lave txt-fil</strong><br>
Klik på dette link: <a href="csv.asp" target="_blank">Åbn txt-fil</a>, hvorefter I markerer hele teksten (Ctrl + a, efterfulgt af Ctrl + c), og sætter det ind i en txt-fil.<br>

<br>
<hr>
<br>

<strong>Udskrift af vagtplaner til afkrydsning</strong><br>
<br>
<form action="visUdskrift.asp" method="get">
<%
Set rsQuery = Con.Execute("SELECT tblSkiPerioder.id, tblSkiPerioder.dato, tblSkiPerioder.tid FROM tblSkiPerioder ORDER BY tblSkiPerioder.sortorder;")
 %>

<select name="periode">
<%
 while rsQuery.eof = false
%>
 
<option value="<%=rsQuery("id") %>"><%=rsQuery("dato") %>&nbsp;&nbsp;&nbsp;<%=rsQuery("tid") %></option> 


<%
rsquery.movenext
wend
 %>
</select>&nbsp;&nbsp;<input type="submit" value="Vis">
</form>
<br>
<hr>
<br>
<strong>Angiv status</strong><br>
<br>

<!--#include file="conn.asp" -->
<%Set rsQuery = Con.Execute("SELECT * FROM tblSkiTid") %>
<%'=rsQuery("tid") %>

<form action="updateStatus.asp" method="get">
<select name="tidStatus">
<option value="1" <%if rsQuery("tid") = "1" then response.write "selected" end if %>>Åben for tilmelding</option>
<option value="2" <%if rsQuery("tid") = "2" then response.write "selected" end if %>>Midlertidig lukket</option>
<option value="3" <%if rsQuery("tid") = "3" then response.write "selected" end if %>>Stop for tilmelding</option>
<option value="4" <%if rsQuery("tid") = "4" then response.write "selected" end if %>>Resultat vises</option>
<option value="5" <%if rsQuery("tid") = "5" then response.write "selected" end if %>>Lukket totalt</option>



<!-- <option value="2">Vælg (= lukket for tilmelding)</option>
<option value="4" <%if rsQuery("tid") = "4" then response.write "selected" end if %>>Midlertidig lukket</option>
<option value="1" <%if rsQuery("tid") = "1" then response.write "selected" end if %>>Åben for tilmelding</option>
<option value="2" <%if rsQuery("tid") = "2" then response.write "selected" end if %>>Lukket for tilmelding</option>
<option value="3" <%if rsQuery("tid") = "3" then response.write "selected" end if %>>Resultat vises</option>
 --></select>
<input type="submit" value="Sæt ny status">
</form>
<br>
<hr>
<br>
<strong>Datoer for validering</strong>
<br>
Bemærk at dato formatet <strong>SKAL</strong> være dd-mm-åååå
<%Set rsQuery = Con.Execute("SELECT * FROM tblSkiData")%>

<form action="updateData.asp" method="get">
18 år: <input type="text" name="18aar" value="<%=rsQuery("18aar") %>">
<br>
15 år: <input type="text" name="15aar" value="<%=rsQuery("15aar") %>">
<br>

<input type="submit" value="Gem nye datoer">
</form>
</body>
</html>
