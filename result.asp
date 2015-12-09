<!--#include file="conn.asp" -->

<%
Set rsQuery1 = Con.Execute("SELECT * FROM tblSkiBruger WHERE email like '" & session("email") & "' and pw like " & session("pw") & "")

 %>
 
 <%Set rsQuery3 = Con.Execute("SELECT tblSkiBrugerTider.vagtled, tblSkiBrugerTider.email, tblSkiBrugerTider.pw, tblSkiPerioder.id FROM tblSkiBrugerTider INNER JOIN tblSkiPerioder ON tblSkiBrugerTider.PeriodeID = tblSkiPerioder.id WHERE (((tblSkiBrugerTider.vagtled)='1')) and email like '" & session("email") & "' and pw like " & session("pw") & "")
strVagtleder = ","
while not rsQuery3.eof = true
strVagtleder = strVagtleder & rsquery3("id") & ","
rsquery3.movenext
wend
'response.write strVagtleder
 %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
	<title></title>
<style>
body{
	font-size:12px
}
table{
	background-color:#aaaaaa
}
td{
	background-color: #ffffff
}
</style>
<script language="javascript" src="jsval.js"></script>
</head>



<body >

<div style="width:450px">
<img src="logo.gif" height="124" width="124" align="right">
<h1>Roskilde Skiklub</h1>
<h2>Roskilde Festival</h2>
</div>
<br>
<a href="logaf.asp">Log af</a>

<br>
<br>


<!-- grunddata -->
Navn: <%=rsQuery1("navn") %>&nbsp;<%=rsQuery1("efternavn") %>
<br>
Medlemsnummer:&nbsp;<%=rsQuery1("mnummer") %>
<br>
<br>

<% if len(rsQuery1("andrebesk")) > 2 then %>
<strong>Vær opmærksom på:</strong><br>
<%=rsQuery1("andrebesk") %>
<br>
<br><%end if %>

<!-- grunddata slut -->
<%

'Set Con = Nothing
'Set rsQuery   = Nothing

Set rsQuery = Con.Execute("SELECT tblSkiBrugerTider.PeriodeID FROM tblSkiBrugerTider WHERE tblSkiBrugerTider.pw like " &  session("pw") & "")

strPerioder = ","
while not rsQuery.eof = true
	strPerioder = strPerioder & rsQuery("periodeID") & ","
	rsQuery.movenext
wend

'response.write strPerioder

Set rsQuery = Con.Execute("SELECT tblSkiBrugerTider.vagtled, Count(tblSkiBrugerTider.pw) AS AntalOfpw, tblSkiPerioder.value, tblSkiPerioder.dato, tblSkiPerioder.tid, tblSkiPerioder.ekstra, tblSkiPerioder.max, tblSkiPerioder.graa, tblSkiPerioder.id, tblSkiPerioder.sortorder FROM tblSkiBrugerTider RIGHT JOIN tblSkiPerioder ON tblSkiBrugerTider.PeriodeID = tblSkiPerioder.id WHERE email like '" & session("email") & "' and pw like " & session("pw") & " GROUP BY tblSkiBrugerTider.vagtled, tblSkiPerioder.dato, tblSkiPerioder.tid, tblSkiPerioder.ekstra, tblSkiPerioder.value,tblSkiPerioder.max, tblSkiPerioder.graa, tblSkiPerioder.id, tblSkiPerioder.sortorder ORDER BY tblSkiPerioder.sortorder" )
 %> 


<%
rsquery1.movefirst
strSum = rsQuery1("OPS1") & rsQuery1("OPS2") & rsQuery1("OPS3") & rsQuery1("NED1") 
'response.write strSum
rsquery1.movefirst
 %>

<%i = 0 %>
 
 
 <%if len(strSum) > 0 then %>

Du har fået tildelt følgende ekstraordinære vagter:
<ul>
<%if rsQuery1("1") = "1" then%>
<li>Opsætning - <strong>25. juni 16-??</strong></li>
<%'i = i + 6 %>
<%end if %>

 <%if rsQuery1("2") = "1" then%>
<li>Opsætning - <strong>26. juni 16-??</strong></li>
<%'i = i + 6 %>
<%end if %>

 <%if rsQuery1("3") = "1" then%>
<li>Opsætning - <strong>28. juni 16-??</strong></li>
<%'i = i + 6 %>
<%end if %>

 <%if rsQuery1("4") = "1" then%>
<li>Nedtagning - <strong>6. juli 16-??</strong></li>
<%'i = i + 6 %>
<%end if %>


</ul>

<%end if %>


Du har fået tildelt følgende tider under festivalen:<br>
<!-- tider -->
<ul>
<%while not rsQuery.eof = true %>

<li><%=weekdayname(weekday(rsQuery("dato"))) %>&nbsp;<%=formatdateTime(rsQuery("dato"),1) %> - <%=rsQuery("tid") %>&nbsp;<%=rsQuery("ekstra") %>


<%if rsQuery("vagtled") = "1" then %>
<strong>(Vagtleder !)</strong>
<%end if %>

</li>
<%i = i + rsQuery("value") %>

<%
rsQuery.movenext
wend
 %>
</ul>
 <br>
   
Du har ialt <%=i %> timer.
<br>
<br>
<input onclick="javascript:window.print()" type="button" value="Udskriv">
</body>
</html>
<%
Set Con = Nothing
Set rsQuery   = Nothing
 %>
