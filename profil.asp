

<!--#include file="conn.asp" -->

<%Set rsQuery = Con.Execute("SELECT * FROM tblSkiTid") 
strTid = rsQuery("tid")
%>



<%
Set rsQuery1 = Con.Execute("SELECT *, camping, mnummer, bemaerk, OPS1, OPS2, OPS3, NED1, andrebesk, edited, gruppe FROM tblSkiBruger WHERE email like '" & session("email") & "' and pw like " & session("pw") & "")

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

<%
'if session("overlap") = 1 then 
'strOnload = "alert('Du har valgt overlappende perioder.\n Vær opmærksom på, at perioder afmærket med rød ramme\n overlapper de sidste vagter.');"
'end if 
'response.write session("overlap")
'session("overlap") = 0
%>

<%
if session("dublet") = 1 then 
strOnload = "alert('Du kan ikke opdatere dit navn, din adresse eller postnummer som ønsket.\n En sådan bruger er allerede oprettet i databasen.');"
end if 
session("dublet") = 0
%>

<%
if session("24") < 18  and session("24") <> "" then 
strOnload = "alert('Du skal vælge vagter svarende til mindst 18 timer.\n Du har valgt " & session("24") & " timer');"
end if 
session("24") = null
%>

<%
'session("12") = 0
if session("12") < 9 and session("12") <> "" then 
strOnload = "alert('Du skal vælge vagter svarende til mindst 9 timer i de grå perioder.\n Du har kun valgt " & session("12") & " timer')"
end if 
session("12") = 13
%>

<%if session("obl") = 1 then 
strOnload = "alert('Du skal huske at udfylde de obligatoriske felter.')"
end if 
session("obl") = 0
%>

<%if session("tak") = 1 then 
strOnload = "alert('Tak for din tilmelding.\n Vi håber, at kunne give dig de ønskede vagter, men forbeholder os ret til at justere dem. \n Du vil i starten af juni modtage en mail, hvor vi bekræfter dine endelige vagter.')"
end if 
session("tak") = 0
%>


<body onload="<%=strOnload %>">

<div style="width:450px">
<img src="logo.gif" height="124" width="124" align="right">
<h1>Roskilde Skiklub</h1>
<h2>Tilmelding til Roskilde Festival</h2>
</div>
<br>

<a href="logaf.asp">Log af</a>
<br>
<br>

<form name="frmSample" action="updateProfil.asp" method="get" onSubmit="return validateStandard(this);">
<table border="0" cellpadding="5" cellspacing="1">
<tr>
<td valign="top">
<!-- grunddata -->
<table border="0" cellpadding="5" cellspacing="1">
<tr>
<td>Fornavn&nbsp;<span style="color:red">*</span></td><td><input type="text" name="navn" value="<%=rsQuery1("navn") %>" required="1" err="Indtast fornavn"></td>
</tr>
<tr>
<td>Efternavn&nbsp;<span style="color:red">*</span></td><td><input type="text" name="efternavn" value="<%=rsQuery1("efternavn") %>" required="1" err="Indtast efternavn"></td>
</tr>
<tr>
<td>Fødselsdato&nbsp;<span style="color:red">*</span><br>
(format: dd-mm-åååå)</td><td><input type="text" name="fdato" value="<%=formatdatetime(rsQuery1("fdato"),2) %>" required="1" regexp="JSVAL_RX_TEL" err="Indtast fødselsdato i følgende format: dd-mm-åååå"></td>
</tr>
<tr>
<td>Adresse&nbsp;<span style="color:red">*</span></td><td><input type="text" name="adr" value="<%=rsQuery1("adresse") %>" required="1" err="Indtast adresse"></td>
</tr>
<tr>
<td>Postnummer&nbsp;<span style="color:red">*</span></td><td><input type="text" name="postnr" maxlength="4"  value="<%=rsQuery1("postnr") %>" maxlength="4" minlength="4" required="1" regexp="JSVAL_RX_ZIP" err="Indtast postnummer" ></td>
</tr>
<tr>
<td>By&nbsp;<span style="color:red">*</span></td><td><input type="text" name="by"  value="<%=rsQuery1("bye") %>" required="1" err="Indtast by"></td>
</tr>
<tr>
<td>Land&nbsp;<span style="color:red">*</span></td><td><input type="text" name="land"  value="<%=rsQuery1("land") %>" required="1" err="Indtast land"></td>
</tr>
<tr>
<td>Telefon&nbsp;<span style="color:red">*</span></td><td><input type="text" name="tlf" maxlength="8" value="<%=rsQuery1("telefon") %>" required="1"  regexp="JSVAL_RX_ZIP" maxlength="8" minlength="8" err="Indtast telefonnummer"></td>
</tr>
<tr>
<td>Email</td><td><input type="text" readonly name="email" value="<%=rsQuery1("email") %>"></td>
</tr>
<tr>
<td valign="top">Password<br>
(husk dette til næste log-ind)</td><td valign="top"><input type="text" readonly name="pw" value="<%=rsQuery1("pw") %>"></td>
</tr>
<tr>
<td valign="top">Medlemsnummer<br>
(Står på bagsiden af Bindingen)</td><td valign="top"><input type="text" name="medlem" value="<%=rsQuery1("medlem") %>"></td>
</tr>
<tr>
<td>Hvis ikke medlem, <br>
hvem er da din kontaktperson ?</td><td><input type="text" name="kontakt"  value="<%=rsQuery1("kontakt") %>"></td>
</tr>
<tr>
<td>Jeg har arbejdet i boden før</td><td><input type="checkbox" name="tidlArb" value="Ja" <%if rsQuery1("tidlArb") = "Ja" then response.write "checked" %>></td>
</tr>
<tr>
<td>T-shirt størrelse</td><td>
<select name="tshirtStr">
<option <%if rsQuery1("tshirtstr") = "Small" then response.write "selected" %>>Small</option>
<option <%if rsQuery1("tshirtstr") = "Medium" then response.write "selected" %>>Medium</option>
<option <%if rsQuery1("tshirtstr") = "Large" then response.write "selected" %>>Large</option>
<option <%if rsQuery1("tshirtstr") = "X-Large" then response.write "selected" %>>X-Large</option>
<option <%if rsQuery1("tshirtstr") = "XX-Large" then response.write "selected" %>>XX-Large</option>

</select>
</td>
</tr>
<!-- <tr>
<td>Jeg ønsker at benytte <br>
medarbejder camping (kr. 150,-)</td><td><input type="checkbox" name="camping" value="Ja" <%if rsQuery1("camping") = "Ja" then response.write "checked" %>></td>
</tr> -->
<tr><td>Bemærkninger</td><td>
<textarea name="bemaerk">
<%if session("bemaerk") <> "" then
response.write session("bemaerk")
else
response.write rsQuery1("bemaerk")
end if %>
</textarea>
</td></tr>
</table>

<!-- grunddata slut -->
</td>
<%

'Set Con = Nothing
'Set rsQuery   = Nothing

Set rsQuery = Con.Execute("SELECT tblSkiBrugerTider.PeriodeID FROM tblSkiBrugerTider WHERE email like '" & session("email") & "' and tblSkiBrugerTider.pw like " &  session("pw") & "")

strPerioder = ","
while not rsQuery.eof = true
strPerioder = strPerioder & rsQuery("periodeID") & ","
rsQuery.movenext
wend

'response.write strPerioder

sql = "SELECT tblSkiPerioder.dato, tblSkiPerioder.tid, Count(tblSkiBrugerTider.pw) AS AntalOfpw, tblSkiPerioder.value, tblSkiPerioder.max, tblSkiPerioder.graa, tblSkiPerioder.id, tblSkiPerioder.sortorder,tblSkiPerioder.valis FROM tblSkiBrugerTider RIGHT JOIN tblSkiPerioder ON tblSkiBrugerTider.PeriodeID = tblSkiPerioder.id GROUP BY tblSkiPerioder.dato, tblSkiPerioder.tid, tblSkiPerioder.value, tblSkiPerioder.max, tblSkiPerioder.valis, tblSkiPerioder.graa, tblSkiPerioder.id, tblSkiPerioder.sortorder ORDER BY tblSkiPerioder.sortorder"

'"SELECT tblPerioder.periode, Count(tblBrugerTider.pw) AS AntalOfpw, tblPerioder.value, tblPerioder.max, tblPerioder.graa, tblPerioder.dato, tblPerioder.tid, tblPerioder.id, tblPerioder.sortorder FROM tblBrugerTider RIGHT JOIN tblPerioder ON tblBrugerTider.PeriodeID = tblPerioder.id GROUP BY tblPerioder.dato,tblPerioder.tid,tblPerioder.periode, tblPerioder.value, tblPerioder.max, tblPerioder.graa, tblPerioder.id, tblPerioder.sortorder ORDER BY tblPerioder.sortorder"

Set rsQuery = Con.Execute(sql)


'SELECT tblPerioder.graa, tblPerioder.periode,  tblPerioder.id, tblPerioder.value FROM tblPerioder ORDER BY tblPerioder.sortorder")

'SELECT tblPerioder.graa, tblPerioder.value, tblPerioder.id, tblPerioder.periode, tblBruger.email, tblBruger.pw FROM (tblBrugerTider RIGHT JOIN tblPerioder ON tblBrugerTider.PeriodeID = tblPerioder.id) LEFT JOIN tblBruger ON tblBrugerTider.pw = tblBruger.pw")

Intvalue = 0
IntValue2 = 0

while not rsQuery.eof = true

if instr(strPerioder, "," & rsQuery("id") & ",") > 0  then
Intvalue = Intvalue + rsQuery("value") 

	if rsQuery("graa") > 0 and rsQuery("graa") < 3 then
		IntValue2 = IntValue2 + rsQuery("value")
	end if
end if

rsQuery.movenext
wend
rsQuery.movefirst
 %> 

<td valign="top">
<!-- status -->

<strong>Nu skal du vælge de vagter du ønsker dig:</strong>
<br>
<br>
<table border="0" cellpadding="5" cellspacing="1" width="400px">
<!-- <tr>
<td colspan="2">
Opsætning/nedtagning af boden (tæller ikke med i dine vagttimer)
</td>
</tr> -->

<!-- <tr>
<td colspan="2">
<input type="checkbox" name="OPS1" value="Ja" 
<%if session("OPS1") <> "" then
response.write "checked "
end if %>
<%if rsQuery1("OPS1") = "Ja" then response.write "checked" %>>&nbsp;
Opsætning torsdag 25. juni (16.00 til "aften")
</td>
</tr> -->

<!-- <tr>
<td colspan="2">
<input type="checkbox" name="OPS2" value="Ja" 
<%if session("OPS2") <> "" then
response.write "checked "
end if %>
<%if rsQuery1("OPS2") = "Ja" then response.write "checked" %>>&nbsp;
Opsætning fredag 26. juni (16.00 til "aften")
</td>
</tr> -->

<!-- <tr>
<td colspan="2">
<input type="checkbox" name="OPS3" value="Ja" 
<%if session("OPS3") <> "" then
response.write "checked "
end if %>
<%if rsQuery1("OPS3") = "Ja" then response.write "checked" %>>&nbsp;
Opsætning søndag 28. juni (9.00 til "vi er færdige")
</td>
</tr> -->

<!-- <tr>
<td colspan="2">
<input type="checkbox" name="NED1" value="Ja" 
<%if session("NED1") <> "" then
response.write "checked "
end if %>
<%if rsQuery1("NED1") = "Ja" then response.write "checked" %>>&nbsp;
Nedtagning mandag 6. juli (16.00 til "vi er færdige")
</td>
</tr> -->
<!-- <tr>
<td colspan="2">
<strong>Vagter før festivalen </strong><br>
Hvis du arbejder før festivalen (inden 3. juli) skal du have mindst 3 vagter, hvoraf minimum én vagt skal ligge under festivalen.
</td>
</tr> -->
<tr>
<td colspan="2">
For information om opsætnings- og nedtagningsvagter, kontakt <a href="mailto:thomas@svanevej.dk">Thomas Lindegaard Nielsen</a>

</td>
</tr>
<tr>
<td colspan="2">
<strong>Vagter under festivalen</strong>
<br>
Hvis du kun arbejder under festivalen (fra den 1. juli) skal du have 2 vagter, hvoraf den ene skal ligge i det "grå" felt.
</td>
</tr>

<tr>
<td>Jeg vil gerne<br>
være vagtleder</td><td><input type="checkbox" name="vagtleder" value="Ja"  <%if rsQuery1("vagtleder") = "Ja" then response.write "checked" %>></td>
</tr>

<tr>
<td colspan="2">
Nedenfor skal du sætte flueben udfor de vagter du ønsker. 
<br>
Din nuværende vagt-status er:<br>
Samlet timeantal = <strong><%=Intvalue %></strong><br>
Heraf er <strong><%=Intvalue2 %></strong> timer placeret i de travle perioder (med Grå)

<%if Intvalue < 20  then %>
<br><span style="color:red"><strong>OBS</strong></span>: Bemærk, at du skal vælge perioder svarende til mindst 20 timer - du mangler <%=20 - Intvalue %> timer.
<%end if %>

<%if Intvalue > 26  then %>
<br><span style="color:red"><strong>OBS</strong></span>: Bemærk, at du skal vælge perioder svarende til højst 26 timer.
<%end if %>

<%if Intvalue2 < 9 then %>
<br><span style="color:red"><strong>OBS</strong></span>: Bemærk, at du skal vælge perioder svarende til <strong><u>mindst 9 timer i de grå perioder</u></strong> - du mangler <%=9 - Intvalue2 %> timer.
<%end if %>
</td>
</tr>

<tr>
<td colspan="2">
Gruppe: <input type="text" name="gruppe" size="20" value="<%=rsQuery1("gruppe") %>">
<br>
Hvis I er flere som ønsker at arbejde på de samme vagter, skal I enkeltvis udfylde tilmeldingen og hver især angive det samme kodeord i dette felt ('x-factor','ketchup', 'snowboard' ell. lign) 
<br>
<br>
Er du i tvivl om dine registrerede vagtønsker så <a href="profil.asp">klik her for at genindlæse siden.</a>
</td>
</tr>

</table>
<!-- status slut -->

<table border="0" cellpadding="5" cellspacing="1" width="400px">

<%
'slut på nattevagter
'end if
 %>


</table>
<br>
<br>





</td>
</tr>
<tr>
<td colspan="2">
<!-- tider -->
<%
' angiv startdato
strDate = "25-06-2010"
 %>
<table border="0" cellpadding="5" cellspacing="1">
<tr>

<%
' indsættelse af datoer
for i = 0 to 8
strDate =  DateAdd("d", 1, strDate)
 %>
 <td align="center">
 <strong>
 <%=weekdayname(weekday(strDate)) %>
<br>
  <%=formatDateTime(strDate,2) %>
 </strong></td>
<%
next
'loop
 %>
</tr>
<tr>


<%

while not rsQuery.eof = true
'if strDag <> left(rsQuery("periode"),4) then
if strDag <> rsQuery("dato") then
response.write "<td valign=top>"
end if 
'strDag = left(rsQuery("periode"),4)
%>

<div style="padding:2px;<%if rsQuery("graa") > 0  then response.write "background-color:#cccccc" %>;">


<input 

<%
' er der overhovedet nogle data i Valis, dvs skal der indsættes validering på overlap
if rsQuery("valis") <> "" then 
%>
onclick="

<%
	' der var data i Valis. Opsplittes i array
arrValis = split(rsQuery("valis"),";")
	' for hver tal adskilt af semikolon oprettes en scriptlinie
for i = 0 to ubound(arrValis)
 %>
window.document.frmSample.check_<%=arrValis(i) %>.checked=false;

<%next %>
" 
<%
	' slut med at validere på overlap
end if 
%>

id="check_<%=rsQuery("id") %>" name="periode<%=rsQuery("id") %>" value="<%=rsQuery("id") %>" type="checkbox" <%if rsQuery("id") = session("check" & rsQuery("id")) then
response.write "checked "
end if

 %><%if instr(strPerioder, "," & rsQuery("id") & ",") > 0  then
 response.write "checked" 
Intvalue = Intvalue + rsQuery("value") 

	if rsQuery("graa") > 0 then
		IntValue2 = IntValue2 + rsQuery("value")
	end if
else
	if rsQuery("AntalOfpw") >= rsQuery("max") then
	response.write " disabled "
	end if
end if%>>&nbsp;<%=rsQuery("tid") %>&nbsp;&nbsp;
<%
IntPladser = rsQuery("max")-rsQuery("AntalOfpw") 
' hvis der skal angives ledige pladser:
'response.write IntPladser
'	if IntPladser <> 1 then
'	response.write "&nbsp;pladser tilbage."
'	else
'	response.write "&nbsp;plads tilbage."
'	end if

%>
</div>
<%
strDag = rsQuery("dato")
%>
<%
rsQuery.movenext
wend
 %> 
</td>
</tr>
</table>
<!-- tider slut -->
</td>
</tr>
</table>
<%if strTid = "1" then %>
<input type="submit" value="Opdater mine oplysninger">
<%end if %>

</form>

</body>
</html>
<%
Set Con = Nothing
Set rsQuery   = Nothing
 %>
 <%for i = 1 to 50
'if instr(request.servervariables("query_string"),"periode" & i & "=") > 0 then
session("check" & i) = ""
'end if
next
 %>
 <%
 session("bemaerk") = ""
'session("opsaetTelt") = ""
'session("klarBod") = ""
'session("nattevagt") = ""
'session("udlaegPall") = ""
'session("nedtagTelt") = ""
'session("nedtagBod") = ""
'session("nt") = ""
'session("nf") = ""
'session("nl") = ""
'session("ns") = ""
  %>