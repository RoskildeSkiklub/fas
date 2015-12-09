<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>

<!--#include file="conn.asp" -->
<%
Set rsQuery1 = Con.Execute("SELECT *, camping, mnummer, bemaerk, OPS1, OPS2, OPS3, NED1, andrebesk, gruppe, arbopg, edited FROM tblSkiBruger WHERE email like '" & request("email") & "' and pw like " & request("pw") & "")

 %>
 
 <%Set rsQuery3 = Con.Execute("SELECT tblSkiBrugerTider.vagtled, tblSkiBrugerTider.email, tblSkiBrugerTider.pw, tblSkiPerioder.id FROM tblSkiBrugerTider INNER JOIN tblSkiPerioder ON tblSkiBrugerTider.PeriodeID = tblSkiPerioder.id WHERE (((tblSkiBrugerTider.vagtled)='1')) and email like '" & request("email") & "' and pw like " & request("pw") & "")
strVagtleder = ","
while not rsQuery3.eof = true
strVagtleder = strVagtleder & rsquery3("id") & ","
rsquery3.movenext
wend
'response.write strVagtleder
 %>

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
<div style="width:450px">
<img src="../logo.gif" height="62" width="62" align="right">
<h1>Roskilde Skiklub</h1>
<h2>Tilmelding til Roskilde Festival</h2>
</div>
<br>


<form name="frmSample" action="updateProfil-adm.asp" method="get" onSubmit="return validateStandard(this);">
<table border="0" cellpadding="5" cellspacing="1">
<tr>
<td>
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
(format: dd-mm-åååå)</td><td><input type="text" name="fdato" value="<%=rsQuery1("fdato") %>" required="1" regexp="JSVAL_RX_TEL" err="Indtast fødselsdato i følgende format: dd-mm-åååå"></td>
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
<td>Email</td><td><input readonly type="text" name="email" value="<%=rsQuery1("email") %>"></td>
</tr>
<tr>
<td>Password</td><td><input type="text" readonly name="pw" value="<%=rsQuery1("pw") %>"></td>
</tr>
<tr>
<td>Medlemsnummer</td><td><input type="text" name="medlem" value="<%=rsQuery1("medlem") %>"></td>
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
<tr>
<td>Medarbejdernummer</td><td><input type="text" name="mnummer" value="<%=rsQuery1("mnummer") %>"></td>
</tr>
<tr><td><strong>Bemærkninger</strong>
</td><td>
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

Set rsQuery = Con.Execute("SELECT tblSkiBrugerTider.PeriodeID FROM tblSkiBrugerTider WHERE email like '" & request("email") & "' and tblSkiBrugerTider.pw like " &  request("pw") & " ")

strPerioder = ","
while not rsQuery.eof = true
strPerioder = strPerioder & rsQuery("periodeID") & ","
rsQuery.movenext
wend

'response.write strPerioder

sql = "SELECT tblSkiPerioder.dato, tblSkiPerioder.tid, Count(tblSkiBrugerTider.pw) AS AntalOfpw, tblSkiPerioder.value, tblSkiPerioder.max, tblSkiPerioder.graa, tblSkiPerioder.id, tblSkiPerioder.sortorder,tblSkiPerioder.valis FROM tblSkiBrugerTider RIGHT JOIN tblSkiPerioder ON tblSkiBrugerTider.PeriodeID = tblSkiPerioder.id GROUP BY tblSkiPerioder.dato, tblSkiPerioder.tid, tblSkiPerioder.value, tblSkiPerioder.max, tblSkiPerioder.valis, tblSkiPerioder.graa, tblSkiPerioder.id, tblSkiPerioder.sortorder ORDER BY tblSkiPerioder.sortorder"

Set rsQuery = Con.Execute(sql)

'"SELECT tblPerioder.periode, Count(tblBrugerTider.pw) AS AntalOfpw, tblPerioder.value, tblPerioder.max, tblPerioder.graa, tblPerioder.id, tblPerioder.sortorder FROM tblBrugerTider RIGHT JOIN tblPerioder ON tblBrugerTider.PeriodeID = tblPerioder.id GROUP BY tblPerioder.periode, tblPerioder.value, tblPerioder.max, tblPerioder.graa, tblPerioder.id, tblPerioder.sortorder ORDER BY tblPerioder.sortorder")


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
<tr>
<td colspan="2">
Opsætning/nedtagning af boden (tæller ikke med i dine vagttimer)
</td>
</tr>

<tr>
<td colspan="2">
<input type="checkbox" name="OPS1" value="Ja" 
<%if session("OPS1") <> "" then
response.write "checked "
end if %>
<%if rsQuery1("OPS1") = "Ja" then response.write "checked" %>>&nbsp;
Opsætning 25. juni 16-??
&nbsp;(
<%Set rsQuery2 = Con.Execute("SELECT Count(tblSkiBruger.id) AS AntalOfid FROM tblSkiBruger GROUP BY tblSkiBruger.OPS1 HAVING (((tblSkiBruger.OPS1)='ja'));")
 %>
 <%if rsQuery2.eof = true then 
 response.write 0
 else
 response.write rsQuery2("antalofid")
 end if %>
&nbsp;allerede tilmeldt)</td>
</tr>

<tr>
<td colspan="2">
<input type="checkbox" name="OPS2" value="Ja" 
<%if session("OPS2") <> "" then
response.write "checked "
end if %>
<%if rsQuery1("OPS2") = "Ja" then response.write "checked" %>>&nbsp;
Opsætning 26. juni 16-??
&nbsp;(
<%Set rsQuery2 = Con.Execute("SELECT Count(tblSkiBruger.id) AS AntalOfid FROM tblSkiBruger GROUP BY tblSkiBruger.OPS2 HAVING (((tblSkiBruger.OPS2)='ja'));")
 %>
 <%if rsQuery2.eof = true then 
 response.write 0
 else
 response.write rsQuery2("antalofid")
 end if %>
&nbsp;allerede tilmeldt)</td>
</tr>

<tr>
<td colspan="2">
<input type="checkbox" name="OPS3" value="Ja" 
<%if session("OPS3") <> "" then
response.write "checked "
end if %>
<%if rsQuery1("OPS3") = "Ja" then response.write "checked" %>>&nbsp;
Opsætning 28. juni 09-??
&nbsp;(
<%Set rsQuery2 = Con.Execute("SELECT Count(tblSkiBruger.id) AS AntalOfid FROM tblSkiBruger GROUP BY tblSkiBruger.OPS3 HAVING (((tblSkiBruger.OPS3)='ja'));")
 %>
 <%if rsQuery2.eof = true then 
 response.write 0
 else
 response.write rsQuery2("antalofid")
 end if %>
&nbsp;allerede tilmeldt)</td>
</tr>

<tr>
<td colspan="2">
<input type="checkbox" name="NED1" value="Ja" 
<%if session("NED1") <> "" then
response.write "checked "
end if %>
<%if rsQuery1("NED1") = "Ja" then response.write "checked" %>>&nbsp;
Nedtagning 6. juli 16-??
&nbsp;(
<%Set rsQuery2 = Con.Execute("SELECT Count(tblSkiBruger.id) AS AntalOfid FROM tblSkiBruger GROUP BY tblSkiBruger.NED1 HAVING (((tblSkiBruger.NED1)='ja'));")
 %>
 <%if rsQuery2.eof = true then 
 response.write 0
 else
 response.write rsQuery2("antalofid")
 end if %>
&nbsp;allerede tilmeldt)</td>
</tr>

<tr>
<td colspan="2"><input type="checkbox" name="vagtleder" value="Ja"  <%if rsQuery1("vagtleder") = "Ja" then response.write "checked" %>>&nbsp;Jeg vil gerne være vagtleder</td>
</tr>

<tr>
<td colspan="2">
Din nuværende vagt-status er:<br>
Samlet timeantal = <strong><%=Intvalue %></strong><br>
Heraf er <strong><%=Intvalue2 %></strong> timer placeret i de travle perioder (med Grå)

<%if Intvalue < 20  then %>
<br><span style="color:red"><strong>OBS</strong></span>: Bemærk, at du skal vælge perioder svarende til mindst 20 timer - du mangler <%=20 - Intvalue %> timer.
<%end if %>

<%if Intvalue > 26  then %>
<br><span style="color:red"><strong>OBS</strong></span>: Bemærk, at du skal vælge perioder svarende til højst 26 timer.
<%end if %>
</td>
</tr>
<tr>
<td colspan="2">
Er du i tvivl om de registrerede vagter så <a href="profil-adm.asp">klik her for at genindlæse siden.</a>
<br>
<br>
<strong>Beskeder til bruger:</strong><br>
<textarea name="andreBesk" cols="40">
<%if session("bemaerk") <> "" then
response.write session("andreBesk")
else
response.write rsQuery1("andreBesk")
end if %>
</textarea>
<br>
<br>
Gruppe: <input type="text" name="gruppe" size="20" value="<%=rsQuery1("gruppe") %>">
<br>
Arbejdsopgaver: <input type="text" name="arbopg" size="20" value="<%=rsQuery1("arbopg") %>">
<br>
<br>
<input type="checkbox" value="Ja" name="edited" <%if rsQuery1("edited") = "Ja" then response.write "checked" %>>&nbsp;Behandlet (sæt flueben her, hvis du har færdigbehandlet brugerens ønsker)
</td>
</tr>

<!-- status slut -->
</table>


</td>
</tr>
<tr>
<td colspan="2">
<!-- tider -->

<%
strDate = "25-06-2010"
 %>
<table border="0" cellpadding="5" cellspacing="1">
<tr>

<%
' indsættelse af datoer
for i = 0 to 8
strDate =  DateAdd("d", 1, strDate)
 %>
 <td align="center"><strong>
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
'	if rsQuery("AntalOfpw") >= rsQuery("max") then
'	response.write " disabled "
'	end if
end if%>>&nbsp;<%=rsQuery("tid") %><br>
&nbsp;
(
<%
IntPladser = rsQuery("max")-rsQuery("AntalOfpw") 
response.write IntPladser
'	if IntPladser <> 1 then
'	response.write "&nbsp;pladser tilbage."
'	else
'	response.write "&nbsp;plads tilbage."
'	end if

%>
fri)</div>
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
<input type="submit" value="Opdater mine oplysninger">

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
session("opsaetTelt") = ""
session("klarBod") = ""
session("nattevagt") = ""
session("udlaegPall") = ""
session("nedtagTelt") = ""
session("nedtagBod") = ""
  %>