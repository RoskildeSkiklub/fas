<%response.buffer = false %>
<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>
<%
Randomize
rndMaxInterval = 99999999
rndInterval = Int((rndMaxInterval * Rnd)+1) 

select case len(rndInterval)
case 1 
rndInterval = rndInterval & "0000000"
case 2 
rndInterval = rndInterval & "000000"
case 3 
rndInterval = rndInterval & "00000"
case 4 
rndInterval = rndInterval & "0000"
case 5 
rndInterval = rndInterval & "000"
case 6 
rndInterval = rndInterval & "00"
case 7 
rndInterval = rndInterval & "0"
end select
 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
	<title>Opret profil</title>
<LINK HREF="adm.css" REL="stylesheet" TYPE="text/css">
	<style type="text/css">
table{
	background-color:#aaaaaa
}
td{
	background-color: #ffffff
}
</style>
<script language="javascript" src="jsval.js"></script>
</head>

<body OnLoad="setVariables();checkLocation()">
<div style="width:450px">
<img src="logo.jpg" height="81" width="100" align="right">
<h1>Himmelev Volleyball</h1>
<h2>Tilmelding til Roskilde Festival</h2>
</div>
<br>
<a href="logpaa.asp">Tilbage til forsiden</a>
<br>
<br>
Felter markeret med <span style="color:red">*</span> <u>skal</u> udfyldes.
<br>
<br>
<!-- validateStandard(this); -->
<form name="testform" action="adduser.asp" method="get" onSubmit="return validateStandard(this);">

<!-- <form name="frmSample" method="get" action="adduser.asp" onSubmit="return validateStandard(this,'error');"> -->
<table border="0" cellpadding="5" cellspacing="1">
<tr>
<td>Fornavn&nbsp;<span style="color:red">*</span></td><td><input type="text" name="navn" required="1" err="Indtast fornavn"></td>
</tr>
<tr>
<td>Efternavn&nbsp;<span style="color:red">*</span></td><td><input type="text" name="efternavn" required="1" err="Indtast efternavn"></td>
</tr>
<td>Fødselsdato&nbsp;<span style="color:red">*</span><br>
(format: dd-mm-åååå)</td><td><input type="text" name="fdato" required="1" regexp="JSVAL_RX_TEL" err="Indtast fødselsdato i følgende format: dd-mm-åååå"></td>
</tr>
<tr>
<td>Adresse&nbsp;<span style="color:red">*</span></td><td><input type="text" name="adr" required="1" err="Indtast adresse"></td>
</tr>
<tr>
<td>Postnummer&nbsp;<span style="color:red">*</span></td><td><input type="text" name="postnr" maxlength="4" minlength="4" required="1" regexp="JSVAL_RX_ZIP" err="Indtast postnummer" ></td>
</tr>
<tr>
<td>By&nbsp;<span style="color:red">*</span></td><td><input type="text" name="by" required="1" err="Indtast by"></td>
</tr>
<tr>
<td>Land&nbsp;<span style="color:red">*</span></td><td><input type="text" name="land" value="Danmark" required="1" err="Indtast land" ></td>
</tr>
<tr>
<td>Telefon&nbsp;<span style="color:red">*</span></td><td><input type="text" name="tlf" required="1"  regexp="JSVAL_RX_ZIP" maxlength="8" minlength="8" err="Indtast telefonnummer"></td>
</tr>
<tr>
<td>Email&nbsp;<span style="color:red">*</span></td><td><input type="text" name="email" id="email" required="1" regexp="JSVAL_RX_EMAIL" err="Indtast email adresse"></td>
</tr>
<tr>
<td>Password</td><td><input type="text" readonly name="pw" value="<%response.write rndInterval%>"></td>
</tr>
<tr>
<td>T-shirt størrelse</td><td>
<select name="tshirtStr">
<option>Small</option>
<option>Medium</option>
<option>Large</option>
<option>X-Large</option>
<option>XX-Large</option>
</select>
</td>
</tr>
<tr>
<td>Jeg vil gerne være vagtleder</td><td><input type="checkbox" name="vagtleder" value="Ja"></td>
</tr>
<tr>
<td>Jeg vil helst have følgende vagter</td><td>
<input type="checkbox" name="4x6" value="Ja">&nbsp;4 x 6 timer<br>
<input type="checkbox" name="3x8" value="Ja">&nbsp;3 x 8 timer<br>
<input type="checkbox" name="2x12" value="Ja">&nbsp;2 x 12 timer
</td>
</tr>
<tr>
<td>Jeg ønsker at benytte <br>
medarbejder camping (kr. 150,-)</td><td><input type="checkbox" name="camping"  value="Ja"></td>
</tr>
<tr>
<td>Medlem af klubben</td><td><input type="checkbox" name="medlem"  value="Ja"></td>
</tr>
<tr>
<td>Hvis ikke medlem, hvem er da din kontaktperson ?</td><td><input type="text" name="kontakt"></td>
</tr>
<tr><td>Bemærkninger</td><td>
<textarea name="bemaerk"></textarea>
</td></tr>
<tr>
<td colspan="2">&nbsp;&nbsp;<input type="submit" value="Gå til timevalg"></td>
</tr>

</table>
</form>
</body>
</html>
