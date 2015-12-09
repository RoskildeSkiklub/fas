<%response.buffer = false %>

<%
'response.write datediff("d", "02-07-1993", "01-07-1993")
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

<!--#include file="conn.asp" -->
<%
Set rsQuery = Con.Execute("SELECT * FROM tblSkiData")
str18 = rsQuery("18aar")
str15 = rsQuery("15aar")

Set rsQuery = Con.Execute("SELECT CountOfid FROM Vw_SkiUnder")
strUnder = rsQuery("CountOfId")


Set rsQuery = Con.Execute("SELECT CountOfid FROM Vw_SkiAll ")
strAll = rsQuery("CountOfId")

intProcent = strUnder/strAll * 100

'response.write strUnder
'response.write intProcent

 %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
	<title>Opret profil</title>
<style>
table{
	background-color:#aaaaaa
}
td{
	background-color: #ffffff
}
</style>
<script language="javascript" src="jsval.js"></script>

</head>
<%'=session("dublet") %>
<%
if session("dublet") = 1 then 
%>
<body onload="alert('Du kan ikke oprette en profil som ønsket.\n En sådan bruger er allerede oprettet i databasen.')"
<%
end if 
session("dublet") = 0
%>

<%if session("succes") = 1 then %>
<body onload="alert('Din profil er oprettet.')">
<%
session("succes") = null
end if %>

<%if session("succes") = 2 then %>
<body onload="alert('Du har ikke udfyldt alle obligatoriske felter - prøv igen.')">
<%
session("succes") = null
end if %>

<%if session("succes") = 5 then %>
<body onload="alert('Du er født efter den angivne dato. Du kan ikke tilmeldes.')">
<%
session("succes") = null
end if %>

<%if session("succes") = 4 then %>
<body onload="alert('Der var problmer med at logge på - prøv igen.')">
<%
session("succes") = null
end if %>

<%if session("succes") = null then %>
<body>
<%
session("succes") = null
end if%>
<div style="width:450px">
<img src="logo.gif" height="124" width="124" align="right">
<h1>Roskilde Skiklub</h1>
<h2>Tilmelding til Roskilde Festival</h2>
</div>
<br>
<a href="index.asp">Tilbage til forsiden</a>
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
<td>Fornavn&nbsp;<span style="color:red">*</span></td><td><input type="text" name="navn" required="1" err="Indtast fornavn" value="<%=session("Onavn") %>"></td>
</tr>
<tr>
<td>Efternavn&nbsp;<span style="color:red">*</span></td><td><input type="text" name="efternavn" required="1" err="Indtast efternavn" value="<%=session("Oefternavn") %>"></td>
</tr>
<td>Fødselsdato&nbsp;<span style="color:red">*</span><br>
(format: dd-mm-åååå)</td>
<td style="font-size:10px">

<input type="text" name="fdato" required="1" regexp="JSVAL_RX_TEL" err="Indtast fødselsdato i følgende format: dd-mm-åååå">
<br>
<%
'tallet ændres til den procentsats der skal være grænsen for personer under 18
' dvs ved 90 skal blot 10% være over 18
if intProcent > 90 then 
%>
Du skal være født før <%=formatdatetime(str18,1) %>
<%else %>
Du skal være født før <%=formatdatetime(str15,1) %>
<%end if %>
</td>
</tr>
<tr>
<td>Adresse&nbsp;<span style="color:red">*</span></td><td><input type="text" name="adr" required="1" err="Indtast adresse"  value="<%=session("Oadr") %>"></td>
</tr>
<tr>
<td>Postnummer&nbsp;<span style="color:red">*</span></td><td><input type="text" name="postnr" maxlength="4" minlength="4" required="1" regexp="JSVAL_RX_ZIP" err="Indtast postnummer" value="<%=session("Opostnr") %>"></td>
</tr>
<tr>
<td>By&nbsp;<span style="color:red">*</span></td><td><input type="text" name="by" required="1" err="Indtast by" value="<%=session("Oby") %>"></td>
</tr>
<tr>
<td>Land&nbsp;<span style="color:red">*</span></td><td><input type="text" name="land" value="Danmark" required="1" err="Indtast land" ></td>
</tr>
<tr>
<td>Telefon&nbsp;<span style="color:red">*</span></td><td><input type="text" name="tlf" required="1"  regexp="JSVAL_RX_ZIP" maxlength="8" minlength="8" err="Indtast telefonnummer" value="<%=session("Otelefon") %>"></td>
</tr>
<tr>
<td>Email&nbsp;<span style="color:red">*</span></td><td><input type="text" name="email" id="email" required="1" regexp="JSVAL_RX_EMAIL" err="Indtast email adresse" value="<%=session("Oemail") %>"></td>
</tr>
<tr>
<td>Password (husk dette til næste log-ind)</td><td><input type="text" readonly name="pw" value="<%response.write rndInterval%>"></td>
</tr>
<tr>
<td>Medlemsnummer (Står på bagsiden af Bindingen)</td><td><input type="text" name="medlem" value="<%=session("Omedlem") %>"></td>
</tr>
<tr>
<td>Hvis ikke medlem, hvem er da din kontaktperson ?</td><td><input type="text" name="kontakt"></td>
</tr>
<tr>
<td>Jeg har arbejdet i boden før</td><td><input type="checkbox" name="tidlArb" value="Ja"></td>
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
<!-- <tr>
<td>Jeg vil gerne være vagtleder</td><td><input type="checkbox" name="vagtleder" value="Ja"></td>
</tr>
 -->
<!-- <tr>
<td>Jeg vil helst have følgende vagter</td><td>
<input type="checkbox" name="4x6" value="Ja">&nbsp;4 x 6 timer<br>
<input type="checkbox" name="3x8" value="Ja">&nbsp;3 x 8 timer<br>
<input type="checkbox" name="2x12" value="Ja">&nbsp;2 x 12 timer
</td>
</tr> -->


<tr><td>Gruppe
<br>
Hvis I er flere som ønsker at arbejde på de samme vagter<br>
skal I enkeltvis udfylde tilmeldingen og hver især angive<br>
det samme kodeord i dette felt ('x-factor','ketchup',<br>
'snowboard' ell. lign) 
</td><td>
<input type="text" name="gruppe" size="20" value="<%=session("Ogruppe") %>">
</td>
</tr>
<tr>
<td>
Bemærkninger
<br>
Hvis du i øvrigt har bemærkninger kan dette også anføres i dette felt.
</td><td>
<textarea name="bemaerk"><%=session("Obemaerk") %></textarea>
</td></tr>
<tr>
<td colspan="2">&nbsp;&nbsp;<input type="submit" value="Gå til timevalg"></td>
</tr>

</table>
</form>
</body>
</html>
