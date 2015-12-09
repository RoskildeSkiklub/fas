<!--#include file="conn.asp" -->
<%
session.lcid = "1033"
'sessionvariable

session("Onavn") = request("navn")
session("Oadr") = request("adr")
session("Oby") = request("by")
session("Opostnr") = request("postnr")
session("Oemail") = request("email")
session("Otelefon") = request("tlf")
session("Omedlem") = request("medlem")
session("Oefternavn") = request("efternavn")
session("Obemaerk") = request("bemaerk")
session("Ogruppe") = request("gruppe")



' tjek på alder

Set rsQuery = Con.Execute("SELECT * FROM tblSkiData")
str18 = rsQuery("18aar")
str15 = rsQuery("15aar")

Set rsQuery = Con.Execute("SELECT CountOfid FROM Vw_SkiUnder")
strUnder = rsQuery("CountOfId")

Set rsQuery = Con.Execute("SELECT CountOfid FROM Vw_SkiAll ")
strAll = rsQuery("CountOfId")

intProcent = strUnder/strAll * 100
'tallet ændres til den procentsats der skal være grænsen for personer under 18
' dvs ved 90 skal blot 10% være over 18
if intProcent < 90 then 
	strDato2 = str15
else
	strDato2 = str18
end if


'response.write Fdato

'response.write datediff("d",strFdato, strDato2)

session.lcid = "1030"
if datediff("d",request("fdato"), strDato2) < 0 then 
session("succes") = 5
response.redirect "opret.asp"
end if
session.lcid = "1033"


strNavn = request("navn")
strEfternavn = request("efternavn")
strAdr = request("adr")
strPostnr = request("postnr")
arrdate = split(request("fdato"),"-")
strFdato = cdate(arrDate(1) & "-" & arrDate(0) & "-" & arrDate(2))

'valider om samme bruger findes 2 gange
Set rsQuery = Con.Execute("SELECT Count(tblSkiBruger.id) AS AntalOfUser FROM tblSkiBruger WHERE fdato like '" & strFdato & "' and adresse like '" & strAdr & "' and postnr like '" & strPostnr & "'")

' Set rsQuery = Con.Execute("SELECT Count(tblBruger.id) AS AntalOfUser FROM tblBruger WHERE navn like '" & strNavn & "' and efternavn like '" & strEfternavn & "' and adresse like '" & strAdr & "' and postnr like '" & strPostnr & "'")

'session("test") = rsQuery("AntalOfUser")


if rsQuery("AntalOfUser") > 0 then
	session("dublet") = 1
'	response.write "SELECT Count(tblBruger.id) AS AntalOfUser FROM tblBruger WHERE navn like '" & strNavn & "' and efternavn like '" & strEfternavn & "' and adresse like '" & strAdr & "' and postnr like '" & strPostnr & "'"
	response.redirect "opret.asp"
end if


 Function ProtectSQLInjection (TxtData)
TxtData = Replace (TxtData,"%","")
TxtData = Replace (TxtData,"'","")
TxtData = Replace (TxtData,"*","")
TxtData = Replace (TxtData,"--","")
ProtectSQLInjection = TxtData
End Function

strNavn = ProtectSQLInjection (request("navn"))
strAdresse = ProtectSQLInjection (request("adr"))
strBy = ProtectSQLInjection (request("by"))
strPostnummer = ProtectSQLInjection (request("postnr"))
strEmail = ProtectSQLInjection (request("email"))
strTelefon = ProtectSQLInjection (request("tlf"))
strVagtleder = request("vagtleder")
str3x8 = request("3x8")
str4x6 = request("4x6")
str2x12 = request("2x12")
strTshirt = request("tshirtstr")
strPW = ProtectSQLInjection (request("pw"))
'strFdato = ProtectSQLInjection (request("fdato"))
strEfternavn = ProtectSQLInjection (request("efternavn"))
strBemaerk = ProtectSQLInjection (request("bemaerk"))
strLand = ProtectSQLInjection (request("land"))
strCamping = request("camping")
strTidlArb = request("tidlArb")
strGruppe = request("gruppe")

stropsaetTelt = request("opsaetTelt")
strnedtagTelt = request("nedtagTelt")
strudlaegPall = request("udlaegPall")
strnedtagBod = request("nedtagBod")
strklarBod = request("klarBod")
strMedlem = request("medlem")
strKontakt = ProtectSQLInjection (request("kontakt"))
strNattevagt = request("nattevagt")

if strNavn = "" or strTelefon = "" or strEmail = "" or strFdato = "" or strAdresse = "" or strBy = "" or strPostnummer = "" or strLand = "" then
session("succes") = 2
response.redirect "opret.asp"
end if

%>


<%

'response.write "INSERT INTO tblBruger(navn, adresse, postnr, bye, email, telefon, vagtleder, 4x6, 3x8, 2x12, tshirtStr, pw, fdato, efternavn, bemaerk) values ( '" & strNavn & "', '" & strAdresse & "', " & strPostnummer & ", '" & strBy & "', '" & strEmail & "', '" & strTelefon & "', '" & strVagtleder & "', '" & str4x6 & "', '" & str3x8 & "', '" & str2x12 & "', '" & strTshirt & "', '" & strPW & "', '" & strFdato & "', '" & strEfternavn & "', '" & strBemaerk & "')"

'Set rsQuery = Con.Execute("INSERT INTO tblBruger(navn, adresse, postnr, bye, email, telefon, vagtleder, 4x6, 3x8, 2x12, tshirtStr, pw, fdato, efternavn, bemaerk, land, opsaetTelt, nedtagTelt, udlaegPall, nedtagBod, klarBod, nattevagt, medlem, kontakt) values ( '" & strNavn & "', '" & strAdresse & "', " & strPostnummer & ", '" & strBy & "', '" & strEmail & "', '" & strTelefon & "', '" & strVagtleder & "', '" & str4x6 & "', '" & str3x8 & "', '" & str2x12 & "', '" & strTshirt & "', '" & strPW & "', '" & strFdato & "', '" & strEfternavn & "', '" & strBemaerk & "', '" & strLand & "', '" & stropsaetTelt & "', '" & strnedtagTelt & "', '" & strudlaegPall & "', '" & strnedtagBod & "', '" & strklarBod & "', '" & strNattevagt & "', '" & strMedlem & "', '" & strKontakt & "')")

Set rsQuery = Con.Execute("INSERT INTO tblSkiBruger(navn, adresse, postnr, bye, email, telefon, vagtleder, tshirtStr, pw, fdato, efternavn, bemaerk, land, medlem, kontakt, camping, tidlArb, gruppe) values ( '" & strNavn & "', '" & strAdresse & "', " & strPostnummer & ", '" & strBy & "', '" & strEmail & "', '" & strTelefon & "', '" & strVagtleder & "', '" & strTshirt & "', '" & strPW & "', '" & strFdato & "', '" & strEfternavn & "', '" & strBemaerk & "', '" & strLand & "', '" & strMedlem & "', '" & strKontakt & "', '" & strCamping & "', '" & strTidlArb & "', '" & strGruppe & "')")

Set Con = Nothing
Set rsQuery   = Nothing

'send mail
'----------------------------------------------------

	 set msg = Server.CreateOBject("JMail.Message")

   ' Angiv din afsender-adresse
   ' - den skal være oprettet på webhotellet
   msg.From = "noreply@kongebrovej.dk"
   msg.FromName = "Roskilde Skiklub"

	msg.AddRecipientBCC "thomas@hors.dk"
	msg.AddRecipient strEmail
   ' Angiv et emne for emailen
   msg.Subject = "Password fra Roskilde Skiklub"

'			JMail.AddRecipient strEmail
			'"hors@mail.dk"
			'request.form("modtager")

		  msg.body = "Hej" & vbcrlf & vbcrlf & "Du er tilmeldt Roskilde Skiklubs Festivalgruppe på denne mail adresse, med følgende password: " & strPW & vbcrlf & vbcrlf & "Med venlig hilsen"  & vbcrlf & "Roskilde Skiklub"
 
			   ' Send mailen - ved at angive navnet på mail-serveren
   if (not msg.Send("mailout2.surf-town.net")) then
    Response.write "<pre>" & msg.log & "</pre>"
   end if	

'----------------------------------------------------


session("succes") = 1

'if session("adm") = 1 then
'response.redirect "profil-adm.asp?pw=" & strPW & "&email=" & strEmail
'else
response.redirect "login.asp?pw=" & strPW & "&email=" & strEmail
'end if
 %>