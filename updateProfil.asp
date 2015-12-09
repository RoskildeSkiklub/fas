<%for i = 1 to 50
if instr(request.servervariables("query_string"),"periode" & i & "=") > 0 then
session("check" & i) = i
end if
next
 %>
 <%
session("bemaerk") = request("bemaerk")
'session("") = request("")
'session("") = request("")

  %>
 
<%
Function ProtectSQLInjection (TxtData)
TxtData = Replace (TxtData,"%","")
TxtData = Replace (TxtData,"'","")
TxtData = Replace (TxtData,"*","")
TxtData = Replace (TxtData,"--","")
ProtectSQLInjection = TxtData
End Function
 %>

<%
strNavn = request("navn")
strEfternavn = request("efternavn")
strTelefon = request("tlf")
'strFdato = request("fdato")
arrdate = split(request("fdato"),"-")
strFdato = arrDate(1) & "-" & arrDate(0) & "-" & arrDate(2)

response.write fdato

strAdr = request("adr")
strPostnr = request("postnr")
strBy = request("by")

'valider tomme felter
if strNavn = "" or strTelefon = "" or strFdato = "" or strAdr = "" or strPostnr = "" or strBy = "" then
session("obl") = 1
response.redirect "profil.asp"
end if
 %>


<!--#include file="conn.asp" -->

<%
'valider om samme bruger findes 2 gange
 Set rsQuery = Con.Execute("SELECT Count(tblSkiBruger.id) AS AntalOfUser FROM tblSkiBruger WHERE fdato like '" & strFdato & "' and adresse like '" & strAdr & "' and postnr like '" & strPostnr & "'")

if rsQuery("AntalOfUser") > 1 then
	session("dublet") = 1
'	response.write "SELECT Count(tblBruger.id) AS AntalOfUser FROM tblBruger WHERE navn like '" & strNavn & "' and efternavn like '" & strEfternavn & "' and adresse like '" & strAdr & "' and postnr like '" & strPostnr & "'"
	response.redirect "profil.asp"
end if


'valider 19 timer
for i = 1 to 50
if instr(request.servervariables("query_string"),"periode" & i & "=") > 0 then
Set rsQuery = Con.Execute("SELECT value from tblSkiPerioder WHERE id = " & i & "")
TimeAntal = TimeAntal + rsQuery("value")
end if
next

'Valider 9 timer
for i = 1 to 50
if instr(request.servervariables("query_string"),"periode" & i & "=") > 0 then
Set rsQuery = Con.Execute("SELECT value from tblSkiPerioder WHERE id = " & i & " and graa > 0 and graa < 3")

if rsQuery.eof = false then
TimeAntal2 = TimeAntal2 + rsQuery("value")
end if

end if
next

'valider 19 timer slut
if TimeAntal < 18 then
session("24") = cint(TimeAntal)
TimeAntal = 0
response.redirect "profil.asp"
end if

'valider 9 timer slut
if TimeAntal2 < 9 then
session("12") = cint(TimeAntal2)
TimeAntal2 = 0
'response.write session("12")
response.redirect "profil.asp"
end if


 %>



<%
sql = "UPDATE tblSkiBruger SET navn='" & ProtectSQLInjection (request("navn")) & "', adresse='" & ProtectSQLInjection (request("adr")) & "', postnr=" & ProtectSQLInjection (request("postnr")) & ", bye='" & ProtectSQLInjection (request("by")) & "', telefon='" & ProtectSQLInjection (request("tlf")) & "', tshirtstr='" & request("tshirtstr") & "', vagtleder='" & request("vagtleder") & "', fdato ='" & strFdato & "', efternavn ='" & ProtectSQLInjection (request("efternavn")) & "', land='" & ProtectSQLInjection (request("land")) & "', OPS1='" & request("OPS1") & "', OPS2='" & request("OPS2") & "', OPS3='" & request("OPS3") & "', NED1='" & request("NED1") & "', medlem='" & request("medlem") & "', kontakt='" & ProtectSQLInjection (request("kontakt")) & "',  bemaerk ='" & ProtectSQLInjection (request("bemaerk")) & "',  camping ='" & request("camping") & "',  tidlArb ='" & request("tidlArb") & "', gruppe ='" & request("gruppe") & "',  edited = '' WHERE email = '" & session("email") & "' and pw = '" & session("pw") & "'"

response.write sql

Set rsQuery = Con.Execute(sql)

'Set Con = Nothing
Set rsQuery   = Nothing

Set rsQuery = Con.Execute("DELETE FROM tblSkiBrugerTider WHERE email like '" & session("email") & "' and pw like " & session("pw") & "")

Set rsQuery   = Nothing



for i = 1 to 50
if instr(request.servervariables("query_string"),"periode" & i & "=") > 0 then
Set rsQuery = Con.Execute("INSERT INTO tblSkiBrugerTider (periodeID, email, pw) VALUES ('" & i & "', '" & session("email") & "', " & session("pw") & ")")
end if
next

Set Con = Nothing
Set rsQuery   = Nothing

session("tak") = 1

response.redirect "profil.asp"


 %>
