<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>
 
 <%
'for i = 1 to 50
'if instr(request.servervariables("query_string"),"periode" & i & "=") > 0 then
'session("check" & i) = i
'end if
'next
 %>
 <%
'session("bemaerk") = request("bemaerk")
'session("opsaetTelt") = request("opsaetTelt")
'session("klarBod") = request("klarBod")
'session("nattevagt") = request("nattevagt")
'session("udlaegPall") = request("udlaegPall")
'session("nedtagTelt") = request("nedtagTelt")
'session("nedtagBod") = request("nedtagBod")
'session("") = request("")
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

arrdate = split(request("fdato"),"-")
strFdato = arrDate(1) & "-" & arrDate(0) & "-" & arrDate(2)
session.lcid = "1030"
 %>
 
<%
'strNavn = request("navn")
'strEfternavn = request("efternavn")
'strTelefon = request("tlf")
'strFdato = request("fdato")
'strAdr = request("adr")
'strPostnr = request("postnr")
'strBy = request("by")

'valider tomme felter
'if strNavn = "" or strTelefon = "" or strFdato = "" or strAdr = "" or strPostnr = "" or strBy = "" then
'session("obl") = 1
'response.redirect "profil.asp"
'end if
 %>


<!--#include file="conn.asp" -->

<%
'valider om samme bruger findes 2 gange
' Set rsQuery = Con.Execute("SELECT Count(tblBruger.id) AS AntalOfUser FROM tblBruger WHERE fdato like '" & strFdato & "' and adresse like '" & strAdr & "' and postnr like '" & strPostnr & "'")

'if rsQuery("AntalOfUser") > 1 then
'session("dublet") = 1
'	response.write "SELECT Count(tblBruger.id) AS AntalOfUser FROM tblBruger WHERE navn like '" & strNavn & "' and efternavn like '" & strEfternavn & "' and adresse like '" & strAdr & "' and postnr like '" & strPostnr & "'"
'	response.redirect "profil.asp"
'end if




'valider overlap
'for i = 1 to 50
'if instr(request.servervariables("query_string"),"periode" & i & "=") > 0 then
'Set rsQuery = Con.Execute("SELECT overlap from tblPerioder WHERE id = " & i & "")
'OverlapAntal = OverlapAntal & rsQuery("overlap")
'session("testOverlap") = OverlapAntal
'end if
'next

'valider 24 timer
'for i = 1 to 50
'if instr(request.servervariables("query_string"),"periode" & i & "=") > 0 then
'Set rsQuery = Con.Execute("SELECT value from tblPerioder WHERE id = " & i & "")
'TimeAntal = TimeAntal + rsQuery("value")
'end if
'next

'valider 12 timer
'for i = 1 to 50
'if instr(request.servervariables("query_string"),"periode" & i & "=") > 0 then
'Set rsQuery = Con.Execute("SELECT value from tblPerioder WHERE id = " & i & " and graa > 0 and graa < 3")

'if rsQuery.eof = false then
'TimeAntal2 = TimeAntal2 + rsQuery("value")
'end if

'end if
'next

'if instr(OverlapAntal,"ab") > 0 or instr(OverlapAntal,"cd") > 0 or instr(OverlapAntal,"ef") > 0 or instr(OverlapAntal,"gh") > 0 then
'session("overlap") = 1
'OverlapAntal = ""
'response.redirect "profil.asp"
'end if

'if TimeAntal <> 24 then
'session("24") = cint(TimeAntal)
'TimeAntal = 0
'response.redirect "profil.asp"
'end if


'if TimeAntal2 < 12 then
'session("12") = cint(TimeAntal2)
'TimeAntal2 = 0
'response.write session("12")
'response.redirect "profil.asp"
'end if


 %>



<%
'Set rsQuery = Con.Execute("UPDATE tblBruger SET mnummer='" & request("mnummer") & "', navn='" & request("navn") & "', adresse='" & request("adr") & "', postnr=" & request("postnr") & ", bye='" & request("by") & "', telefon='" & request("tlf") & "', tshirtstr='" & request("tshirtstr") & "', vagtleder='" & request("vagtleder") & "', 4x6='" & request("4x6") & "', 3x8='" & request("3x8") & "', 2x12='" & request("2x12") & "', fdato ='" & request("fdato") & "', efternavn ='" & request("efternavn") & "', land='" & request("land") & "', opsaetTelt='" & request("opsaetTelt") & "', nedtagTelt='" & request("nedtagTelt") & "', udlaegPall='" & request("udlaegPall") & "', nedtagBod='" & request("nedtagBod") & "', klarBod='" & request("klarBod") & "', nattevagt='" & request("nattevagt") & "', medlem='" & request("medlem") & "', kontakt='" & request("kontakt") & "',  andreBesk ='" & request("andreBesk") & "',  bemaerk ='" & request("bemaerk") & "', camping ='" & request("camping") & "' WHERE email like '" & request("email") & "' and pw like '" & request("pw") & "'")

'Set rsQuery = Con.Execute("UPDATE tblBruger SET navn='" & ProtectSQLInjection (request("navn")) & "', adresse='" & ProtectSQLInjection (request("adr")) & "', postnr=" & ProtectSQLInjection (request("postnr")) & ", bye='" & ProtectSQLInjection (request("by")) & "', telefon='" & ProtectSQLInjection (request("tlf")) & "', tshirtstr='" & request("tshirtstr") & "', vagtleder='" & request("vagtleder") & "', 4x6='" & request("4x6") & "', 3x8='" & request("3x8") & "', 2x12='" & request("2x12") & "', fdato ='" & ProtectSQLInjection (request("fdato")) & "', efternavn ='" & ProtectSQLInjection (request("efternavn")) & "', land='" & ProtectSQLInjection (request("land")) & "', opsaetTelt='" & request("opsaetTelt") & "', nedtagTelt='" & request("nedtagTelt") & "', udlaegPall='" & request("udlaegPall") & "', nedtagBod='" & request("nedtagBod") & "', klarBod='" & request("klarBod") & "', nattevagt='" & request("nattevagt") & "', nt='" & request("nt") & "', nf='" & request("nf") & "', nl='" & request("nl") & "', ns='" & request("ns") & "', medlem='" & request("medlem") & "', kontakt='" & ProtectSQLInjection (request("kontakt")) & "',  bemaerk ='" & ProtectSQLInjection (request("bemaerk")) & "',  camping ='" & request("camping") & "',  tidlArb ='" & request("tidlArb") & "',  edited = '" & request("edited") & "' WHERE email like '" & request("email") & "' and pw like '" & request("pw") & "'")

sql = "UPDATE tblSkiBruger SET mnummer='" & request("mnummer") & "', navn='" & ProtectSQLInjection (request("navn")) & "', adresse='" & ProtectSQLInjection (request("adr")) & "', postnr=" & ProtectSQLInjection (request("postnr")) & ", bye='" & ProtectSQLInjection (request("by")) & "', telefon='" & ProtectSQLInjection (request("tlf")) & "', tshirtstr='" & request("tshirtstr") & "', vagtleder='" & request("vagtleder") & "', fdato ='" & strFdato & "', efternavn ='" & ProtectSQLInjection (request("efternavn")) & "', land='" & ProtectSQLInjection (request("land")) & "', OPS1='" & request("OPS1") & "', OPS2='" & request("OPS2") & "', OPS3='" & request("OPS3") & "', NED1='" & request("NED1") & "', medlem='" & request("medlem") & "', kontakt='" & ProtectSQLInjection (request("kontakt")) & "',  bemaerk ='" & ProtectSQLInjection (request("bemaerk")) & "',  camping ='" & request("camping") & "',  tidlArb ='" & request("tidlArb") & "', gruppe ='" & request("gruppe") & "', arbopg ='" & request("arbopg") & "', andreBesk ='" & request("andreBesk") & "', edited = '" & request("edited") & "' WHERE email = '" & request("email") & "' and pw = '" & request("pw") & "'"

'response.write sql

Set rsQuery = Con.Execute(sql)

'Set Con = Nothing
Set rsQuery   = Nothing

Set rsQuery = Con.Execute("DELETE FROM tblSkiBrugerTider WHERE email like '" & request("email") & "' and pw like " & request("pw") & "")

Set rsQuery   = Nothing



for i = 1 to 50
if instr(request.servervariables("query_string"),"periode" & i & "=") > 0 then
Set rsQuery = Con.Execute("INSERT INTO tblSkiBrugerTider (periodeID, email, pw) VALUES ('" & i & "', '" & request("email") & "', " & request("pw") & ")")
end if
next

Set Con = Nothing
Set rsQuery   = Nothing

'session("tak") = 1

response.redirect "profil-adm.asp?pw=" & request("pw") & "&email=" & request("email")


 %>
