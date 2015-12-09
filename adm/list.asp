<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>

<!--#include file="conn.asp" -->

<%
Set rsQuery = Con.Execute("SELECT * FROM tblSkiData")
str18 = dateadd("yyyy", 18, rsQuery("18aar"))
'str18 = rsQuery("18aar")
%>


<%
if request("sort1") <> "" then
session("sort1") = request("sort1")
end if

if request("sort2") <> "" then
session("sort2") = request("sort2")
end if

if request("sort3") <> "" then
session("sort3") = request("sort3")
end if

if request("order1") <> "" then
session("order1") = request("order1")
end if

if request("order2") <> "" then
session("order2") = request("order2")
end if

if request("order3") <> "" then
session("order3") = request("order3")
end if
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
<br>
<br>
Ved klik på navn linkes til personen profil, som så herefter kan revideres.
<br>
Vælg herunder for at sortere:<br>

	<form name="order" method="get" action="list.asp">
<table>
<tr>
<td>
<select name="order1">
<option value="9" <%if session("order1") = 9 then response.write "selected" %>>Medarbejdernummer</option>
<option value="8" <%if session("order1") = 8 then response.write "selected" %>>Tilmeldingstidspunkt</option>
<option value="1" <%if session("order1") = 1 then response.write "selected" %> >Fornavn</option>
<option value="4" <%if session("order1") = 4 then response.write "selected" %>>Efternavn</option>
<option value="2" <%if session("order1") = 2 then response.write "selected" %>>Timeantal</option>
<option value="3" <%if session("order1") = 3 then response.write "selected" %>>Vagtantal</option>
<option value="5" <%if session("order1") = 5 then response.write "selected" %>>Fødselsår</option>
<option value="6" <%if session("order1") = 6 then response.write "selected" %>>Camping</option>
<option value="7" <%if session("order1") = 7 then response.write "selected" %>>Aktiv/Inaktiv</option>
<option value="10" <%if session("order1") = 10 then response.write "selected" %>>Behandlet</option>
<option value="11" <%if session("order1") = 11 then response.write "selected" %>>Gruppe</option>
</select>
<br>
<input name="sort1" value="asc" <%if session("sort1") = "asc" then response.write "checked" %> type="radio">&nbsp;Stigende
<br>
<input name="sort1" value="desc" <%if session("sort1") = "desc" then response.write "checked" %> type="radio">&nbsp;Faldende
<br>
	</td>
<td>
<select name="order2">
<option value="9" <%if session("order2") = 9 then response.write "selected" %>>Medarbejdernummer</option>
<option value="8" <%if session("order2") = 8 then response.write "selected" %>>Tilmeldingstidspunkt</option>
<option value="1" <%if session("order2") = 1 then response.write "selected" %> >Fornavn</option>
<option value="4" <%if session("order2") = 4 then response.write "selected" %>>Efternavn</option>
<option value="2" <%if session("order2") = 2 then response.write "selected" %>>Timeantal</option>
<option value="3" <%if session("order2") = 3 then response.write "selected" %>>Vagtantal</option>
<option value="5" <%if session("order2") = 5 then response.write "selected" %>>Fødselsår</option>
<option value="6" <%if session("order2") = 6 then response.write "selected" %>>Camping</option>
<option value="7" <%if session("order2") = 7 then response.write "selected" %>>Aktiv/Inaktiv</option>
<option value="10" <%if session("order2") = 10 then response.write "selected" %>>Behandlet</option>
<option value="11" <%if session("order2") = 11 then response.write "selected" %>>Gruppe</option>
</select>
<br>
<input name="sort2" value="asc" <%if session("sort2") = "asc" then response.write "checked" %> type="radio">&nbsp;Stigende
<br>
<input name="sort2" value="desc" <%if session("sort2") = "desc" then response.write "checked" %> type="radio">&nbsp;Faldende
<br>
	</td>
	<td>
<select name="order3">
<option value="9" <%if session("order3") = 9 then response.write "selected" %>>Medarbejdernummer</option>
<option value="8" <%if session("order3") = 8 then response.write "selected" %>>Tilmeldingstidspunkt</option>
<option value="1" <%if session("order3") = 1 then response.write "selected" %> >Fornavn</option>
<option value="4" <%if session("order3") = 4 then response.write "selected" %>>Efternavn</option>
<option value="2" <%if session("order3") = 2 then response.write "selected" %>>Timeantal</option>
<option value="3" <%if session("order3") = 3 then response.write "selected" %>>Vagtantal</option>
<option value="5" <%if session("order3") = 5 then response.write "selected" %>>Fødselsår</option>
<option value="6" <%if session("order3") = 6 then response.write "selected" %>>Camping</option>
<option value="7" <%if session("order3") = 7 then response.write "selected" %>>Aktiv/Inaktiv</option>
<option value="10" <%if session("order3") = 10 then response.write "selected" %>>Behandlet</option>
<option value="11" <%if session("order3") = 11 then response.write "selected" %>>Gruppe</option>
</select>
<br>
<input name="sort3" value="asc" <%if session("sort3") = "asc" then response.write "checked" %> type="radio">&nbsp;Stigende
<br>
<input name="sort3" value="desc" <%if session("sort3") = "desc" then response.write "checked" %> type="radio">&nbsp;Faldende
<br>

	</td>
</tr>
</table>
<input type="submit" value="Sorter"></form>



<br>
<%

select case session("order1")
case 1
strOrder1 = "tblSkiBruger.navn"
case 2 
strOrder1 = "Vw_SkiAntalTimer.SumOfvalue"
case 3 
strOrder1 = "Vw_SkiAntalVagter.AntalOfPeriodeID"
case 4
strOrder1 = "tblSkiBruger.efternavn"
case 5
strOrder1 = "cdate(tblSkiBruger.fdato)"
case 6
strOrder1 = "tblSkiBruger.camping"
case 7
strOrder1 = "tblSkiBruger.inact"
case 8
strOrder1 = "tblSkiBruger.id"
case 9
strOrder1 = "tblSkiBruger.mnummer"
case 10
strOrder1 = "tblSkiBruger.edited"
case 11
strOrder1 = "tblSkiBruger.gruppe"
case else
strOrder1 = "tblSkiBruger.navn"
end select

select case session("order2")
case 1
strOrder2 = "tblSkiBruger.navn"
case 2 
strOrder2 = "Vw_SkiAntalTimer.SumOfvalue"
case 3 
strOrder2 = "Vw_SkiAntalVagter.AntalOfPeriodeID"
case 4
strOrder2 = "tblSkiBruger.efternavn"
case 5
strOrder2 = "cdate(tblSkiBruger.fdato)"
case 6
strOrder2 = "tblSkiBruger.camping"
case 7
strOrder2 = "tblSkiBruger.inact"
case 8
strOrder2 = "tblSkiBruger.id"
case 9
strOrder2 = "tblSkiBruger.mnummer"
case 10
strOrder2 = "tblSkiBruger.edited"
case 11
strOrder1 = "tblSkiBruger.gruppe"
case else
strOrder2 = "Vw_SkiAntalVagter.AntalOfPeriodeID"
end select

select case session("order3")
case 1
strOrder3 = "tblSkiBruger.navn"
case 2 
strOrder3 = "Vw_SkiAntalTimer.SumOfvalue"
case 3 
strOrder3 = "Vw_SkiAntalVagter.AntalOfPeriodeID"
case 4
strOrder3 = "tblSkiBruger.efternavn"
case 5
strOrder3 = "cdate(tblSkiBruger.fdato)"
case 6
strOrder3 = "tblSkiBruger.camping"
case 7
strOrder3 = "tblSkiBruger.inact"
case 8
strOrder3 = "tblSkiBruger.id"
case 9
strOrder3 = "tblSkiBruger.mnummer"
case 10
strOrder3 = "tblSkiBruger.edited"
case 11
strOrder1 = "tblSkiBruger.gruppe"
case else
strOrder3 = "tblSkiBruger.efternavn"
end select

'Set rsQuery = Con.Execute("SELECT tblBruger.pw, tblBruger.7 , tblBruger.navn,tblBruger.fdato, tblBruger.Efternavn, tblBruger.opsaetTelt, tblBruger.adresse, tblBruger.bemaerk, tblBruger.postnr, tblBruger.bye, tblBruger.email, tblBruger.telefon, tblBruger.vagtleder, tblBruger.[4x6], tblBruger.[3x8], tblBruger.[2x12], tblBruger.tshirtStr, qryAntalTimer.SumOfvalue, qryAntalVagter.AntalOfPeriodeID FROM qryAntalVagter RIGHT JOIN (qryAntalTimer RIGHT JOIN tblBruger ON qryAntalTimer.pw = tblBruger.pw) ON qryAntalVagter.pw = tblBruger.pw ORDER BY " & strOrder & " " & request("sort") & "")



Set rsQuery = Con.Execute("SELECT tblSkiBruger.*, Vw_SkiAntalTimer.SumOfvalue, Vw_SkiAntalVagter.AntalOfPeriodeID, tblSkiBruger.bemaerk, tblSkiBruger.andrebesk, tblSkiBruger.gruppe, tblSkiBruger.vente FROM Vw_SkiAntalVagter RIGHT JOIN (Vw_SkiAntalTimer RIGHT JOIN tblSkiBruger ON Vw_SkiAntalTimer.pw = tblSkiBruger.pw) ON Vw_SkiAntalVagter.pw = tblSkiBruger.pw ORDER BY " & strOrder1 & " " & session("sort1") & ", " & strOrder2 & " " & session("sort2") & ", " & strOrder3 & " " & session("sort3") & " ")


 %>
<table border="0" cellpadding="4" cellspacing="1">
<%
i = 1
while not rsQuery.eof = true

%>
<tr>
<td  align="center"  style="<%if i mod 2 = 0 then%>background-color: #dddddd <%end if %>;font-size:18px;<%if rsQuery("inact") = "1" then%>color:red;font-weight:bold<%end if %>"><%=i %>
</td>
<td align="left" valign="top" <%if i mod 2 = 0 then%> style="background-color: #dddddd <%end if %>;white-space:nowrap">
Timer:
<%
if isnull(rsQuery("SumOfvalue")) then%>
<strong>Ingen</strong>
<%elseif rsQuery("SumOfvalue") <> 24 then
 %>
 <strong><%=rsQuery("SumOfvalue") %></strong>
<% 
else
response.write rsQuery("SumOfvalue")
end if
%><br>
Vagter: <%=rsQuery("AntalOfPeriodeID") %>
<br>
ID: <strong><%=rsQuery("id") %></strong>
<br>
Status:&nbsp;<a href="updateInactive.asp?id=<%=rsQuery("id")%>&action=inact&status=<%=rsQuery("inact") %>" target="_self"><%if rsQuery("inact") = "1" then%>Inaktiv<%else %>Aktiv<%end if %></a>
<br>
Behandlet: <strong><%=rsQuery("edited") %></strong>
<br>
Tshirt:&nbsp;<i><%=rsQuery("tshirtStr") %> </i>
<br>
PW:&nbsp;<i><%=rsQuery("pw") %> </i>

 </td>
 <td <%if i mod 2 = 0 then%> style="background-color: #dddddd" <%end if %>><a target="_self" href="profil-adm.asp?pw=<%=rsQuery("pw") %>&email=<%=rsQuery("email") %>" >
<%if rsQuery("efternavn") <> "" then %>
<%=rsQuery("navn") %>&nbsp;<%=rsQuery("efternavn") %>
<%end if %>
</a>&nbsp;-&nbsp;
<% 
    ' use DateSerial(y,m,d) to avoid locale issues 
	 
	 
	 
    date1 = DateSerial(year(rsQuery("fdato")), month(rsQuery("fdato")), day(rsQuery("fdato")))  
    date2 = DateSerial(year(now()), month(now()), day(now()) )
 
    ' make sure we have a valid date! 
    if date2 >= date1 then 
 
        ' determine if leapYearBaby 
        if month(date1) = 2 and day(date1) = 29 then 
            leapBaby = true 
        end if      
 
        ' get absolute number of years 
        ageInYears = cint(datediff("YYYY", date1, date2)) 
 
        ' get date1's month and day in terms of date2's year 
        date1alt = dateadd("yyyy", ageInYears, date1) 
 
 
        if date1alt > date2 then 
            ' their birthday hasn't hit yet in date2's year 
            ageInYears = ageInYears - 1 
        end if 
 
        if leapBaby = true then 
            ' need to format output slightly 
            yearsPassed = ageInYears 
            ageInYears = ageInYears \ 4 
        end if 
 
        response.write "" & ageInYears 
        if leapBaby then response.write " (" & yearsPassed & " siden fødsel)" 
    else 
        response.write "Invalid date." 
    end if 
%>

<%
'response.write datediff("yyyy",rsQuery("fdato"),cdate(str18),2,1) 
%>
år
<br>
<%=rsQuery("Adresse") %><br>
<%=rsQuery("postnr") %>&nbsp;<%=rsQuery("bye") %><br>
<%=rsQuery("telefon") %><br>
<a href="mailto:<%=rsQuery("email") %>"><%=rsQuery("email") %></a>
</a><br>
</td>
<td <%if i mod 2 = 0 then%> style="background-color: #dddddd" <%end if %>><%=rsQuery("bemaerk") %><hr>
<%=rsQuery("andreBesk") %>
<hr>
<%=rsQuery("gruppe") %>
</td>
<td width="200" <%if i mod 2 = 0 then%> style="background-color: #dddddd" <%end if %> align="right">
<form action="updateMnummer-adm.asp" method="get">
<input type="hidden" value="<%=rsQuery("email") %>" name="email">
<input type="hidden" value="<%=rsQuery("pw") %>" name="pw">
Medarbejdernummer:<input type="text" value="<%=rsQuery("mnummer") %>" name="mnummer" size="4"><br>
Båndnummer:<input type="text" value="<%=rsQuery("bnummer") %>" name="bnummer" size="4"><br>
Venteliste: <input type="checkbox" name="vente" value="1" <%if rsQuery("vente")="1" then response.write "checked" end if%>>
<br>
<input type="submit" value="Gem">
</form>
<form action="deleteUser.asp" method="get">
<input type="hidden" value="<%=rsQuery("email") %>" name="email">
<input type="hidden" value="<%=rsQuery("pw") %>" name="pw">
<input type="text" value="skriv delete her" name="delete" size="14"><input type="submit" value="Slet">
</form>
</td>
</tr>


<%
rsQuery.movenext
i = i + 1

wend


 %>
</table>

</body>
</html>
<%
Set Con = Nothing
Set rsQuery   = Nothing
 %>