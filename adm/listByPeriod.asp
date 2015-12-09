<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>

<!--#include file="conn.asp" -->



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
Klik på navn for at tildele denne person vagt. Klik igen for at afmelde vagt.<br>
Resultatet vil fremgå ved personens navn.
<br>
<br>
<div style="width:100%;float:left">
<table width="200px" style="float:left">
<tr>
<td valign="top">
<strong>Opsætning 25. juni</strong>
</td>
</tr><tr>
<td valign="top" style="white-space:nowrap">
<%
Set rsQuery = Con.Execute("SELECT * FROM tblSkiBruger WHERE (((tblSkiBruger.OPS1) Like 'ja')) and ( inact <> '1' or inact Is Null ) ORDER BY id")
%>
<%
while not rsQuery.eof = true
%>
&nbsp;
<%if rsQuery("1") = "1" then%>
<strong>Vagt!</strong>
<%else %>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%end if %><a href="updateVagt.asp?id=<%=rsQuery("id") %>&action=1&status=<%=rsQuery("1") %>" target="_self"><%=rsQuery("navn") %>&nbsp;<%=rsQuery("efternavn") %></a>
- <a href="profil-adm.asp?pw=<%=rsQuery("pw") %>&email=<%=rsQuery("email") %>">Vis profil</a>
<br>

<%
rsQuery.movenext
i = i + 1
wend

Set rsQuery   = Nothing

 %>
</td>
</tr>
</table>

<table width="200px" style="float:left">
<tr>
<td valign="top">
<strong>Opsætning 26. juni</strong>
</td></tr><tr>
<td valign="top" style="white-space:nowrap">
<%
Set rsQuery = Con.Execute("SELECT * FROM tblSkiBruger WHERE (((tblSkiBruger.OPS2) Like 'ja'))and ( inact <> '1' or inact Is Null ) ORDER BY id")
%>
<%
while not rsQuery.eof = true
%>
&nbsp;
<%if rsQuery("2") = "1" then%>
<strong>Vagt!</strong>
<%else %>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%end if %><a href="updateVagt.asp?id=<%=rsQuery("id") %>&action=2&status=<%=rsQuery("2") %>" target="_self"><%=rsQuery("navn") %>&nbsp;<%=rsQuery("efternavn") %></a> - <a href="profil-adm.asp?pw=<%=rsQuery("pw") %>&email=<%=rsQuery("email") %>">Vis profil</a>
<br>

<%
rsQuery.movenext
i = i + 1
wend
Set rsQuery   = Nothing
 %>
</td>
</tr>
</table>

<table width="200px" style="float:left">
<tr>
<td valign="top">
<strong>Opsætning 27. juni</strong>
</td></tr><tr>
<td valign="top" style="white-space:nowrap">
<%
Set rsQuery = Con.Execute("SELECT * FROM tblSkiBruger WHERE (((tblSkiBruger.OPS3) Like 'ja'))and ( inact <> '1' or inact Is Null ) ORDER BY id")
%>
<%
while not rsQuery.eof = true
%>
&nbsp;
<%if rsQuery("3") = "1" then%>
<strong>Vagt!</strong>
<%else %>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%end if %><a href="updateVagt.asp?id=<%=rsQuery("id") %>&action=3&status=<%=rsQuery("3") %>" target="_self"><%=rsQuery("navn") %>&nbsp;<%=rsQuery("efternavn") %></a> - <a href="profil-adm.asp?pw=<%=rsQuery("pw") %>&email=<%=rsQuery("email") %>">Vis profil</a>
<br>

<%
rsQuery.movenext
i = i + 1
wend
Set rsQuery   = Nothing
 %>
</td>
</tr>
</table>
</div>

<div style="width:100%;float:left;margin-top:10px">
<table width="200px" style="float:left">
<tr>
<td valign="top">
<strong>Nedtagning 6. juli</strong>
</td></tr><tr>
<td valign="top" style="white-space:nowrap">
<%
Set rsQuery = Con.Execute("SELECT * FROM tblSkiBruger WHERE (((tblSkiBruger.NED1) Like 'ja'))and ( inact <> '1' or inact Is Null ) ORDER BY id")
%>
<%
while not rsQuery.eof = true
%>
&nbsp;
<%if rsQuery("4") = "1" then%>
<strong>Vagt!</strong>
<%else %>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%end if %><a href="updateVagt.asp?id=<%=rsQuery("id") %>&action=4&status=<%=rsQuery("4") %>" target="_self"><%=rsQuery("navn") %>&nbsp;<%=rsQuery("efternavn") %></a> - <a href="profil-adm.asp?pw=<%=rsQuery("pw") %>&email=<%=rsQuery("email") %>">Vis profil</a>
<br>

<%
rsQuery.movenext
i = i + 1
wend
Set rsQuery   = Nothing
 %>
</td>
</tr>
</table>
</div>

</body>
</html>
<%
Set Con = Nothing
Set rsQuery   = Nothing
 %>