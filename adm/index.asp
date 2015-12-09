<%
if request("logaf") = "1" then
session.abandon
end if

 %>
<%
if request("pw") <> "" and request("email") <> "" then
 %>


<%

if request("pw") = "svanevej" and request("email") = "thomas" then
session("adm") = "1"
session.timeout = 120
response.redirect "list.asp"
else
session("adm") = "0"
end if

end if %>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
	<title>Log ind</title>
<LINK HREF="adm.css" REL="stylesheet" TYPE="text/css">
</head>


<body>
<div style="width:450px">
<img src="../logo.jpg" height="81" width="100" align="right">
<h1>Roskilde Skiklub</h1>
<h2>Roskilde Festival</h2>
</div>
<%
if session("adm") = "0" then
 %>
Brugernavn eller password var forkert. Prøv igen.
<br>
<br>
<%end if 
session("adm") = ""
%>
 
<form id="logon" action="index.asp" method="get">
Email: <input type="text" name="email"><br>
Password: <input type="password" name="pw"><br>
<input type="submit" value="Log på">
</form>
</body>
</html>
