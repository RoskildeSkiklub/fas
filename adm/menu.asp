<%
if session("adm") <> "1" then
response.redirect "index.asp"
end if
 %>
 <SCRIPT LANGUAGE="JavaScript">

/*
Created by Randy Bennet http://home.thezone.net/~rbennett/utility/javahead.htm
Featured on JavaScript Kit (http://javascriptkit.com)
For this and over 400+ free scripts, visit http://javascriptkit.com
*/

function setVariables() {
if (document.layers) {
v=".top=";
dS="document.";
sD="";
y="window.pageYOffset";
}
else if (document.all){
v=".pixelTop=";
dS="";
sD=".style";
y="document.body.scrollTop";
}
else if (document.getElementById){
y="window.pageYOffset";
}
}
function checkLocation() {
object="object1";
yy=eval(y);
if (document.getElementById)
document.getElementById("object1").style.top=yy
else
eval(dS+object+sD+v+yy)
setTimeout("checkLocation()",10);
}
</script>

<div id="object1" style="position:absolute; visibility:show; left:15px; top:15px; z-index:5;background-color:transparent">

<table width=750 border=0 cellspacing=1 cellpadding=5 style="background-color:#AAAAAA">
<tr>
<td bgcolor="#FFFFFF"><a href="index.asp?logaf=1">Log af</a></td>
<td bgcolor="#FFFFFF"><a href="list.asp">Vis alle personer</a></td>
<td bgcolor="#FFFFFF"><a href="listbyperiod.asp">Vis specielle vagter</a></td>
<td bgcolor="#FFFFFF"><a href="listall.asp">Vis personer på vagter</a></td>
<td bgcolor="#FFFFFF"><a href="listallgraph.asp">Vis personer på vagter - Graf</a></td>
<td bgcolor="#FFFFFF"><a href="listbyvagtleder.asp">Vis vagtledere</a></td>
<td bgcolor="#FFFFFF"><a href="oversigter.asp">Oversigter</a></td>
</tr>
</table>
</div>