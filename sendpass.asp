<% @ LANGUAGE = VBSCRIPT %>

<!--#include file="conn.asp" -->

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
Set rsQuery = Con.Execute("SELECT * FROM tblSkiBruger WHERE tblSkiBruger.email = '" & ProtectSQLInjection (request("email")) & "'")

if rsQuery.eof = false then

do while rsQuery.eof = false

strPass = strPass & rsQuery("navn") & " " & rsQuery("efternavn") & ": " &rsQuery("pw") & vbcrlf

rsQuery.movenext
loop
%>



<%

	 set msg = Server.CreateOBject("JMail.Message")

   ' Angiv din afsender-adresse
   ' - den skal være oprettet på webhotellet
   msg.From = "noreply@kongebrovej.dk"
   msg.FromName = "Roskilde Skiklub"

'	msg.AddRecipientBCC ""
	msg.AddRecipient request("email")		
   ' Angiv et emne for emailen
   msg.Subject = "Password fra Roskilde Skiklub"

	msg.body =  "Hej" & vbcrlf & vbcrlf & "Der er tilknyttet følgende navne og password(s) til denne mailadresse:" & vbcrlf & vbcrlf & strPass & vbcrlf & vbcrlf & "Med venlig hilsen"  & vbcrlf & "Roskilde Skiklub"
 

		  
   if (not msg.Send("mailout2.surf-town.net")) then
    Response.write "<pre>" & msg.log & "</pre>"
   end if	
			
end if

session("pass-sendt") = 1			
response.redirect "index.asp"
			
			%>
			