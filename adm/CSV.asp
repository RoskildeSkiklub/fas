<!--#include file="conn.asp" -->
id;navn;efternavn;adresse;postnr;by;land;email;telefon;tshirtstr;fødselsdato;medlem;kontakt;camping;medlemsnummer;båndnummer;gruppe;antal_timer;bmrk<%=vbcrlf %><br>
<%
sql = "SELECT tblSkiBruger.*, Vw_SkiAntalTimer.SumOfvalue FROM Vw_SkiAntalTimer RIGHT JOIN tblSkiBruger ON Vw_SkiAntalTimer.pw = tblSkiBruger.pw"
Set rsQuery1 = Con.Execute(sql)    '"SELECT * FROM tblSkiBruger")
 %>

 <%
 while rsQuery1.eof = false
 %>
 <%=rsQuery1("id") %>; <%=rsQuery1("navn") %>; <%=rsQuery1("efternavn") %>; <%=rsQuery1("adresse") %>; <%=rsQuery1("postnr") %>; <%=rsQuery1("bye") %>; <%=rsQuery1("land") %>; <%=rsQuery1("email") %>; <%=rsQuery1("telefon") %>; <%=rsQuery1("tshirtstr") %>; <%=rsQuery1("fdato") %>; <%=rsQuery1("medlem") %>; <%=rsQuery1("kontakt") %>; <%=rsQuery1("camping") %>; <%=rsQuery1("mnummer") %>; <%=rsQuery1("bnummer") %>; <%=rsQuery1("gruppe") %>; <%=rsQuery1("sumOfValue") %>; <%=rsQuery1("bemaerk") %>;<%=vbcrlf %><br>



 <%
 rsQuery1.movenext
 wend
 
 
  %>
