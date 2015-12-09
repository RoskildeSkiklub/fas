<%Set Con = Server.CreateObject("ADODB.Connection")
Set rsQuery   = Server.CreateObject("ADODB.Recordset")
Path ="DRIVER={SQL Server}; SERVER=sqlservernavn/IP adresse; DATABASE=databasenavn; UID=username; PWD=password"

Con.Open Path 

session.lcid = "1030"
%>