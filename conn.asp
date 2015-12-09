<%Set Con = Server.CreateObject("ADODB.Connection")
Set rsQuery   = Server.CreateObject("ADODB.Recordset")
'Path = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
'Path = Path & Server.MapPath("../../../../database/ski.mdb") & ""
'Path = Path & Server.MapPath("volley.mdb") & ""
Path ="DRIVER={SQL Server}; SERVER=sqlservernavn/IP adresse; DATABASE=databasenavn; UID=username; PWD=password"

Con.Open Path 

session.lcid = "1030"
%>