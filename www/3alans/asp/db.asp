<%
dim adocon
Set adocon = Server.CreateObject ("ADODB.Connection")
adocon.Mode = 3
adocon.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="&Server.Mappath("../../")&"\database\3alans.mdb"
%>