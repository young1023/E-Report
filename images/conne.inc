<%
Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
DBPath = "Data Source=" & Server.MapPath("database/UOBIntranet.mdb")
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open Provider & DBPath
%>
