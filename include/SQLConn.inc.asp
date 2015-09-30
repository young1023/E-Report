<%
Set Conn = Server.CreateObject("ADODB.Connection")
StrCnn = "Provider=sqloledb; User ID=intranet; Password=intranet; Initial Catalog=UOBIntranet; Data Source=localhost"


Conn.CommandTimeout=0
Conn.ConnectionTimeout=0

Conn.Open StrCnn

%>
