<%
'Set Conn = Server.CreateObject("ADODB.Connection")
'StrCnn = "Data Source=PROLIANT;Initial Catalog=UOBIntranet;Integrated Security=True"
'Conn.Open StrCnn

Dim strSQL
set Global_DBConnection=Server.CreateObject("ADODB.Connection")
DSN="Driver={SQL Server};Server=192.168.1.100;Database=UOBIntranet;UID=sa;PWD="
Global_DBConnection.Open(DSN)


%>
