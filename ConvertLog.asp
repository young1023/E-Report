<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=AuditLog.xls"

if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if
%>

<!--#include file="include/SQLConn.inc.asp" -->

<%


'**************
'Initialisation
'**************
Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adCmdText = &H0001

ConvertLog = Request("ConvertLog")

' Create a server recordset object
Set rs = Server.CreateObject("ADODB.Recordset")


'Execute the sql
 rs.open ("Exec Retrieve_AuditLog_To_Excel '"&ConvertLog&"'") , Conn ,3,1
 

%>
<html><meta http-equiv="Content-Type" content="text/html; charset=big5">

<body>
<Head>
<STYLE TYPE="text/css">
<!--

TD 
{
  color: black;
  font-family: verdana, Garamond, Times, sans-serif;
  FONT-SIZE: 10px;
  TEXT-ALIGN: left 
}

TD.caption
{
  color: red;
  font-family: verdana, Garamond, Times, sans-serif;
  FONT-SIZE: 10px;
  TEXT-ALIGN: left 
}
-->
</STYLE>
</head>

<div align="center">

<table BORDER="1" width="98%">
<tr>
<td class=caption width="20%">Performed By</td>
<td class=caption>Description</td>
<td class=caption width="20%">Date and Time</td>
</tr>


<%
' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
%>

<tr>

<td>
<% = rs("Name") %>
</td>

<td>
<% = rs("Description") %>
</td>

<td>
<% = rs("CreateDate") %>
</td>

</tr>

<%
' Move to the next record
rs.movenext
' Loop back to the do statement
loop %>
</table>

</div>

</body>
</html>

<%
' Close and set the recordset to nothing
rs.close
set rs=nothing
%>