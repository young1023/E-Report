<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=MemberList.xls"

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

' Create a server recordset object
Set rs = Server.CreateObject("ADODB.Recordset")


'Execute the sql
 rs.open ("Exec Retrieve_Member_To_Excel") , Conn ,3,1


       ' Get the number of day for the password expired in the system
   
        sql2 = "Select SettingValue From SystemSetting Where SettingName = 'PasswordMaximumAge'"
   
        Set Rs2 = Conn.Execute(sql2)
   
        ExpiredAge = Rs2("SettingValue")

%>


<html>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
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
<td class=caption width="20%">Name</td>
<td class=caption>Login Name</td>
<td class=caption>Department</td>
<td class=caption>Email</td>
<td class=caption>User Right</td>
<td class=caption>Branch</td>
<td class=caption>Shared Group</td>
<td class=caption width="20%">Password Expired Date</td>
<td class=caption width="20%">Created Date</td>

</tr>


<%
' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
%>

<tr>

<td>
<% = rs("MemberName") %>
</td>

<td>
<% = rs("LoginName") %>
</td>

<td>
<% = rs("Dept") %>
</td>

<td>
<% = rs("Email") %>
</td>


<td>
<% = rs("levelName") %>
</td>

<td>
<% = rs("GroupName") %>
</td>

<td>
<%

   ' Show Share Group
   
   Sql1 = " Select * From SharedGroup s Join UserGroup u on s.SharedGroupID = u.GroupID "
   
   Sql1 = Sql1 & " and u.sharing = 1 and  s.MemberID = "&rs("MemberID")
   
   Set Rs1 = Conn.Execute(Sql1)
      
   If Not Rs1.EoF Then
   Rs1.MoveFirst
   Do While Not Rs1.EoF
   response.write Rs1("Name")&"<br>"
   Rs1.MoveNext
   Loop
   End If

%>
</td>

<td>
<% = dateadd("d", ExpiredAge, datevalue(rs("LastPasswordChangeDate"))) %>
</td>

<td>
<% = datevalue(rs("CreationDate")) %>
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