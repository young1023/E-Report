<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel"  
Response.addHeader "content-disposition","attachment;filename=MarginSummary.xls"


if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if
%>

<!--#include file="include/SQLConn.inc.asp" -->

<%


'**************
'Argument handler
'**************

	
	Search_ClientFrom                  = session("ClientFrom")
	Search_ClientTo                    = session("ClientTo")
	Search_AEFrom                      = session("AEFrom")
	Search_AETo                        = session("AETo")
    Search_MinDebitBalance             = session("MinDebitBalance")
    Search_MarginExceedPercent         = session("MarginExceedPercent")
    Search_AccountType                 = session("AccountType")
   Search_AEGroup      = Request("Search_SharedGroup")
        Search_SharedGroupMember= Request("Search_SharedGroupMember")


' If User enter From value only, change the "To" value to "From"
if Search_ClientTo = "" then
   Search_ClientTo = Search_ClientFrom
end if
if Search_AETo = "" then
   Search_AETo = Search_AEFrom
end if

'**************
'Initialisation
'**************
Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adCmdText = &H0001
	

' Create a server recordset object
Set rs = Server.CreateObject("ADODB.Recordset")

 	Response.Write ("Exec Retrieve_MarginCall_To_Excel '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_MinDebitBalance&"', '"&Search_MarginExceedPercent&"', '"&Search_AccountType&"','"&Search_AEGroup&"',  '"&Search_SharedGroupMember&"' ")
 	rs.open ("Exec Retrieve_MarginCall_To_Excel '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_MinDebitBalance&"', '"&Search_MarginExceedPercent&"', '"&Search_AccountType&"',  '"&Search_AEGroup&"',  '"&Search_SharedGroupMember&"' ") , Conn ,3,1

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
  FONT-SIZE: 9px;
  TEXT-ALIGN: left 
}

TD.caption
{
  color: red;
  font-family: verdana, Garamond, Times, sans-serif;
  FONT-SIZE: 9px;
  TEXT-ALIGN: left 
}
-->
</STYLE>
</head>

<div align="center">

<table BORDER="1" width="98%">
			<tr bgcolor="#ADF3B6">
			   <td width="10%">Client No.<br>客戶編號</a></td>
			   <td width="18%">Client Name<br>客戶名稱</a></td>
			   <td width="5%">Currency<br>貨幣</a></td>
			   <td width="10%">Cash Balance to be settled (HKD)</a></td>
			   <td width="5%">AE Code<br>經紀編號</a></td>
			   <td width="12%">AE Name<br>經紀名稱</td>
			   <td width="12%">Portfilio Mkt Value (HKD)<br>組合總值(港元)</a></td> 
			   <td width="12%">A/C Bal (Due Amt)</a></td> 
			   <td width="12%">Margin %<br>按倉比率</a></td> 
			   <td width="12%">Loss<br>損失</a></td> 
			   <td width="12%">Margin Value<br>按倉價值</a></td> 
			   <td width="12%">Accrued Int</a></td> 
			</tr>

<%
' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
%>

<tr bgcolor="#ADF3B6" align="center">
			<td><% =rs("CLNTCODE")  %></td>
			<td><% =rs("CLNTNAME")  %></td>
			<td><% =rs("CCY")  %></td>
			<td><% =rs("CURRENCY")  %></td>
			<td><% =rs("CLNTAECODE")  %></td>
			<td><% =rs("CLNTAENAME")  %></td>
			<td><% =rs("PORTFILIO")  %></td>
			<td><% =rs("BALANCE")  %></td>
			<td><% =rs("MARGINPERCENT")  %></td>
			<td><% =rs("LOSS")  %></td>
			<td><% =rs("MARGINVALUE")  %></td>
			<td><% =rs("ACCUREDINT")  %></td>

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