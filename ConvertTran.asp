<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=Transaction.xls"


if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if
%>

<!--#include file="include/SQLConn.inc.asp" -->

<%


'**************
'Argument handler
'**************

	Search_From_Day = request("From_Day")
	Search_From_Month = request("From_Month")
	Search_From_Year = request("From_Year")
	Search_To_Day = request("To_Day")
	Search_To_Month = request("To_Month")
	Search_To_Year = request("To_Year")
    Search_Transaction_Type = Request("Search_Transaction_Type")
    Search_Market           = Request("Search_Market")
    Search_Currency         = Request("Search_Currency")
    Search_Instrument       = Request("Search_Instrument")
    Search_Amount_type      = Request("Search_Amount_type")
    Search_SharedGroup      = Request("Search_SharedGroup")
    Search_SharedGroupMember= Request("Search_SharedGroupMember")

		
	Search_ClientFrom   = session("ClientFrom")
	Search_ClientTo     = session("ClientTo")
	Search_AEFrom       = session("AEFrom")
	Search_AETo         = session("AETo")


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

 	'Response.Write ("Exec Retrieve_TransactionHistory_To_Excel '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Transaction_Type&"', '"&Search_Market&"', '"&Search_Currency&"', '"&Search_instrument&"', '"&Search_Amount_type&"', '"&Search_SharedGroup&"', '"&Search_SharedGroupMember&"' ")
 	rs.open ("Exec Retrieve_TransactionHistory_To_Excel '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Transaction_Type&"', '"&Search_Market&"', '"&Search_Currency&"', '"&Search_instrument&"', '"&Search_Amount_type&"', '"&Search_SharedGroup&"', '"&Search_SharedGroupMember&"' ") , Conn , 3,1

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
   <td width="8%">Trade Date<br>交易日期</td>
   <td width="10%">Client No.<br>客戶編號</td>
   <td width="10%">Client Name<br>客戶名稱</td>
   <td width="10%">Value Date<br>評價日期</td>
   <td width="8%">Trans Type<br>交易類型</td>
   <td width="8%">Market<br>市場</td>
   <td width="8%">Curr<br>貨幣</td>
   <td width="9%">Instrument<br>股票號碼</td> 
   <td width="12%">Instrument Name<br>股票名稱</td> 
   <td width="9%">Quantity<br>數量</td> 
   <td width="17%">Price<br>價錢</td> 
   <td width="30%">Amount<br>總值</td> 
   <td width="17%">Remark<br>備註</td> 
</tr>

<%
' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
%>

<tr bgcolor="#FFFFCC"> 
   <td width="8%"><% = rs("TradeDate") %></td>
   <td width="13%"><% = rs("Clnt") %></td>
   <td width="10%"><% = rs("ClntName") %>　</td>
   <td width="10%"><% = rs("ValueDate") %>　</td>
   <td width="8%"><% = rs("TradeType") %>　</td>
   <td width="8%"><% = rs("Market") %>　</td>
   <td width="8%"><% = rs("Ccy") %>　</td>
   <td width="9%"><% = rs("Instrument") %>　</td> 
   <td width="12%"><% = rs("InstrumentDesc") %>　</td> 

<! -- copy codes for strquantity and strprice from transactionhistory.asp by gary on 11 Jan 2011 -->
  <td width="9%"><%  if rs("strquantity") <> "" then
   	mystr = replace(rs("strquantity"), chr(13), " ")
   		response.write	left(mystr, len(mystr)-1) 
   			end if%>　</td> 
 <td width="17%"><%  if rs("strPrice") <> "" then
   		mystr = replace(rs("strPrice"), chr(13), " ")
   	response.write	left(mystr, len(mystr)-1) 
   		end if%>　</td> 

   <td width="17%"><% = formatnumber(rs("Amount"),2,-1,-1) %>　</td> 
   <td width="17%"><% = rs("Remark") %>　</td> 
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