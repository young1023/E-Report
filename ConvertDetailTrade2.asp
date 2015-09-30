<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=DetailTrade.xls"


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
    Search_Market = request("Search_Market")
    Search_Instrument = request("Search_Instrument")

		
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

 	Response.Write ("Exec Retrieve_TransactionHistory_To_Excel '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"', '"&Search_instrument&"' ")
 	rs.open ("Exec Retrieve_DetailTrade_To_Excel '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"', '"&Search_instrument&"' ") , Conn , 3,1

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

<tr bgcolor="#ADF3B6" align="center">
   <td width="14%">Trade Code<br>交易編號</a></td>
   <td width="16%">Currency<br>貨幣</a></td>
   <td width="14%">Trade Date<br>交易日期</a></td>
   <td width="14%">Settle Date<br>結算日期</a></td>
   <td width="30%">Client Code <br> 客戶編號 </a></td>
   <td width="10%">Buy/Sell<br>買/賣</a></td>
   <td width="10%">Location<br>地點</a></td>
   <td width="10%">Stock Code<br>股票編號</a></td>
   <td width="10%">Share No<br>股票數量</a></td>
   <td width="10%">Price<br>價錢</a></td>
   <td width="10%">Consideration<br>交易總額</a></td>
   <td width="17%">Client Brokerage<br>客戶佣金 </a></td> 
   <td width="17%">Client Rebate<br>客戶回扣</a></td> 
   <td width="17%">Broker Brokerage<br>經紀佣金</a></td> 
   <td width="17%">Charge 1<br>收入 一</a></td> 
   <td width="17%">Charge 2<br>收入 二</a></td> 
   <td width="17%">Charge 3<br>收入 三</a></td> 
   <td width="17%">Charge 4<br>收入 四</a></td> 
   <td width="17%">Charge 5<br>收入 五</a></td> 
   <td width="17%">Charge 6<br>收入 六</a></td> 
   <td width="17%">Charge 7<br>收入 七</a></td> 
   <td width="17%">Broker Rebate<br>經紀回扣</a></td> 
   <td width="17%">Turnover<br>交易量</a></td>
   <td>Rate</td> 
   <td width="17%">Confirmation Date<br>確認日期</a></td> 
   <td width="17%">Brokerage Rate<br>佣金比率</a></td> 
</tr>

<%
' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
%>

<tr bgcolor="#ADF3B6" align="center">
			<td><% =rs("TRADENO")  %></td>
			<td><% =rs("TRADINGCCY")  %></td>
			<td><% =rs("TRADEDATE")  %></td>
			<td><% =rs("SETTLEDATE")  %></td>
			<td><% =rs("CLIENTCODE")  %></td>
			<td><% =rs("BUYSELL")  %></td>
			<td><% =rs("MARKET")  %></td>
			<td><% =rs("INSTRUMENT")  %></td>
			<td><% =rs("TTLQTY")  %></td>
			<td><% =rs("PRICE")  %></td>
			<td><% =rs("SETFXAMOUNT")  %></td>
			<td><% '=rs("a")  %></td> 
			<td><% '=rs("b")  %></td> 
			<td><% =rs("ORFEE1")  %></td> 
			<td><% =rs("ORFEE2")  %></td> 
			<td><% =rs("ORFEE3")  %></td> 
			<td><% =rs("ORFEE4")  %></td> 
			<td><% =rs("ORFEE5")  %></td> 
			<td><% =rs("ORFEE6")  %></td> 
			<td><% =rs("ORFEE7")  %></td> 
			<td><% =rs("ORFEE8")  %></td> 
			<td><% =rs("REBATEAMOUNT")  %></td> 
			<td><% =rs("CONSIDERATION")  %></td>
                        <td><% =rs("XRate")  %></td>   
			<td><% =rs("CONFIRMATIONDATE")  %></td> 
			<td><% =rs("BROKERAGERATE")  %></td> 

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