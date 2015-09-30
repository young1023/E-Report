<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=DetailTrade.xls"


if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

Dim itotalturnover, itotalconsideration, itotalbrokerage, itotalCCY, itotalNetAmount
itotalturnover = 0
itotalconsideration = 0
itotalbrokerage = 0
itotalNetAmount = 0

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
        Search_AEGroup      = Request("Search_SharedGroup")
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

 	'Response.Write ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '8', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") 

 	Rs.open ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '8', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1

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
   <td width="15%">Client Code <br> 客戶編號 </td>
   <td width="25%">Client Name<br>客戶</td>
   <td width="15%">Turnover (HKD)<br>交易總額</td>
   <td width="15%">Brokerage (HKD)<br>客戶佣金</td> 
   <td width="15%">Net Comm (HKD)<br>經紀佣金</td> 
   <td width="15%">Net Amount (HKD)<br>總額</td> 
</tr>

<%
' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
%>

<tr bgcolor="#FFFFCC"> 
   <td width="15%"><%=rs("ClientCode") %></td>
   <td width="25%"><%=rs("ClientName") %></td>
   <td width="15%"><%=formatnumber(rs("totalturnover"),2)  %></td> 
   <td width="15%"><%=formatnumber(rs("totalBrokerage"),2)%></td> 
   <td width="15%"><%=formatnumber(rs("totalBrokerage"),2) %></td> 
   <td width="15%"><%=formatnumber(rs("totalNetAmount"),2) %></td> 
</tr>

<%

   itotalturnover = itotalturnover + formatnumber(rs("totalturnover"))

   itotalbrokerage = itotalbrokerage + formatnumber(rs("totalBrokerage"))

   itotalconsideration = itotalconsideration + formatnumber(rs("totalconsideration"))

   itotalNetAmount = itotalNetAmount + formatnumber(rs("totalNetamount"))


' Move to the next record
rs.movenext
' Loop back to the do statement
loop 

Rs.Close
Set Rs=Nothing

%>


<tr bgcolor="#FFFFCC"> 
   <td>&nbsp;</td>
   <td align="right">Subtotal (HKD)<BR></td>
   <td ><%=formatnumber(itotalturnover,2)  %></td>
   <td ><%=formatnumber(itotalbrokerage,2) %></td>
   <td ><%=formatnumber(itotalbrokerage,2) %></td>
   <td ><%=formatnumber(itotalNetAmount,2) %></td>

</tr>

<%

    set Rs2 = server.createobject("adodb.recordset")

	Select Case  Search_SharedSelection
	case "share2"
		'shared group member
         response.write "Shared"
 			
 			Rs2.open ("Exec retrieve_DetailTrade_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '',  '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
			
	case "share3"

 			Rs2.open ("Exec retrieve_DetailTrade_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
			
	case else
		'normal
 
 			Rs2.open ("Exec retrieve_DetailTrade_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '8', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '1', '100', 'ClientCode', 'ASC' ") ,  StrCnn,3,1
		    'Response.write  ("retrieve_DetailTrade_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '8', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ")        
	
	end select	

  


    If Not Rs2.EoF Then

    MTDTurnover  = formatnumber(rs2("totalturnover"),2)
    MTDBrokerage = formatnumber(rs2("totalBrokerage"),2)
    
    MTDNetAmount = rs2("totalNetAmount")

    MTDNetAmount = formatnumber(rs2("totalNetAmount"),2)

 
  
%>

<tr bgcolor="#FFFFCC">
   <td >&nbsp;</td> 
   <td align="right">MTD Subtotal (HKD)</td>
   <td ><%=MTDTurnover%></td>
   <td ><%=MTDBrokerage%></td>
   <td ><%=MTDBrokerage%></td>
   <td ><%=MTDNetAmount%></td>

</tr>

<%

  End If

Rs2.Close
Set Rs2=Nothing

   set Rs3 = server.createobject("adodb.recordset")

	Select Case  Search_SharedSelection
	case "share2"
		'shared group member
         response.write "Shared"
 			
 			Rs3.open ("Exec retrieve_DetailTrade_YTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '',  '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
			
	case "share3"

 			Rs3.open ("Exec retrieve_DetailTrade_YTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
			
	case else
		'normal
 
 			Rs3.open ("Exec retrieve_DetailTrade_YTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
			
	end select	

If Not Rs3.EoF Then

    YTDTurnover  = formatnumber(rs3("totalturnover"),2)
    YTDBrokerage = formatnumber(rs3("totalBrokerage"),2)
    YTDNetAmount = formatnumber(rs3("totalNetAmount"),2)
    
    End If

Rs3.Close
Set Rs3=Nothing

%>

<tr bgcolor="#FFFFCC"> 
   <td>&nbsp;</td>
   <td align="right">YTD Subtotal (HKD)<BR></td>
   <td ><%=YTDTurnover%></td>
   <td ><%=YTDBrokerage%></td>
   <td ><%=YTDBrokerage%></td>
   <td ><%=YTDNetAmount%></td>

</tr>


</table>

</div>

</body>
</html>

