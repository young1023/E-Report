<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=DetailTrade.xls"


if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



'**************
'Initialisation
'**************


'Const adOpenStatic = 3
'Const adLockReadOnly = 1
'Const adCmdText = &H0001

Const RECORDPERPAGE = 10  ' The size of our pages.

Dim iPageCurrent ' The page we're currently on
Dim iPageCount   ' Number of pages of records
Dim iRecordCount ' Count of the records returned
Dim I            ' Standard looping variable
Dim iRecord ' Counter for page natvigator

Dim cnnSearch  ' ADO connection
Dim rstSearch  ' ADO recordset

Dim  itotalCCY
Dim iPageturnover(), iPageconsideration(), iPagebrokerage(), iPageCCY()
Dim itotalAECommFC, MTDAEcommFC
itotalturnover = 0
itotalconsideration = 0
itotalbrokerage = 0
itotalNetAmount = 0
itotalAECommFC  = 0
MTDTurnover = 0
MTDBrokerage = 0
MTDNetAmount = 0
YTDTurnover  = 0
YTDBrokerage = 0
YTDNetAmount = 0
MTDAEcommFC = 0

'Dim itotalturnover, itotalconsideration, itotalbrokerage, itotalCCY, itotalNetAmount
itotalturnover = 0
itotalconsideration = 0
itotalbrokerage = 0
itotalNetAmount = 0
 itotalNetComm = 0

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

 	'Response.Write ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") 

 	Rs.open ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1

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

<%

         set Rs4 = server.createobject("adodb.recordset")

        Rs4.open ("exec Retrieve_RebateAmount_HKD '"&rs("ClientCode")&"','"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' "),  Conn,3,1

  'Response.Write  ("exec Retrieve_RebateAmount_HKD '"&rs("ClientCode")&"','"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ")


  Dim NetComm

            NetComm = 0

            totalBrokerage = formatnumber(rs("totalBrokerage"),2)

            ClientRebateFC = formatnumber(rs4("ClientRebateFC"),6)

            AECommFC = formatnumber(Rs4("AECommFC"),6)

            BrokerCommFC = formatnumber(Rs4("BrokerCommFC"),2)

            BrokerRebateFC = formatnumber(Rs4("BrokerRebateFC"),2)

            IntroducerRebateFC = formatnumber(Rs4("IntroducerRebateFC"),6)

            ResearchCreditFC = formatnumber(Rs4("ResearchCreditFC"),6)

            'Response.write AECommFC & "<br>"
	
if AEcommFC > 0 then

	Netcomm = totalbrokerage - AECommFC
        
else
	Netcomm = Totalbrokerage - IntroducerRebateFC 
end if



if ClientRebateFC > 0 then
	Netcomm = Netcomm - ClientRebateFC 
else
	Netcomm = Netcomm - ResearchCreditFC
end if


%>


   <td width="15%"><%=formatnumber(NetComm,2)%></td> 
   <td width="15%"><%=formatnumber(rs("totalNetAmount"),2) %></td> 
</tr>

<%

   itotalturnover = itotalturnover + formatnumber(rs("totalturnover"),6)

   itotalbrokerage = itotalbrokerage + formatnumber(rs("totalBrokerage"),6)

   itotalNetComm  = itotalNetComm + formatnumber(NetComm,6)

   itotalNetAmount = itotalNetAmount + formatnumber(rs("totalNetamount"),6)


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
   <td ><%=formatnumber(itotalNetComm,2) %></td>
   <td ><%=formatnumber(itotalNetAmount,2) %></td>

</tr>

<%

    set Rs2 = server.createobject("adodb.recordset")
    Set Rs8 = server.createobject("adodb.recordset")  
  

	Select Case  Search_SharedSelection
	case "share2"
		'shared group member
         response.write "Shared"
 			
 			Rs2.open ("Exec retrieve_DetailTrade_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '',  '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1

                       Rs8.open ("Exec Retrieve_RebateAmount_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") ,  Conn,3,1
  
			
	case "share3"

 			Rs2.open ("Exec retrieve_DetailTrade_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1

                        Rs8.open ("Exec Retrieve_RebateAmount_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") ,  Conn,3,1
   
			
	case else
		'normal
 
 			Rs2.open ("Exec retrieve_DetailTrade_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '1', '100', 'ClientCode', 'ASC' ") ,  StrCnn,3,1
		 'Response.write  ("Exec retrieve_DetailTrade_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ")        

                    Rs8.open ("Exec Retrieve_RebateAmount_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") ,  Conn,3,1
                 'Response.write   	("Exec Retrieve_RebateAmount_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") 
 
	
	end select		

    MTDAECommFC = 0

    If Not Rs2.EoF Then

    MTDTurnover  = formatnumber(rs2("totalturnover"),2)
    MTDBrokerage = formatnumber(rs2("totalBrokerage"),2)
    
    
    MTDAECommFC = formatnumber(Rs8("AECommFC"),2)

    If IsNull(MTDAECommFC) Then

    MTDAECommFC = 0

    End If

    MTDIntroducerRebateFC =   formatnumber(Rs8("IntroducerRebateFC"),2)

    If IsNull(MTDIntroducerRebateFC) Then

    MTDIntroducerRebateFC = 0

    End If

    MTDClientRebateFC =  formatnumber(Rs8("ClientRebateFC"),2)

    If IsNull(MTDClientRebateFC) Then

    MTDClientRebateFC = 0

    End If

    MTDResearchCreditFC = formatnumber(Rs8("ResearchCreditFC"),2)

    If IsNull(MTDResearchCreditFC) Then

    MTDResearchCreditFC = 0

    End If


    MTDIntroducerRebateFC =  formatnumber(Rs8("IntroducerRebateFC"),2)

    If IsNull(MTDIntroducerRebateFC) Then

    MTDIntroducerRebateFC = 0

    End If


    MTDNetAmount = formatnumber(rs2("totalNetAmount"),2)

 
    MTDNetcomm = MTDBrokerage - MTDAECommFC - MTDResearchCreditFC - MTDResearchCreditFC - MTDIntroducerRebateFC
   


  
    MTDNetAmount = formatnumber(rs2("totalNetAmount"),2)

    End If
 
  
%>

<tr bgcolor="#FFFFCC">
   <td>&nbsp;</td> 
   <td >MTD Subtotal (HKD)</td>
   <td ><%=MTDTurnover%></td>
   <td ><%=rs2("totalBrokerage")%></td>
    <td ><% =MTDNetComm %></td>
   <td ><%=MTDNetAmount%></td>

</tr>

<%

  'End If

Rs2.Close
Set Rs2=Nothing

   set Rs3 = server.createobject("adodb.recordset")
   Set Rs9 = server.createobject("adodb.recordset")  


	Select Case  Search_SharedSelection
	case "share2"
		'shared group member
         response.write "Shared"
 			
 			Rs3.open ("Exec retrieve_DetailTrade_YTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '',  '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1


                        Rs9.open ("Exec Retrieve_RebateAmount_YTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") ,  Conn,3,1
  
			
	case "share3"

 			Rs3.open ("Exec retrieve_DetailTrade_YTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1

               Rs9.open ("Exec Retrieve_RebateAmount_YTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") ,  Conn,3,1

			
	case else
		'normal
 
 			Rs3.open ("Exec retrieve_DetailTrade_YTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1

                        Rs9.open ("Exec Retrieve_RebateAmount_YTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") ,  Conn,3,1
   	
	'response.write	    ("Exec Retrieve_RebateAmount_YTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") 
			
	end select


If Not Rs3.EoF Then	

Dim YTDNetComm
    YTDNetComm = 0

    
    YTDTurnover  = formatnumber(rs3("totalturnover"),2)
    YTDBrokerage = formatnumber(rs3("totalBrokerage"),2)


     
    YTDAECommFC = formatnumber(Rs9("AECommFC"),2)

    If IsNull(YTDAECommFC) Then

    YTDAECommFC = 0

    End If

    YTDIntroducerRebateFC =   formatnumber(Rs9("IntroducerRebateFC"),2)

    If IsNull(YTDIntroducerRebateFC) Then

    YTDIntroducerRebateFC = 0

    End If

    YTDClientRebateFC =  formatnumber(Rs9("ClientRebateFC"),2)

    If IsNull(YTDClientRebateFC) Then

    YTDClientRebateFC = 0

    End If

    YTDResearchCreditFC = formatnumber(Rs9("ResearchCreditFC"),2)

    If IsNull(YTDResearchCreditFC) Then

    YTDResearchCreditFC = 0

    End If


    YTDIntroducerRebateFC =  formatnumber(Rs9("IntroducerRebateFC"),2)

    If IsNull(YTDIntroducerRebateFC) Then

    YTDIntroducerRebateFC = 0

    End If



       YTDNetcomm = YTDBrokerage - YTDAECommFC - YTDResearchCreditFC - YTDResearchCreditFC - YTDIntroducerRebateFC
  
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
   <td ><%=YTDNetComm %></td>
   <td ><%=YTDNetAmount%></td>

</tr>


</table>

</div>

</body>
</html>

