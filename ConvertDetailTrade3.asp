<%
' Tells the browser to open excel
'Response.ContentType = "application/vnd.ms-excel" 
'Response.addHeader "content-disposition","attachment;filename=DetailTrade.xls"



On Error resume Next




'**************
'Initialisation
'**************


Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adCmdText = &H0001

Const RECORDPERPAGE = 10  ' The size of our pages.

Dim iPageCurrent ' The page we're currently on
Dim iPageCount   ' Number of pages of records
Dim iRecordCount ' Count of the records returned
Dim I            ' Standard looping variable
Dim iRecord ' Counter for page natvigator

Dim cnnSearch  ' ADO connection
Dim rstSearch  ' ADO recordset

Dim itotalturnover, itotalconsideration, itotalbrokerage, itotalCCY, itotalNetAmount
Dim iPageturnover(), iPageconsideration(), iPagebrokerage(), iPageCCY()
Dim itotalAECommFC
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

strURL = Request.ServerVariables("URL") ' Retreive the URL of this page from Server Variables
%>






<html>
<head>
	
	    <style type="text/css">
    <!-- Hide from legacy browsers
    .print { 
    display: none;
    }
    @media print {
    	.noprint {
    	 display: none;
    	}'
    }  -->
    
    </style>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<script src="include/sorttable.js"></script>
<SCRIPT language=JavaScript>
<!--

function datevalidate(inDay, inMonth, inYear){
	
		var myDayStr = inDay;
		var myMonthStr = inMonth;
		var myYearStr = inYear;	
		var myMonth = new Array('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'); 
		var myDateStr = myDayStr + ' ' + myMonth[myMonthStr] + ' ' + myYearStr;


		var myDate = new Date();
		myDate.setFullYear( myYearStr, myMonthStr, myDayStr );

		if ( myDate.getMonth() != myMonthStr ) {
		  alert( myDateStr + ' is NOT a valid date.' );
		  return false;
		}

		return true;
}


function PopupClientContact(clientnumber) {
	 
		var str='ListClientContact.asp?sid=<%=SessionID%>&clientnumber=' + clientnumber
		
		newwindow=window.open(str , "myWindow", 
									"status = 1, height = 300, width = 600, resizable = 1'"  )
		 if (window.focus) {
           newwindow.focus();
       }
 			
}

function validateUserEntry(){

		//date validation
		if ((datevalidate(document.fm1.FromDay.value, document.fm1.FromMonth.value -1, document.fm1.FromYear.value) == false) || 
				(datevalidate(document.fm1.ToDay.value, document.fm1.ToMonth.value -1, document.fm1.ToYear.value) == false))
		{
			return false;

		}
		
		//User must enter Client From field
		if (document.fm1.ClientFrom.value == ""){
  			alert("Please enter client number");
        document.fm1.ClientFrom.focus();
        return false;
		}

<% if session("shell_power") > 1 then %>
		if  (isNaN(document.fm1.AEFrom.value) == true){
  			alert("AE Code should in numeric format");
        document.fm1.AEFrom.focus();
        return false;
		}
		
		if  (isNaN(document.fm1.AETo.value) == true ){
  			alert("AE Code should in numeric format");
        document.fm1.AETo.focus();
        return false;
		}		
<% end if %>		
				
		return true;
}

function dosubmit(what){
  
			if (validateUserEntry() == false)
			{
				return false

			}
				document.fm1.submitted.value=1;
			  document.fm1.action="<%= strURL %>?sid=<%=SessionID%>";
				document.fm1.page.value=what;
			  document.fm1.submit();
	
}


function ordersubmit(iorder, idirection){
	
	if (validateUserEntry == false)
	{
		return false

	}
		document.fm1.submitted.value=1;
  document.fm1.action="<%= strURL %>?sid=<%=SessionID%>";
	document.fm1.Order.value=iorder;
	document.fm1.Direction.value=idirection;
  document.fm1.submit();
}


//-->
</SCRIPT>

</head>

<body leftmargin="0" topmargin="0" >








<%


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



Search_AEGroup	        = Request.form("GroupID")
'Search_ClientFrom       = Request.form("ClientFrom")
'Search_ClientTo         = Request.form("ClientTo")
'Search_AEFrom           = Request.form("AEFrom")
'Search_AETo             = Request.form("AETo")
'Search_From_Day         = Request.form("FromDay")
'Search_From_Month       = Request.form("FromMonth")
'Search_From_Year        = Request.form("FromYear")
'Search_To_Day           = Request.form("ToDay")
'Search_To_Month         = Request.form("ToMonth")
'Search_To_Year          = Request.form("ToYear")
Search_Transaction_Type = Request.form("TranType")
'Search_Market           = Request.form("Market")
'Search_Instrument       = Request.form("Instrument")
'Search_Order            = Request.form("Order")
'Search_Direction        = Request.form("Direction")
Search_SharedSelection  = Request.form("ShareSelection")	
Search_SharedGroup      = Request.form("SharedGroup")
Search_SharedGroupMember= Request.form("SharedGroupMember")


Search_SharedGroup = "None"

	
	
 set Rs1 = server.createobject("adodb.recordset")


 'Rs return 2 value
 '1) Total number of matched client
 '2) all records for targeted client


	Select Case  Search_SharedSelection
	case "share2"
		'shared group member
         response.write "Shared"
 			
 			Rs1.open ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '',  '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
			
	case "share3"

 			Rs1.open ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
			
	case else
		'normal
 
 			Rs1.open ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '1', '10', 'TRADEDATE', 'ASC' ") ,  StrCnn,3,1
	
	end select	


	          Response.write ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '1', '10', 'TRADEDATE', 'ASC' ") 




  
  dim itotalCCYcount
	erase itotalturnover
	erase itotalconsideration		
	erase itotalBrokerage		
  itotalCCYcount=0

  
if Rs1.EoF then




Else
           
 
%>    

</span>




   
<%
'**********
' Start of page navigation 
'**********
%> 
    

  




<table width="97%" border="0" class="normal"  cellspacing="1" cellpadding="4">
<tr bgcolor="#FFFFCC"> 
<td  width="20%">　</td>
      <td align="center">詳細交易紀錄<br><u>Detail Trade Information</u></td> 
      <td align="right" width="20%"><span class="noprint">
							
			</span></td>
</tr>
</table>
<br>
<table class="sortable"  width="99%" border="0" style="border-width: 0;FONT-SIZE: 11px;TEXT-ALIGN: Right;FONT-FAMILY: Verdana, 'MS Sans Serif', Arial" bgcolor="#808080" cellspacing="1" cellpadding="2">
<tr class="alignright" bgcolor="#ADF3B6" align="center">
   <td width="15%"><span style="cursor:hand">Client Code <br> 1客戶編號</span></td>
   <td width="25%"><span style="cursor:hand">Client Name<br>客戶</span></td>
   <td width="15%"><span style="cursor:hand">Turnover (HKD)<br>交易總額</span></td>
   <td width="15%"><span style="cursor:hand">Brokerage (HKD)<br>佣金</span></td> 
   <td width="15%"><span style="cursor:hand">Net Comm (HKD)<br>剩收入</sapn></td> 
   <td width="15%"><span style="cursor:hand">Net Amount (HKD)<br>總額</span></td> 
</tr>

		<%


  
                        Dim itotalNetComm
  
                       itotalNetComm = 0


            
			dim iPageCCYcount
			dim k
			dim iPageUpdate

			

			dim mystr
			do while (Not rs1.EOF)
				k=1



      
            
            set Rs4 = server.createobject("adodb.recordset")

            'Rs4.open ("exec Retrieve_RebateAmount_HKD '"&rs1("ClientCode")&"','"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' "),  Conn,3,1

     Response.Write  ("exec Retrieve_RebateAmount_HKD '"&rs1("ClientCode")&"','"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ")

            

            Dim NetComm

            NetComm = 0

            totalBrokerage = formatnumber(rs1("totalBrokerage"),2)

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
		
<tr class="alignright" bgcolor="#FFFFCC"> 
   <td width="15%"><a href="DetailTrade2.asp?PrintAllowed=<%=PrintAllowed%>&DisplayFirst=<%=Trim(rs1("ClientCode"))%>&ClientFrom=<%=Search_ClientFrom%>&ClientTo=<%=Search_ClientTo%>&AEFrom=<%=Search_AEFrom%>&AETo=<%=Search_AETo%>&Instrument=<%=Search_Instrument%>&Market=<%=Search_Market%>&FromDay=<%=Search_From_Day%>&FromMonth=<%=Search_From_Month%>&FromYear=<%=Search_From_Year%>&Today=<%=Search_To_Day%>&ToMonth=<%=Search_To_Month%>&ToYear=<%=Search_To_Year%>&Search_Order=ClientCode&Search_Direction=ASC&sid=<%=SessionID%>#DisplayFirst" target=_blank><%=rs1("ClientCode") %></a><span class="noprint"><img border=0 src='images/tel.gif' onClick="PopupClientContact('<%=rs1("ClientCode") %>')"></img></span></td>
   <td width="25%"><%=rs1("ClientName") %></td>
   <td width="15%"><%=formatnumber(rs1("totalturnover"),2)  %></td> 
   <td width="15%"><%=formatnumber(totalBrokerage,2) %></td> 
   <td width="15%"><%=formatnumber(NetComm,2)%></td> 
   <td width="15%"><%=formatnumber(rs1("totalNetAmount"),2) %></td> 
</tr>

<%

   
   itotalturnover = itotalturnover + formatnumber(rs1("totalturnover"))

   itotalbrokerage = itotalbrokerage + formatnumber(rs1("totalBrokerage"))

   itotalconsideration = itotalconsideration + formatnumber(rs1("totalconsideration"))

   itotalNetAmount = itotalNetAmount + formatnumber(rs1("totalNetamount"))

   itotalNetComm  = itotalNetComm + formatnumber(NetComm,6)

   'response.write itotalNetComm &"<br>"

    rs1.movenext

 
				
		loop


Rs1.Close
Set Rs1=Nothing



%>
</table>
<br>
<table class="alignright"  width="99%" border="0" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">


<tr bgcolor="#FFFFCC"> 

   <td colspan="2" width="40%" align="right">Subtotal (HKD)<BR></td>
   <td width="15%"><%=formatnumber(itotalturnover,2)  %></td>
   <td width="15%"><%=formatnumber(itotalbrokerage,2) %></td>
   <td width="15%"><%=formatnumber(itotalNetComm,2) %></td>
   <td width="15%"><%=formatnumber(itotalNetAmount,2) %></td>

</tr>

<tr bgcolor="#FFFFFF"> 
   <td colspan="6">&nbsp;</td>
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


   
  
 


    If Not Rs2.EoF Then

    MTDTurnover  = formatnumber(rs2("totalturnover"),2)
    MTDBrokerage = formatnumber(rs2("totalBrokerage"),2)
    
    
    MTDAECommFC = formatnumber(Rs8("AECommFC"),2)
    MTDIntroducerRebateFC =   formatnumber(Rs8("IntroducerRebateFC"),2)
    MTDClientRebateFC =  formatnumber(Rs8("ClientRebateFC"),2) 
    MTDResearchCreditFC = formatnumber(Rs8("ResearchCreditFC"),2)
    MTDIntroducerRebateFC =  formatnumber(Rs8("IntroducerRebateFC"),2)


    MTDNetAmount = rs2("totalNetAmount")

   ' MTDNetComm = MTDBrokerage - MTDAECommFC - MTDIntroducerRebateFC - MTDResearchCreditFC - MTDIntroducerRebateFC




if MTDAEcommFC > 0 then

	MTDNetcomm = MTDBrokerage - MTDAECommFC
else
	MTDNetcomm = MTDBrokerage - MTDIntroducerRebateFC 
end if

if MTDResearchCreditFC > 0 then

	MTDNetcomm = MTDNetcomm - MTDResearchCreditFC
else
	MTDNetcomm = MTDNetcomm - MTDIntroducerRebateFC
end if
  
    MTDNetAmount = formatnumber(rs2("totalNetAmount"),2)

    End If

%>
   
<tr bgcolor="#FFFFCC"> 
	 
   <td colspan="2" align="right">MTD Subtotal (HKD)</td>
   <td ><%=MTDTurnover%></td>
   <td ><%=formatnumber(rs2("totalBrokerage"),2)%></td>
   <td ><%=formatnumber(MTDNetComm,2)%></td>
   <td ><%=MTDNetAmount%></td>

</tr>

<%    

Rs2.Close
Set Rs2=Nothing
 
%>




<tr bgcolor="#FFFFFF"> 
   <td colspan="6">&nbsp;</td>
</tr>

<%

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
    YTDIntroducerRebateFC =   formatnumber(Rs9("IntroducerRebateFC"),2)
    YTDClientRebateFC =  formatnumber(Rs9("ClientRebateFC"),2) 
    YTDResearchCreditFC = formatnumber(Rs9("ResearchCreditFC"),2)
    YTDIntroducerRebateFC =  formatnumber(Rs9("IntroducerRebateFC"),2)

   ' YTDNetComm = YTDBrokerage - YTDAECommFC - YTDIntroducerRebateFC - YTDResearchCreditFC - YTDIntroducerRebateFC

if YTDAEcommFC > 0 then

	YTDNetcomm = YTDBrokerage - YTDAECommFC
else
	YTDNetcomm = YTDBrokerage - YTDIntroducerRebateFC 
end if

if YTDResearchCreditFC > 0 then

	YTDNetcomm = YTDNetcomm - YTDResearchCreditFC
else
	YTDNetcomm = YTDNetcomm - YTDIntroducerRebateFC
end if
  
    YTDNetAmount = formatnumber(rs3("totalNetAmount"),2)
  
    

   
    End If

Rs3.Close
Set Rs3=Nothing

%>

<tr bgcolor="#FFFFCC"> 

   <td colspan="2" align="right">YTD Subtotal (HKD)<BR></td>
   <td ><%=YTDTurnover%></td>
   <td ><%=YTDBrokerage%></td>
   <td ><%=formatnumber(YTDNetComm,2)%></td>
   <td ><%=YTDNetAmount%></td>
</tr>

<% End If %>

</table>

                
              
              </body>
              </html>
              
<%
'*****************************************************************
' Termination
'*****************************************************************

 'Rs1.Close
 set Rs1 = Nothing
 'Rs2.close
 Set Rs2 = Nothing
 Conn.Close
 Set Conn = Nothing
 
%>
