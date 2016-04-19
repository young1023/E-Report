<% 
'*********************************************************************************
'NAME       : DetailTrade2.asp           
'DESCRIPTION: Instrument trading with all details info
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 090401 Gary Yeung   Prototype
'MODIFIED   : 090415 Roger Wong   Record and page control
'			  090712 Roger Wong		Add Shared Group
'********************************************************************************

'On Error resume Next

Server.ScriptTimeout = 7200000
%>

<!--#include file="include/SessionHandler.inc.asp" -->
<%

if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

Title = "Detail Trade"


'**************
'Initialisation
'**************


Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adCmdText = &H0001

Const RECORDPERPAGE = 100000000  ' The size of our pages.

Dim iPageCurrent ' The page we're currently on
Dim iPageCount   ' Number of pages of records
Dim iRecordCount ' Count of the records returned
Dim I            ' Standard looping variable
Dim iRecord ' Counter for page natvigator

Dim cnnSearch  ' ADO connection
Dim rstSearch  ' ADO recordset

Dim itotalturnover, itotalconsideration, itotalbrokerage, itotalCCY, itotalNetAmount
Dim iPageturnover(), iPageconsideration(), iPagebrokerage(), iPageCCY()
Dim TotalAECommFC, TotalIntroducerRebateFC, TotalClientRebateFC, TotalResearchCreditFC
Dim SumAECommFC, SumIntroducerRebateFC, SumClientRebateFC, SumTotalResearchCreditFC
Dim NetComm
DIM Search_SharedSelection  


TotalTurnover = 0
NetComm = 0
TotalResearchCreditFC  = 0
TotalClientRebateFC = 0
TotalAECommFC = 0
TotalIntroducerRebateFC = 0
itotalturnover = 0
itotalconsideration = 0
itotalbrokerage = 0
itotalNetAmount = 0

SumAECommFC = 0
SumIntroducerRebateFC=0
SumClientRebateFC =0
SumTotalResearchCreditFC=0
SumTotalNewComm=0
SumTotalResearchCreditFC=0

strURL = Request.ServerVariables("URL") ' Retreive the URL of this page from Server Variables

if session("shell_power")="" then
  response.redirect "Default.asp"
end if


%>


<html>
<head>
	

<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<script src="include/common.js"></script>
<script src="include/common_original.js"></script>
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


function dosubmit(what){
  
			document.fm1.submitted.value=what;
			document.fm1.action="<%= strURL %>?sid=<%=SessionID%>";
		    document.fm1.submit();
	
}



// --> 
</script>

</head>

<%

'**************
'Argument handler
'**************

Dim Search_ClientFrom
Dim Search_ClientTo
Dim Search_AEFrom
Dim Search_AETo
Dim Search_AEGroup
Dim Search_From_Day
Dim Search_From_Month
Dim Search_From_Year
Dim Search_To_Day
Dim Search_To_Month
Dim Search_To_Year
Dim Search_Market
Dim Search_Instrument
Dim Search_Order 
Dim Search_Direction
Dim Search_SharedGroup
Dim Search_SharedGroupMember
Dim ShowSubTotal


SharedNumber            = Trim(Request("SharedNumber"))
SharedName              = Trim(Request("SharedName"))
Search_ClientFrom1      = session("ClientFrom")                           
Search_ClientTo1        = session("ClientTo") 

Search_AEGroup	        = Request("GroupID")
Search_ClientFrom       = Trim(Request("DisplayFirst"))
Search_ClientTo         = Trim(Request("DisplayFirst"))
Search_AEFrom           = Request("AEFrom")
Search_AETo             = Request("AETo")
Search_From_Day         = Request("FromDay")
Search_From_Month       = Request("FromMonth")
Search_From_Year        = Request("FromYear")
Search_To_Day           = Request("ToDay")
Search_To_Month         = Request("ToMonth")
Search_To_Year          = Request("ToYear")
Search_Transaction_Type = Request("TranType")
Search_Market           = Request("Market")
Search_Instrument       = Request("Instrument")
Search_Order            = Request("Order")
Search_Direction        = Request("Direction")
Search_SharedGroup      = Request("SharedGroup")
Search_SharedGroupMember= Request("SharedGroupMember")
Search_Order            = Request("Search_Order")
Search_Direction        = Request("Search_Direction")
Search_SharedSelection  = Request("ShareSelection")

'AECode search permission
Select Case Session("shell_power")
	case "1"
		'AE shall access their own clients only
		Search_AEFrom = Session("id")
		Search_AETo = Session("id")
		Search_AEGroup = Session("GroupID")
     	
    case "5"
		' Branch Manager shall access all AE's clients belongs to 
		Search_AEGroup = Session("GroupID")
	
		'Other having full access
end select


' If User enter From value only, change the "To" value to "From"
if Search_ClientTo = "" then
   Search_ClientTo = Search_ClientFrom
end if

if Search_AETo = "" then
   Search_AETo = Search_AEFrom
end if

' Retrieve page to show or default to the first

If request.form("submitted")  = 1 Then


    Search_ClientFrom    = SharedNumber
    Search_ClientTo      = SharedNumber
    SharedName           = SharedNumber
	Search_Order         = "ClientCode"
	Search_Direction     = "ASC"
	Search_From_Day      = session("Search_From_Day")  
	Search_From_Month    = session("Search_From_Month")
	Search_From_Year     = session("Search_From_Year")
	Search_To_Day        = session("Search_To_Day")
	Search_To_Month      = session("Search_To_Month") 
	Search_To_Year       = session("Search_To_Year")
	Search_AEGroup	     = session("GroupID")
	Search_AEFrom        = session("AEFrom")
	Search_AETo          = session("AETo")
	Search_Market        = session("Search_Market")
    Search_Instrument    = session("Search_Instrument") 
   
End If

If request.form("submitted")  = 2 Then


    Search_ClientFrom    = SharedName
    Search_ClientTo      = SharedName
    SharedNumber         = SharedName
	Search_Order         = "ClientCode"
	Search_Direction     = "ASC"
	Search_From_Day      = session("Search_From_Day")  
	Search_From_Month    = session("Search_From_Month")
	Search_From_Year     = session("Search_From_Year")
	Search_To_Day        = session("Search_To_Day")
	Search_To_Month      = session("Search_To_Month") 
	Search_To_Year       = session("Search_To_Year")
	Search_AEGroup	     = session("GroupID")
	Search_AEFrom        = session("AEFrom")
	Search_AETo          = session("AETo")
	Search_Market        = session("Search_Market")
    Search_Instrument    = session("Search_Instrument") 
   
End If

If request("submitted")  = 3 Then


    Search_ClientFrom    = request("ClientFrom")
    Search_ClientTo      = request("ClientTo")
	Search_Order         = "ClientCode"
	Search_Direction     = "ASC"
	Search_From_Day      = session("Search_From_Day")  
	Search_From_Month    = session("Search_From_Month")
	Search_From_Year     = session("Search_From_Year")
	Search_To_Day        = session("Search_To_Day")
	Search_To_Month      = session("Search_To_Month") 
	Search_To_Year       = session("Search_To_Year")
	Search_AEGroup	     = session("GroupID")
	Search_AEFrom        = session("AEFrom")
	Search_AETo          = session("AETo")
	Search_Market        = session("Search_Market")
    Search_Instrument    = session("Search_Instrument") 
   
End If


    '  Query for Client Number and Client Name

    set Rs = server.createobject("adodb.recordset")
     
    'response.write  ("Exec Retrieve_DetailTrade_ClientCode '"&Search_ClientFrom1&"', '"&Search_ClientTo1&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") 
 
    Rs.open ("Exec Retrieve_DetailTrade_ClientCode '"&Search_ClientFrom1&"', '"&Search_ClientTo1&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") ,  StrCnn,3,1

    set RsN = server.createobject("adodb.recordset")
     
    'response.write  ("Exec Retrieve_DetailTrade_ClientName '"&Search_ClientFrom1&"', '"&Search_ClientTo1&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") 
 
    RsN.open ("Exec Retrieve_DetailTrade_ClientName '"&Search_ClientFrom1&"', '"&Search_ClientTo1&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") ,  StrCnn,3,1


%>


<%
'*****************************************************************
' Start of report body
'*****************************************************************


            PrintAllowed =  0  

          ' Permission for Printing 
          '************************
         set pRs = server.createobject("adodb.recordset")

		 pRs.open ("Exec Check_PrintPermission '"&Session("MemberID")&"', '"&Title&"' ") , StrCnn,3,1

  
           iRecordCount = pRs(0)

           If iRecordCount > 0 then
 
           PrintAllowed =  1     

           End if    
%>    
 
<body leftmargin="0" rightmargin="0" topmargin="8"  onkeypress="return disableCtrlKeyCombination(event);" onkeydown="return disableCtrlKeyCombination(event);" >
<form name="fm1" method="post" action="">
<table width="99%" border="0" class="normal"  cellspacing="2" cellpadding="4">
		<tr> 
			<td colspan="4">
			&nbsp;
			<select name="SharedNumber" class="common" onChange="dosubmit(1)">
              <option value=""> -- Please select the client code -- </option>
    		   
                       <%      Rs.MoveFirst 

                               Do While Not Rs.EoF %>

               <option value="<%=Rs("ClientCode")%>" <%If Trim(SharedNumber)=Trim(Rs("ClientCode")) Then%>Selected<%End If%>>  <%=Rs("ClientCode")%> |  <%=Rs("ClientName")%>  </option>
                       <%
								Rs.Movenext
								Loop
						%>
              
			</select>	
			&nbsp;
			<select name="SharedName" class="common" onChange="dosubmit(2)">
              <option value=""> -- Please select the client Name -- </option>
             			   
                       <%      RsN.MoveFirst 

                               Do While Not RsN.EoF %>

               <option value="<%=RsN("ClientCode")%>" <%If Trim(SharedName)=Trim(RsN("ClientCode")) Then%>Selected<%End If%>>    <%=RsN("ClientName")%> | <%=RsN("ClientCode")%>  </option>
                       <%
								RsN.Movenext
								Loop
						%>
              
			</select>	
					</td>
             <td align="left"><a href="DetailTrade2.asp?submitted=3&ClientFrom=<%=Search_ClientFrom1%>&ClientTo=<%=Search_ClientTo1%>&AEFrom=<%=Search_AEFrom%>&AETo=<%=Search_AETo%>&Instrument=<%=Search_Instrument%>&Market=<%=Search_Market%>&FromDay=<%=Search_From_Day%>&FromMonth=<%=Search_From_Month%>&FromYear=<%=Search_From_Year%>&Today=<%=Search_To_Day%>&ToMonth=<%=Search_To_Month%>&ToYear=<%=Search_To_Year%>&Search_Order=ClientCode&Search_Direction=ASC&PrintAllowed=1&sid=<%=SessionID%>"  onclick="alert('The searching process may take a long time. Please be patient!');">Show All</a>
             </td>
		</tr> 
</table>

<br>
<table width="99%" border="0" class="normal"  cellspacing="2" cellpadding="4">
<tr bgcolor="#FFFFCC"> 
<td  width="20%">　</td>
      <td align="center">詳細交易紀錄<br><u>Detail Trade Information</u></td> 
      <td align="right" width="20%"><span class="noprint">
							<%if Request("PrintAllowed") = 1 then %>  
							<a href="javascript:window.print()">Friendly Print</a><% 'end if %> &nbsp;&nbsp;
<% End If %>
     	
			</span></td>
</tr>
 	<input type=hidden   name="submitted" value="">
 

</table>
</form>
<table width="99%" border="0" class="alignright" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

<tr bgcolor="#ADF3B6" align="center">
   <td width="14%">Trade Date<br>交易日期</td>
   <td width="14%">Trade No<br>交易編號</td>
   <td width="14%">Status<br></td>
   <td width="30%">Client Code <br> 客戶編號</td>
   <td width="30%">Client Name <br> 客戶名稱</td>
   <td width="10%">Location<br>地點</td>
   <td width="16%">Currency<br>貨幣</td>
   <td width="10%">Instrument Code<br>股票編號</td>
   <td width="10%">Instrument Name<br>股票編號</td>
   <td width="10%">Buy/Sell<br>買/賣</td>
   <td width="10%">QTY<br>股票數量</td>
   <td width="10%">Price<br>價錢</td>
   <td width="20%">Turnover<br>交易量</td>
   <td width="20%">Brokerage Rate<br>佣金比率</td> 
   <td width="20%">Broker Brokerage<br>經紀佣金</td>
<span class="noprint">
   <td width="14%">AE/Intro Reb</td>
   <td width="14%">Clnt Rebate/<br>Research Cr</td>
</span>
   <td width="150">Net Comm<br>佣金總額</td> 
   <td width="300">Net Amount<br>交易總額</td> 
   <td width="20%">Charge 1<br>收入 一</td> 
   <td width="20%">Charge 2<br>收入 二</td> 
   <td width="20%">Charge 3<br>收入 三</td> 
   <td width="20%">Charge 4<br>收入 四</td> 
   <td width="20%">Charge 5<br>收入 五</td> 
   <td width="20%">Charge 6<br>收入 六</td> 
   <td width="20%">Charge 7<br>收入 七</td>
<span class="noprint"> 
   <td width="20%">Client Brokerage<br>客戶佣金</td> 
   <td width="20%">Client Rebate<br>客戶回扣</td> 
   <td width="20%">Broker Rebate<br>經紀回扣</td> 
</span>
   <td width="20%">Confirmation Date<br>確認日期</td>
   <td width="20%">Exchange Rate</td>  
</tr>


<%

    ' New Object

Session("Search_Market") = Search_Market

    
'    set Rs1 = server.createobject("adodb.recordset")
     
	'response.write  ("Exec retrieve_DetailTrade_OrderBy_Market_ClientCode '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") 
 
' 	Rs1.open ("Exec retrieve_DetailTrade_OrderBy_Market_ClientCode '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") , StrCnn,2,3

  

'response.write "<BR> " &  Now & "<BR>"


'Start of program

dim lObj

set lObj = server.createobject ("StringHandle.clsDetailTrade")


lObj.SessionID         = Session("id")
lObj.SearchClientFrom  = Search_ClientFrom
lObj.SearchClientTo    = Search_ClientTo
lObj.SearchOrder       = Search_Order
lObj.SearchDirection   = Search_Direction
lObj.SearchFromDay     = Search_From_Day
lObj.SearchFromMonth   = Search_From_Month
lObj.SearchFromYear    = Search_From_Year
lObj.SearchToDay       = Search_To_Day
lObj.SearchToMonth     = Search_To_Month
lObj.SearchToYear      = Search_To_Year
lObj.SearchAEGroup     = Search_AEGroup
lObj.SearchAEFrom      = Search_AEFrom        
lObj.SearchAETo        = Search_AETo          
lObj.SearchMarket      = Search_Market        
lObj.SearchInstrument  = Search_Instrument    

lObj.Search_SharedSelection = Search_SharedSelection  
lObj.SearchSharedGroupMember   = Search_SharedGroupMember
lObj.SearchLevel  = session("shell_power")


Response.write lObj.lExecute

  'Response.write  ("Exec Retrieve_DetailTrade_MTD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Trim(ThisTradingCCY)&"',  '"&Trim(Search_Market)&"','"&Search_Instrument&"' ") 

   

%>
                
                
                
                </td>
                </tr>
              </table>
              



              </body>

              </html>
              
<%



'*****************************************************************
' Termination
'*****************************************************************

 Conn.Close
 Set Conn = Nothing

'pass all selection to session 
session("GroupID")                 =  Search_AEGroup	               
session("ClientFrom")              =  Search_ClientFrom1              
session("ClientTo")                =  Search_ClientTo1                
session("AEFrom")                  =  Search_AEFrom                  
session("AETo")                    =  Search_AETo                    
session("Search_From_Day")         =  Search_From_Day
session("Search_From_Month")       =  Search_From_Month
session("Search_From_Year")        =  Search_From_Year
session("Search_To_Day")           =  Search_To_Day
session("Search_To_Month")         =  Search_To_Month
session("Search_To_Year")          =  Search_To_Year
session("Search_Instrument")       =  Search_Instrument


Server.ScriptTimeout = 180

%>
<SCRIPT language=JavaScript>
<!--
function doConvert(){
window.open("ConvertDetailTrade2.asp?Search_Instrument=<%=Search_Instrument%>&Search_Market=<%=Search_Market%>&From_Day=<%=Search_From_Day%>&From_Month=<%=Search_From_Month%>&From_Year=<%=Search_From_Year%>&To_day=<%=Search_To_Day%>&To_Month=<%=Search_To_Month%>&To_Year=<%=Search_To_Year%>"); 

}

//-->
</SCRIPT>