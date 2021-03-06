<% 
'*********************************************************************************
'NAME       : ClientSummaryList.asp           
'DESCRIPTION: Client Summary List (adopted from Client summary report)
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 100222 Roger Wong   Prototype
'MODIFIED   : 
'							
'********************************************************************************

%>

<%
On Error resume Next
%>

<!--#include file="include/SessionHandler.inc.asp" -->
<%


Conn.CommandTimeout=0
Conn.ConnectionTimeout=0

Dim MSG_BUSY
MSG_BUSY = "The selection criteria are too broad for the system to process. Please use more specific selection criteria then resubmit."

if session("shell_power")="" then
	response.redirect "logout.asp?r=1"
end if

Title = "Client Summary"




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
set RsSettle = server.createobject("adodb.recordset")

Dim itotalturnover, itotalconsideration, itotalbrokerage, itotalCCY
Dim iPageturnover(), iPageconsideration(), iPagebrokerage(), iPageCCY()
itotalturnover = 0
itotalconsideration = 0
itotalbrokerage = 0

strURL = Request.ServerVariables("URL") ' Retreive the URL of this page from Server Variables

Dim rsChangeFlag
Dim rsClientCode
Dim rsClientName
Dim rsPortValue
Dim rsMarginValue
Dim rsMarginCall
Dim rsPercentUsed
Dim rsCashValue
Dim rsNetValue
Dim rsAvailableBalance
Dim rsT1HKD
Dim rsT2HKD
Dim rsT3HKD
Dim rsIntAccrued
Dim rsMarginLimit
Dim rsTradingLimit
Dim rsChequeName
Dim rsBankInfo
Dim rsSettleMethod
Dim rsAcctType
%>





<%
if session("shell_power")="" then
  response.redirect "Default.asp"
end if


%>


<html>
<head>

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
		if ((datevalidate(document.fm1.FromDay.value, document.fm1.FromMonth.value -1, document.fm1.FromYear.value) == false) )
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
				

		//Balance greater must be numeric 
		if (isNaN(document.fm1.BalanceGreater.value) == true ){
  			alert("Please enter a numeric number");
        	document.fm1.BalanceGreater.focus();
        	return false;
		}

		//Balance less must be numeric 
		if (isNaN(document.fm1.BalanceLess.value) == true ){
  			alert("Please enter a numeric number");
        	document.fm1.BalanceLess.focus();
        	return false;
		}



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

function PopupSearchAE() {
		newwindow=window.open( "SearchAE.asp?sid=<%=SessionID%>", "myWindow", 
									"status = 1, height = 300, width = 800, resizable = 1'"  )
		 if (window.focus) {
           newwindow.focus();
       }
 			
}

function PopupWindow() {
		newwindow=window.open( "SearchClientNumber.asp?sid=<%=SessionID%>", "myWindow", 
									"status = 1, height = 300, width = 800, resizable = 1'"  )
		 if (window.focus) {
           newwindow.focus();
       }
 			
}

//-->
</SCRIPT>


</head>

<body leftmargin="0" topmargin="0" OnLoad="document.fm1.submitted.value=0;document.fm1.ClientFrom.focus();" onkeypress="return disableCtrlKeyCombination(event);" onkeydown="return disableCtrlKeyCombination(event);" >



<span class="noprint">
	
	
<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">


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
Dim Search_Market
Dim Search_AccountType
Dim Search_Instrument
Dim Search_Order 
Dim Search_Direction
Dim Search_SharedSelection 
Dim Search_SharedGroup
Dim Search_SharedGroupMember
Dim Search_Amount_Type
Dim Search_balance_greater
Dim Search_balance_less


Search_AEGroup	    = Request.form("GroupID")
Search_ClientFrom       = Request.form("ClientFrom")
Search_ClientTo         = Request.form("ClientTo")
Search_AEFrom           = Request.form("AEFrom")
Search_AETo             = Request.form("AETo")
Search_From_Day         = Request.form("FromDay")
Search_From_Month       = Request.form("FromMonth")
Search_From_Year        = Request.form("FromYear")
Search_Transaction_Type = Request.form("TranType")
Search_Market           = Request.form("Market")
Search_Instrument       = Request.form("Instrument")
Search_Order            = Request.form("Order")
Search_AccountType      = Request.form("AccountType")
Search_Amount_type      = Request.form("AmountType")
Search_balance_greater  = Request.form("BalanceGreater")
Search_balance_less  = Request.form("BalanceLess")
Search_Direction        = Request.form("Direction")
Search_SharedSelection  = Request.form("ShareSelection")	
Search_SharedGroup      = Request.form("SharedGroup")
Search_SharedGroupMember= Request.form("SharedGroupMember")


'response.write "1 " & Search_AEGroup	    & "<BR>"
'response.write "2 " & Search_ClientFrom   & "<BR>"
'response.write "3 " & Search_ClientTo     & "<BR>"
'response.write "4 " & Search_AEFrom       & "<BR>"
'response.write "5 " & Search_AETo         & "<BR>"
'response.write "6 " & Search_From_Day     & "<BR>"
'response.write "7 " & Search_From_Month   & "<BR>"
'response.write "8 " & Search_From_Year    & "<BR>"
'response.write "9 " & Search_Transaction_Type & "<BR>"
'response.write "10 " & Search_Market           & "<BR>"
'response.write "11 " & Search_Instrument       & "<BR>"
'response.write "12 " & Search_Order            & "<BR>"
'response.write "13 " & Search_AccountType      & "<BR>"
'response.write "14 " & Search_Amount_type      & "<BR>"
'response.write "15 " & Search_balance_greater  & "<BR>"
'response.write "16 " & Search_balance_less  & "<BR>"
'response.write "17 " & Search_Direction        & "<BR>"
'response.write "18 " & Search_SharedSelection  & "<BR>"
'response.write "19 " & Search_SharedGroup      & "<BR>"
'response.write "20 " & Search_SharedGroupMember & "<BR>"



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
If Request.form("page") = "" Then
	iPageCurrent = 1
	Search_Order = "TRADEDATE"
	Search_Direction = "ASC"
	Search_From_Day = day(Session("DBLastModifiedDateValue"))
	Search_From_Month = month(Session("DBLastModifiedDateValue"))
	Search_From_Year = year(Session("DBLastModifiedDateValue"))



	Search_AEGroup	    = session("GroupID")
	Search_ClientFrom   = session("ClientFrom")
	Search_ClientTo     = session("ClientTo")
	Search_AEFrom       = session("AEFrom")
	Search_AETo         = session("AETo")
	
	Search_SharedSelection = "share1"

Else
	iPageCurrent = Clng(Request.form("page"))
End If


'pass all selection to session 
session("GroupID")               =  Search_AEGroup	               
session("ClientFrom")            =  Search_ClientFrom              
session("ClientTo")              =  Search_ClientTo                
session("AEFrom")                =  Search_AEFrom                  
session("AETo")                  =  Search_AETo                    


set RsMarket = server.createobject("adodb.recordset")
RsMarket.open ("Exec Retrieve_AvailableMarket ") ,  Conn,3,1


set RsGroupID = server.createobject("adodb.recordset")
RsGroupID.open ("Exec Retrieve_AvailableGroupID ") ,  Conn,3,1


'**************
' Sub procedures
'**************

sub OrderVariable(iorder)
'response.write iorder & "," & Search_Order 
  'User click the same field
	if iorder = Search_Order then 
		'reverse the direction
		if Search_Direction = "ASC" then
			response.write "'" & iorder & "','DESC'"
		else
			response.write "'" & iorder & "','ASC'"
		end if
	else
		'User click a different field, default is ascending order
		response.write "'" & iorder & "','ASC'"
	end if 

End sub 

sub OrderImage(iorder)
  'User click the same field
	if iorder = Search_Order then 
		'reverse the direction
		if Search_Direction = "ASC" then
			response.write "<img border=0 src='images/up.jpg'>" 
		else
			response.write "<img border=0 src='images/down.jpg'>" 
		end if
	else
		'User click a different field, default is ascending order
		' do nothing
	end if 
		
end sub

%>
 </span>
<%
'*****************************************************************
' Start of form
'*****************************************************************
%>

<form name="fm1" method="post" action="">
  <table width="97%" border="0" class="normal">

			<% if ( session("shell_power") = 3   or session("shell_power") = 4   or session("shell_power") = 8 )	 then %>
			<tr> 
				<td width="20%" >Branch:</td> 
				<td width="30%" >
				 
				<select size="1" name="GroupID" class="common">
						<option value="" <% if Search_AEGroup="ALL" then response.write "selected" %> >All</option>
						<%
								do while (  Not RsGroupID.EOF)
						%>
								<option value="<%=RsGroupID("GroupID")%>" <% if Search_AEGroup=cstr(RsGroupID("GroupID")) then response.write "selected" %> ><%=RsGroupID("Name")%></option>
						
						<%
								RsGroupID.movenext
								Loop
						%>
				</select></td>
				<td align="right" colspan="2"><font color="red">*</font> Denotes a mandatory field
				</td>
			</tr>
			<% End If %>

			<tr>
					<td width="20%">Client Number: (From)<font color="red">*</font> </td> 
					<td width="30%">
					<input name="ClientFrom" type=text value="<%= Search_ClientFrom %>" size="15">
                    <img align="top" style="cursor:pointer" onClick="PopupWindow()" src="images/search.gif"> </td>
					
      <td width="20%">Client Number: (To)</td> 
      <td width="30%">
      	     
<input name="ClientTo" type=text value="<% If Search_ClientFrom <> Search_ClientTo Then Response.Write Search_ClientTo End If%>" size="15">
</td>
    </tr>
    
<% if session("shell_power") >=3 then %>

	<tr>
      <td width="20%">AE Code: (From)</td> 
      <td width="30%">
      	     
<input name="AEFrom" type=text value="<%= Search_AEFrom %>" size="15">
<img align="top" style="cursor:pointer" onClick="PopupSearchAE()" src="images/search.gif"></td>
    
      <td width="20%">AE Code: (To)</td> 
      <td>
      	     
<input name="AETo" type=text value="<% If Search_AEFrom <> Search_AETo Then Response.Write Search_AETo End If%>" size="15"></td>
    </tr>

 
<% End If %>
 
 <tr> 
      <td width="20%">Date of Report:</td> 
      <td width="30%">
      	     
			<select name="FromDay" class="common">
			<option value="1" <% if Search_From_Day=1 then response.write "selected"%>>1</option>
			<option value="2" <% if Search_From_Day=2 then response.write "selected"%>>2</option>
			<option value="3" <% if Search_From_Day=3 then response.write "selected"%>>3</option>
			<option value="4" <% if Search_From_Day=4 then response.write "selected"%>>4</option>
			<option value="5" <% if Search_From_Day=5 then response.write "selected"%>>5</option>
			<option value="6" <% if Search_From_Day=6 then response.write "selected"%>>6</option>
			<option value="7" <% if Search_From_Day=7 then response.write "selected"%>>7</option>
			<option value="8" <% if Search_From_Day=8 then response.write "selected"%>>8</option>
			<option value="9" <% if Search_From_Day=9 then response.write "selected"%>>9</option>
			<option value="10" <% if Search_From_Day=10 then response.write "selected"%>>10</option>
			<option value="11" <% if Search_From_Day=11 then response.write "selected"%>>11</option>
			<option value="12" <% if Search_From_Day=12 then response.write "selected"%>>12</option>
			<option value="13" <% if Search_From_Day=13 then response.write "selected"%>>13</option>
			<option value="14" <% if Search_From_Day=14 then response.write "selected"%>>14</option>
			<option value="15" <% if Search_From_Day=15 then response.write "selected"%>>15</option>
			<option value="16" <% if Search_From_Day=16 then response.write "selected"%>>16</option>
			<option value="17" <% if Search_From_Day=17 then response.write "selected"%>>17</option>
			<option value="18" <% if Search_From_Day=18 then response.write "selected"%>>18</option>
			<option value="19" <% if Search_From_Day=19 then response.write "selected"%>>19</option>
			<option value="20" <% if Search_From_Day=20 then response.write "selected"%>>20</option>
			<option value="21" <% if Search_From_Day=21 then response.write "selected"%>>21</option>
			<option value="22" <% if Search_From_Day=22 then response.write "selected"%>>22</option>
			<option value="23" <% if Search_From_Day=23 then response.write "selected"%>>23</option>
			<option value="24" <% if Search_From_Day=24 then response.write "selected"%>>24</option>
			<option value="25" <% if Search_From_Day=25 then response.write "selected"%>>25</option>
			<option value="26" <% if Search_From_Day=26 then response.write "selected"%>>26</option>
			<option value="27" <% if Search_From_Day=27 then response.write "selected"%>>27</option>
			<option value="28" <% if Search_From_Day=28 then response.write "selected"%>>28</option>
			<option value="29" <% if Search_From_Day=29 then response.write "selected"%>>29</option>
			<option value="30" <% if Search_From_Day=30 then response.write "selected"%>>30</option>
			<option value="31" <% if Search_From_Day=31 then response.write "selected"%>>31</option>
		
			
			</select>


			<select name="FromMonth" class="common">            	
					<option value="1" <% if Search_From_Month=1 then response.write "selected"%>>Jan</option>
					<option value="2" <% if Search_From_Month=2 then response.write "selected"%>>Feb</option>
					<option value="3" <% if Search_From_Month=3 then response.write "selected"%>>Mar</option>
					<option value="4" <% if Search_From_Month=4 then response.write "selected"%>>Apr</option>
					<option value="5" <% if Search_From_Month=5 then response.write "selected"%>>May</option>
					<option value="6" <% if Search_From_Month=6 then response.write "selected"%>>Jun</option>
					<option value="7" <% if Search_From_Month=7 then response.write "selected"%>>Jul</option>
					<option value="8" <% if Search_From_Month=8 then response.write "selected"%>>Aug</option>
					<option value="9" <% if Search_From_Month=9 then response.write "selected"%>>Sep</option>
					<option value="10" <% if Search_From_Month=10 then response.write "selected"%>>Oct</option>
					<option value="11" <% if Search_From_Month=11 then response.write "selected"%>>Nov</option>
					<option value="12" <% if Search_From_Month=12 then response.write "selected"%>>Dec</option>
			</select>


			<select name="FromYear" class="common">   
<% 


Year_starting = Year(DateAdd("yyyy", -1, Now()))
year_ending = Year(Now())

for i=Year_starting to Year_ending
%>			         
			<option value="<%=i%>" <% if clng(i)=clng(Search_From_Year) then response.write "selected"%>><%=i%></option>

<% next %>

			</select> </td>
     
      <td></td>
      <td></td>
    
    </tr>
    
 <tr> 
	<td width="20%">Market:</td> 
	<td width="30%">
	 
	<select size="1" name="Market" class="common">
			<option value="ALL" <% if Search_Market="ALL" then response.write "selected" %> >All</option>
			<%
					do while (  Not rsMarket.EOF)
			%>
					<option value="<%=rsMarket("Market")%>" <% if Search_Market=rsMarket("Market") then response.write "selected" %> ><%=rsMarket("Market")%></option>
			
			<%
					rsMarket.movenext
					Loop
			%>
	</select></td>
	
      <td width="20%">Instrument:</td> 
      <td>
     <input name="Instrument" type=text value="<%= Search_Instrument %>" size="15"></td>   
 	     
    </tr>
 <tr> 
	<td width="20%">Ledger Balance Greater Than:</td> 
      <td>
     <input name="BalanceGreater" type=text value="<%= Search_balance_greater %>" size="15"></td> 
	
      <td width="20%">Ledger Balance Less Than:</td> 
      <td>
     <input name="BalanceLess" type=text value="<%= Search_balance_less %>" size="15"></td> 
 	     
    </tr>

 <tr> 

 <tr> 
	<td width="20%">Ledger Balance Type:</td> 
				<td width="30%">
				 
			<select size="1" name="AmountType" class="common">
				<option value="ALL" <% if Search_Amount_Type="ALL" then response.write "selected" %> >All</option>
				<option value="DEBIT" <% if Search_Amount_Type="DEBIT" then response.write "selected" %>>Debit</option>
				<option value="CREDIT" <% if Search_Amount_Type="CREDIT" then response.write "selected" %>>Credit</option>
			</select>
			</td>
	
      <td width="20%">Include Margin A/C:</td> 
      <td>
			<select size="1" name="AccountType" class="common">
			<option value="ALL" <% if Search_AccountType="ALL" then response.write "selected" %> >All</option>
			<option value="MRGN" <% if Search_AccountType="MRGN" then response.write "selected" %> >Margin</option>
			</select>

	</td>	     
    </tr>

    
<%if session("SharedGroup") > 0 then 


'List Shared Group member
set RsSharedGroupMember = server.createobject("adodb.recordset")
RsSharedGroupMember.open ("Exec List_SharedGroupMember '"&Session("id")&"', '"&Session("shell_power")&"' ") ,  Conn,3,1


'List shared group 
set RsSharedGroup = server.createobject("adodb.recordset")
RsSharedGroup.open ("Exec List_SharedGroup '"&Session("id")&"' ") ,  Conn,3,1





%>		  
		<tr> 
			<td colspan="4">&nbsp;<input type="radio" name="ShareSelection" value="share1" onClick=""  <%if Search_SharedSelection = "share1" then response.write "checked" end if %>  > Viewing AE only
			|
			<input type="radio" name="ShareSelection" value="share2" onClick=""  <%if Search_SharedSelection = "share2" then response.write "checked" end if %>   > Particular AE in the Sales Team&nbsp;

			<select name="SharedGroupMember" class="common">
						<%
								do while (  Not RsSharedGroupMember.EOF)
						%>
								<option value="<%=RsSharedGroupMember("loginname")%>" <% if Search_SharedGroupMember=RsSharedGroupMember("loginname") then response.write "selected" end if %> > <%=RsSharedGroupMember("loginname")%></option>
						<%
								RsSharedGroupMember.movenext
								Loop
						%>
			</select>	
			
			|	
			
<input type="radio" name="ShareSelection" value="share3" onClick=""  <%if Search_SharedSelection = "share3" then response.write "checked" end if %>   > All AEs in the Sales Team&nbsp;
		
			<select name="SharedGroup" class="common">
						<%
								do while (  Not RsSharedGroup.EOF)
						%>
								<option value="<%=RsSharedGroup("groupid")%>" 
										<% 
											if IsNumeric(Search_SharedGroup) then 
												response.write Search_SharedGroup
												if  (cint(Search_SharedGroup) = cint(RsSharedGroup("groupid"))) then 
													response.write "selected" 
												end if 
											end if
										%>  
										> <%=RsSharedGroup("name")%> </option>
						<%
								RsSharedGroup.movenext
								Loop
						%>
			</select>				

			</td>
		</tr>  

<% end if %>  		
		
		<tr> 
			<td></td>
			<td colspan="3">
  	<input type=hidden   value="<%=iPageCurrent%>"   name="page"> 
 	<input type=hidden   value="<%=Search_Order%>"   name="Order"> 
 	<input type=hidden   value="<%=Search_Direction%>"   name="Direction"> 
 	<input type=hidden   name="submitted"> 

          <input id="Submit1" type="button" value="Submit" onClick="dosubmit(1);"></td>

		</tr>    

    </table>
</form>    
<%
'*****************************************************************
' End of form
'*****************************************************************
%>


<%
'*****************************************************************
' Start of report body
'*****************************************************************
%>    
    
<%


If Request.form("submitted") = 0 Then


'**********
' If no argument
'**********

'do nothing
  
  
else 

     


'**********
' If passing arguments
'**********



 set Rs1 = server.createobject("adodb.recordset")
 

 'Rs return 2 value
 '1) Total number of matched client
 '2) all records for targeted client

	Select Case  Search_SharedSelection
	case "share2"
		'shared group member
         'response.write "Shared"

 			
 			Rs1.open ("Exec Retrieve_ClientSummaryDetail_GroupBy_Client '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '',  '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"',  '"&Search_Market&"','"&Search_Instrument&"',  '"&Search_Amount_type&"', '"&Search_balance_greater&"','"&Search_balance_less&"','"&Search_AccountType&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  Conn,3,1

   			
			
	case "share3"

 			Rs1.open ("Exec Retrieve_ClientSummaryDetail_GroupBy_Client '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_Market&"','"&Search_Instrument&"',  '"&Search_Amount_type&"', '"&Search_balance_greater&"','"&Search_balance_less&"','"&Search_AccountType&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  Conn,3,1
			
			
	case else
		'normal
 
 			Rs1.open ("Exec Retrieve_ClientSummaryDetail_GroupBy_Client '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_Market&"','"&Search_Instrument&"',  '"&Search_Amount_type&"', '"&Search_balance_greater&"','"&Search_balance_less&"','"&Search_AccountType&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  Conn,3,1


response.write "11 " & iPageCurrent & "<BR>"
response.write "22 " & RECORDPERPAGE & "<BR>"
response.write "33 " & Search_Order & "<BR>"
response.write "44 " & Search_Direction & "<BR>"
response.write "55 " & Search_Amount_type & "<BR>"
response.write "66 " & Search_balance_greater & "<BR>"
response.write "77 " & Search_balance_less & "<BR>"
response.write "88 " & Search_AccountType & "<BR>"
			
	end select	

 			'response.write ("Exec Retrieve_ClientSummaryDetail_GroupBy_Client '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_Market&"','"&Search_Instrument&"',  '"&Search_Amount_type&"', '"&Search_balance_greater&"','"&Search_balance_less&"','"&Search_AccountType&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") 
        




Search_AccountType      = Request.form("AccountType")
Search_Amount_type      = Request.form("AmountType")
Search_balance_greater  = Request.form("BalanceGreater")
Search_balance_less  	= Request.form("BalanceLess")


  if Rs1.EoF then
		
		If Err.Number <> 0 then
			
			'SQL connection error handler
			'response.write  "<table><tr><td class='RedClr'>The server is currently too busy to process your request right now. Please wait a moment and then try again. If the problem persists, please contact systems administrator.<br></td></tr></table>"
			
			'no record found
			response.write ("No record found")

		else
			'no record found
			response.write ("No record found")
				
		End If

	else
 

%>    
   
<div id="reportbody1" >


   
<%
'**********
' Start of page navigation 
'**********
%> 
    
<table width="97%" border="0" class="normal"  cellspacing="1" cellpadding="4">
<tr bgcolor="#FFFFCC"> 
<td  width="20%">　</td>
      <td align="center"><br><u>Client Summary</u></td> 
      <td align="right" width="20%">
          <span class="noprint">
							<%if Trim(PrintAllowed) = 1 then %>  
							<a href="javascript:window.print()">Friendly Print</a><% 'end if %> &nbsp;&nbsp;
							<% End If %>
			
          </span>      	
			</td>
</tr>
</table>
<br>
<table width="97%" border="0" class="sortable" style="border-width: 0;FONT-SIZE: 11px;TEXT-ALIGN: Right;FONT-FAMILY: Verdana, 'MS Sans Serif', Arial" bgcolor="#808080" cellspacing="1" cellpadding="2">
<thead>

		<tr bgcolor="#ADF3B6" align="center">
		   <td><span style="cursor:hand">Client Code <br> 客戶編號</span></td>
		   <td><span style="cursor:hand">Client Name<br>客戶姓名 </span></td>
		   <td><span style="cursor:hand">Portfolio Value (HKD)<br>股份組合市值  </span></td>
		   <td><span style="cursor:hand">Marginable Value (HKD) <br>按倉市值</span></td>
		   <td><span style="cursor:hand">Available / (Margin Call) <br>可按倉值價</span></td>
		
		   <td><span style="cursor:hand">% Used <br>已使用%</span></td>
		   <td><span style="cursor:hand">Cash Value / Ledger Balance (HKD) <br>現金值</span></td>
		   <td width="300"><span style="cursor:hand">Net Value (HKD)<br>總值</span></td>
		   <td width="300"><span style="cursor:hand">Available Balance (HKD)<br>可用結餘</span></td> 
		   <td><span style="cursor:hand">T+1 (HKD)<br>交易當日</span></td> 
		
		   <td><span style="cursor:hand">T+2 (HKD)<br>交易次日</span></td> 
		   <td><span style="cursor:hand">T+3 or beyond (HKD)<br>交易之後第三日</span></td> 
		   <td><span style="cursor:hand">Int Accrued (HKD)<br>應計利息</span></td> 
		   <td><span style="cursor:hand">Margin Limit<br>按倉限額</span></td> 
		   <td><span style="cursor:hand">Trading Limit<br>交易限額</span></td> 
		
		   <td><span style="cursor:hand">Cheque Name<br>支票名稱</span></td> 
		   <td><span style="cursor:hand">Bank Info<br>銀行資料</span></td> 
		   <td><span style="cursor:hand">Settlement Method<br>結算方法</span></td> 
		</tr>
</thead>
<tbody>

		<%
			dim iPageCCYcount
			dim k
			dim iPageUpdate
			
		'	ReDim Preserve iPageCCY(1)
		'	ReDim Preserve iPageturnover(1)
		'	ReDim Preserve iPageconsideration(1)
		'	ReDim Preserve iPagebrokerage(1)
			
		'	iPageCCYcount = 0
		'	iPageturnover(0) = 0
		'	iPageconsideration(0)= 0
		'	iPagebrokerage(0)= 0
			'iPageCCY(0) = ""
			

			
			
			do while (Not rs1.EOF)

				'clean all variable

					erase rsClientCode
					erase rsClientName
					erase rsPortValue
					erase rsMarginValue
					erase rsMarginCall
					erase rsPercentUsed
					erase rsCashValue
					erase rsNetValue
					erase rsAvailableBalance
					erase rsT1HKD
					erase rsT2HKD
					erase rsT3HKD
					erase rsIntAccrued
					erase rsMarginLimit
					erase rsTradingLimit
					erase rsChequeName
					erase rsBankInfo
					erase rsBankName
					erase rsSettleMethod		
					erase rsAcctType
					rsNetValue=0
					
					rsClientCode = rs1("clnt")
					rsMarginLimit   =  rs1("MarginLimit")
					rsTradingLimit  =  rs1("tradingLimit")

                              
                    
					rsSettleMethod  =  rs1("SettleMethod")
					rsChequeName    =  rs1("chequeName")				
					rsIntAccrued       = rs1("IntHKD")
					rsCashValue        = rs1("CashValue")
					rsAcctType         = rs1("Accttype")
					rsAvailableBalance = rs1("T0HKD")
					rsT1HKD            = rs1("T1HKD")
					rsT2HKD            = rs1("T2HKD")
					rsT3HKD            = rs1("T3HKD")
					rsClientName    = rs1("clntname") 
					rsMarginCall       = rs1("MarginCall")
					rsPortValue     = rs1("PortValue")
					rsMarginValue      = rs1("MarginValue")

'response.write "Margin Limit :     " & rsMarginLimit & "<br>"
'response.write " Marginable Value :    " & rsMarginValue & "<br>"
'response.write " Cash Value :    " & rsCashValue & "<br>"


					'Calculate % Used
                                        TempMarginValue = 0


                                         if   cDbl(rsMarginLimit) < cDbl(rsMarginValue) and formatnumber(rsMarginLimit) <> 0 then
 
                                          TempMarginValue =  rsMarginLimit

                                         else 

                                          TempMarginValue =  rsMarginValue

                                         end if      

'response.write " TempMarginValue :    " & TempMarginValue & "<br>"
                                    

				        if (formatnumber(rsCashValue,2) < 0 and formatnumber(rsCashValue,2) <> 0 ) then

						        rsPercentUsed      = abs(cDbl(rsCashValue) / cDbl(TempMarginValue) * 100)
                                                      
										
					else
						rsPercentUsed = 0
					end if

					rsNetValue         = cDbl(rs1("CashValue")) + cDbl(rs1("PortValue"))
					rsBankInfo         = rs1("BankInfo")
					rsBankName         = rs1("BankName")

			


 
			%>
					<tr bgcolor="#FFFFCC"> 
					   <td align="left"><a href="ClientSummaryListDetail.asp?PrintAllowed=<%=Trim(PrintAllowed)%>&DisplayFirst=<%=rsClientCode%>&CurrentClient=<%=rsClientCode%>&ClientFrom=<%=Search_ClientFrom%>&ClientTo=<%=Search_ClientTo%>&AEFrom=<%=search_AEFrom%>&AETo=<%=search_AETo%>&Instrument=<%=Search_Instrument%>&Market=<%=Search_Market%>&SDay=<%=Search_From_Day%>&SMonth=<%=Search_From_Month%>&SYear=<%=Search_From_Year%>&submitted=1&sid=<%=SessionID%>#DisplayFirst" target=_blank><%=rsClientCode %></a><img border=0 src='images/tel.gif' onClick="PopupClientContact('<%=rs1("clnt") %>')"></img></td>
					   <td align="left"><%=rsClientName %></td>
					   <td style="white-space: nowrap"><%=formatnumber(rsPortValue,2)  %></td> 
					   <td style="white-space: nowrap"><% if rsAcctType = "MRGN" then response.write formatnumber(rsMarginValue,2) end if %></td> 
					   <td style="white-space: nowrap"><% if rsAcctType = "MRGN" then response.write formatnumber(rsMarginCall,2)  end if %></td> 
		
					   <td style="white-space: nowrap"><% if rsAcctType = "MRGN" then response.write formatnumber(rsPercentUsed,2) end if %></td> 
					   <td style="white-space: nowrap"><%=formatnumber(rsCashValue,2)  %></td> 
					   <td style="white-space: nowrap"><%=formatnumber(rsNetValue,2)  %></td> 
					   <td style="white-space: nowrap"><%=formatnumber(rsAvailableBalance,2)  %></td> 


					   <td style="white-space: nowrap"><%=formatnumber(rsT1HKD,2) %></td> 
					   <td style="white-space: nowrap"><%=formatnumber(rsT2HKD,2)  %></td> 
					   <td style="white-space: nowrap"><%=formatnumber(rsT3HKD,2)  %></td> 
					   <td style="white-space: nowrap"><%=formatnumber(rsIntAccrued,2)  %></td> 
					   <td style="white-space: nowrap"><% if rsAcctType = "MRGN" then response.write formatnumber(rsMarginLimit,2) end if %></td> 
					   <td style="white-space: nowrap"><% if rsAcctType <> "MRGN" then response.write formatnumber(rsTradingLimit,2) end if %></td> 
		
		
					   <td align="left"><% =rsChequeName %></td> 
					   <td align="left">
					   	<% 
					   		if rsBankName <> "" then
					   			response.write 	rsBankName + " : " + rsBankInfo
					   		end if
					   	%> 
					   </td> 
					   <td align="left">
								<%
								RsSettle.open ("Exec Retrieve_ClientSummaryDetail_SettleMethod '"&rsClientCode&"'  ") ,  Conn,3,1
								
								do while (  Not RsSettle.EOF)
									response.write RsSettle("FieldValue") & "<br>" 
									RsSettle.movenext
								Loop
								
								RsSettle.Close
								
								
								%>

					   </td> 
					</tr>
		
									<%
					
									
					               				
					
							rs1.movenext	  
					
						loop
					
					
							Rs1.Close
							Set Rs1=Nothing
						
						%>
</tbody>





</table>
<br>
<br>





<% if (session("shell_power") = 1 or session("shell_power") = 5) then %>  
		</span>
<% end if %>

<span class="noprint">



<%
'**********
' End of page navigation 
'**********
%>


<%
	    
    end if 'record found if statement
end if   'having client number if statement
%>



<%
'*****************************************************************
' End of report body
'*****************************************************************
%>
                

              
</div>

</span>
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
 'Conn.Close
 'Set Conn = Nothing

set RsSettle = Nothing
%>
<SCRIPT language=JavaScript>
<!--
function doConvert(){
window.open("ConvertClientSummary.asp?Search_Instrument=<%=Search_Instrument%>&Search_Market=<%=Search_Market%>&From_Day=<%=Search_From_Day%>&From_Month=<%=Search_From_Month%>&From_Year=<%=Search_From_Year%>&To_day=<%=Search_To_Day%>&To_Month=<%=Search_To_Month%>&To_Year=<%=Search_To_Year%>"); 

}

//-->
</SCRIPT>