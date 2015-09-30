
<% 
'*********************************************************************************
'NAME       : TransactionHistory.asp           
'DESCRIPTION: Transaction History (Cash Voucher, Cashbank Movement, Instrument Trade)
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 090401 Gary Yeung   Prototype
'MODIFIED   : 090407 Roger Wong   Record and page control
'			:  090712 Roger Wong	    Add Shared Group
'                      :  090929  Gary Yeung       Add Excel Function
'********************************************************************************

%>

<%
'On Error resume Next
%>


<!--#include file="include/SessionHandler.inc.asp" -->
<%

Dim MSG_BUSY
MSG_BUSY = "The selection criteria are too broad for the system to process. Please use more specific selection criteria then resubmit."

if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

Title = "Transaction History"




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



strURL = Request.ServerVariables("URL") ' Retreive the URL of this page from Server Variables
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
<script src="include/common.js"></script>
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

function PopupClientContact(clientnumber) {
	 
		var str='ListClientContact.asp?sid=<%=SessionID%>&clientnumber=' + clientnumber
		
		newwindow=window.open(str , "myWindow", 
									"status = 1, height = 300, width = 600, resizable = 1'"  )
		 if (window.focus) {
           newwindow.focus();
       }
 			
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

function popup() {
window.open( "trancode.asp", "myWindow")
}


//-->
</SCRIPT>

</head>
<body leftmargin="0" topmargin="0" OnLoad="document.fm1.submitted.value=0;document.fm1.ClientFrom.focus();" >


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
Dim Search_To_Day
Dim Search_To_Month
Dim Search_To_Year
Dim Search_Transaction_Type
Dim Search_Market
Dim Search_Currency
Dim Search_Instrument
Dim Search_Account_type
Dim Search_Order 
Dim Search_Direction
Dim Search_SharedSelection 
Dim Search_SharedGroup
Dim Search_SharedGroupMember



Search_AEGroup	    = Request.form("GroupID")
Search_ClientFrom       = Request.form("ClientFrom")
Search_ClientTo         = Request.form("ClientTo")
Search_AEFrom           = Request.form("AEFrom")
Search_AETo             = Request.form("AETo")
Search_From_Day         = Request.form("FromDay")
Search_From_Month       = Request.form("FromMonth")
Search_From_Year        = Request.form("FromYear")
Search_To_Day           = Request.form("ToDay")
Search_To_Month         = Request.form("ToMonth")
Search_To_Year          = Request.form("ToYear")
Search_Transaction_Type = Request.form("TranType")
Search_Market           = Request.form("Market")
Search_Currency         = Request.form("Currency")
Search_Instrument       = Request.form("Instrument")
Search_Amount_type      = Request.form("AmountType")
Search_Order            = Request.form("Order")
Search_Direction        = Request.form("Direction")
Search_SharedSelection      =  Request.form("ShareSelection")	
Search_SharedGroup  = Request.form("SharedGroup")
Search_SharedGroupMember  = Request.form("SharedGroupMember")



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


'response.write session("id")



' If User enter From value only, change the "To" value to "From"
if Search_ClientTo = "" then
   Search_ClientTo = Search_ClientFrom
end if
if Search_AETo = "" then
   Search_AETo = Search_AEFrom
end if


'  default value of variable in first page
If Request.form("page") = "" Then
	iPageCurrent = 1
	Search_Order = "TRADEDATE"
	Search_Direction = "ASC"

	Search_From_Day = day(Session("DBLastModifiedDateValue"))
	Search_From_Month = month(Session("DBLastModifiedDateValue"))
	Search_From_Year = year(Session("DBLastModifiedDateValue"))
	Search_To_Day = day(Session("DBLastModifiedDateValue"))
	Search_To_Month = month(Session("DBLastModifiedDateValue"))
	Search_To_Year = year(Session("DBLastModifiedDateValue"))

	Search_AEGroup	    = session("GroupID")
	Search_ClientFrom   = session("ClientFrom")
	Search_ClientTo     = session("ClientTo")
	Search_AEFrom       = session("AEFrom")
	Search_AETo         = session("AETo")
	
	Search_SharedSelection = "share1"

Else
	iPageCurrent = Clng(Request.form("page"))
End If


set RsMarket = server.createobject("adodb.recordset")
RsMarket.open ("Exec Retrieve_AvailableMarket ") ,  StrCnn,3,1

set RsCCY = server.createobject("adodb.recordset")
RsCCY.open ("Exec Retrieve_AvailableCCY ") ,  StrCnn,3,1

set RsGroupID = server.createobject("adodb.recordset")
RsGroupID.open ("Exec Retrieve_AvailableGroupID ") ,  StrCnn,3,1

session("GroupID")               =  Search_AEGroup	               
session("ClientFrom")            =  Search_ClientFrom              
session("ClientTo")              =  Search_ClientTo                
session("AEFrom")                =  Search_AEFrom                  
session("AETo")                  =  Search_AETo                    


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

<%
'*****************************************************************
' Start of form
'*****************************************************************
%>
<form name="fm1" method="post" action="<%= strURL %>?sid=<%=SessionID%>">
  <table width="97%" border="0" class="normal">
					<% if ( session("shell_power") = 3   or session("shell_power") = 4   or session("shell_power") = 8)	 then %>
							<tr> 
								<td >Branch:</td> 
								<td >
								 
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
					<td width="30%"><input name="ClientFrom" type=text value="<%= Search_ClientFrom %>" size="15">
                    <img align="top" style="cursor:pointer" onClick="PopupWindow()" src="images/search.gif"> </td>
					<td width="20%">Client Number: (To)</td> 
					<td width="30%"><input name="ClientTo" type=text value="<% If Search_ClientFrom <> Search_ClientTo Then Response.Write Search_ClientTo End If%>" size="15"></td>
						
				</tr>
				
				<% if session("shell_power") >=3 then %>
				<tr>
							<td>AE Code: (From)</td> 
							<td><input name="AEFrom" type=text value="<%= Search_AEFrom %>" size="15">
              <img align="top" style="cursor:pointer" onClick="PopupSearchAE()" src="images/search.gif"></td>
						
							<td>AE Code: (To)</td> 
							<td><input name="AETo" type=text value="<% If Search_AEFrom <> Search_AETo Then Response.Write Search_AETo End If%>" size="15"></td>
							
				</tr>
				    
				<% End If %>        
 				<tr> 
				      <td >Period From:</td> 
				      <td >
				      	     
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
							
							
							Year_starting = Year(DateAdd("yyyy", -10, Now()))
							year_ending = Year(Now())
							
							for i=Year_starting to Year_ending
							%>			         
									<option value="<%=i%>" <% if clng(i)=clng(Search_From_Year) then response.write "selected"%>><%=i%></option>
				
							<% next %>
				
							</select> </td>
				     
      <td >Period To:</td> 
      <td >
      	     
		<select name="ToDay" class="common">
			<option value="1" <% if Search_To_Day=1 then response.write "selected"%>>1</option>
			<option value="2" <% if Search_To_Day=2 then response.write "selected"%>>2</option>
			<option value="3" <% if Search_To_Day=3 then response.write "selected"%>>3</option>
			<option value="4" <% if Search_To_Day=4 then response.write "selected"%>>4</option>
			<option value="5" <% if Search_To_Day=5 then response.write "selected"%>>5</option>
			<option value="6" <% if Search_To_Day=6 then response.write "selected"%>>6</option>
			<option value="7" <% if Search_To_Day=7 then response.write "selected"%>>7</option>
			<option value="8" <% if Search_To_Day=8 then response.write "selected"%>>8</option>
			<option value="9" <% if Search_To_Day=9 then response.write "selected"%>>9</option>
			<option value="10" <% if Search_To_Day=10 then response.write "selected"%>>10</option>
			<option value="11" <% if Search_To_Day=11 then response.write "selected"%>>11</option>
			<option value="12" <% if Search_To_Day=12 then response.write "selected"%>>12</option>
			<option value="13" <% if Search_To_Day=13 then response.write "selected"%>>13</option>
			<option value="14" <% if Search_To_Day=14 then response.write "selected"%>>14</option>
			<option value="15" <% if Search_To_Day=15 then response.write "selected"%>>15</option>
			<option value="16" <% if Search_To_Day=16 then response.write "selected"%>>16</option>
			<option value="17" <% if Search_To_Day=17 then response.write "selected"%>>17</option>
			<option value="18" <% if Search_To_Day=18 then response.write "selected"%>>18</option>
			<option value="19" <% if Search_To_Day=19 then response.write "selected"%>>19</option>
			<option value="20" <% if Search_To_Day=20 then response.write "selected"%>>20</option>
			<option value="21" <% if Search_To_Day=21 then response.write "selected"%>>21</option>
			<option value="22" <% if Search_To_Day=22 then response.write "selected"%>>22</option>
			<option value="23" <% if Search_To_Day=23 then response.write "selected"%>>23</option>
			<option value="24" <% if Search_To_Day=24 then response.write "selected"%>>24</option>
			<option value="25" <% if Search_To_Day=25 then response.write "selected"%>>25</option>
			<option value="26" <% if Search_To_Day=26 then response.write "selected"%>>26</option>
			<option value="27" <% if Search_To_Day=27 then response.write "selected"%>>27</option>
			<option value="28" <% if Search_To_Day=28 then response.write "selected"%>>28</option>
			<option value="29" <% if Search_To_Day=29 then response.write "selected"%>>29</option>
			<option value="30" <% if Search_To_Day=30 then response.write "selected"%>>30</option>
			<option value="31" <% if Search_To_Day=31 then response.write "selected"%>>31</option>
		
			
			</select>


			<select name="ToMonth" class="common">            	
					<option value="1" <% if Search_To_Month=1 then response.write "selected"%>>Jan</option>
					<option value="2" <% if Search_To_Month=2 then response.write "selected"%>>Feb</option>
					<option value="3" <% if Search_To_Month=3 then response.write "selected"%>>Mar</option>
					<option value="4" <% if Search_To_Month=4 then response.write "selected"%>>Apr</option>
					<option value="5" <% if Search_To_Month=5 then response.write "selected"%>>May</option>
					<option value="6" <% if Search_To_Month=6 then response.write "selected"%>>Jun</option>
					<option value="7" <% if Search_To_Month=7 then response.write "selected"%>>Jul</option>
					<option value="8" <% if Search_To_Month=8 then response.write "selected"%>>Aug</option>
					<option value="9" <% if Search_To_Month=9 then response.write "selected"%>>Sep</option>
					<option value="10" <% if Search_To_Month=10 then response.write "selected"%>>Oct</option>
					<option value="11" <% if Search_To_Month=11 then response.write "selected"%>>Nov</option>
					<option value="12" <% if Search_To_Month=12 then response.write "selected"%>>Dec</option>
			</select>


			<select name="ToYear" class="common">   
<% 


Year_starting = Year(DateAdd("yyyy", -10, Now()))
year_ending = Year(Now())

for i=Year_starting to Year_ending
%>			         
			<option value="<%=i%>" <% if clng(i)=clng(Search_To_Year) then response.write "selected"%>><%=i%></option>

<% next %>

			</select>
			</td>
      
    </tr>
    

    
 <tr> 
      <td >Transaction Type:</td> 
      <td >
      	     
<select size="1" name="TranType" class="common">
		<option value="ALL" <% if Search_Transaction_Type="ALL" then response.write "selected" %> >All</option>
                <option value="TRADE" <% if Search_Transaction_Type="TRADE" then response.write "selected" %> >TRADE</option>
		<option value="VOUCHER" <% if Search_Transaction_Type="VOUCHER" then response.write "selected" %> >Vouchers</option>
		<option value="INSTRUMENT" <% if Search_Transaction_Type="INSTRUMENT" then response.write "selected" %> >Instrument</option>
</select></td>
      
 
	<td >Market:</td> 
	<td >
	 
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
	
</tr>
    
<tr> 
	<td >Currency:</td> 
	<td >
	 
	<select size="1" name="Currency" class="common">
			<option value="ALL" <% if Search_Currency="ALL" then response.write "selected" %> >All</option>
			<%
					do while (  Not rsCCY.EOF)
			%>
					<option value="<%=rsCCY("CCY")%>" <% if Search_Currency=rsCCY("CCY") then response.write "selected" %> ><%=rsCCY("CCY")%></option>
			
			<%
					rsCCY.movenext
					Loop
			%>
	</select></td>
	
      <td >Instrument:</td> 
      <td >
      	     
<input name="Instrument" type=text value="<%= Search_Instrument %>" size="15"></td>
      
    </tr>
    

    
 <tr> 
      <td >Amount Type:</td> 
      <td >
      	     
<select size="1" name="AmountType" class="common">
	<option value="ALL" <% if Search_Amount_Type="ALL" then response.write "selected" %> >All</option>
	<option value="DEBIT" <% if Search_Amount_Type="DEBIT" then response.write "selected" %>>Debit</option>
	<option value="CREDIT" <% if Search_Amount_Type="CREDIT" then response.write "selected" %>>Credit</option>
</select>&nbsp;&nbsp;   
      <td colspan="2">
      	     
　</td>
    </tr>
    
<%if session("SharedGroup") > 0 then 


'List Shared Group member
set RsSharedGroupMember = server.createobject("adodb.recordset")
RsSharedGroupMember.open ("Exec List_SharedGroupMember '"&Session("id")&"', '"&Session("shell_power")&"' ") ,  StrCnn,3,1


'List shared group 
set RsSharedGroup = server.createobject("adodb.recordset")
RsSharedGroup.open ("Exec List_SharedGroup '"&Session("id")&"' ") ,  StrCnn,3,1





%>		  
		<tr> 
			<td colspan="4"><input type="radio" name="ShareSelection" value="share1" onClick=""  <%if Search_SharedSelection = "share1" then response.write "checked" end if %>  >&nbsp; Viewing AE only&nbsp; |<input type="radio" name="ShareSelection" value="share2" onClick=""  <%if Search_SharedSelection = "share2" then response.write "checked" end if %>   > Particular AE in the Sales Team&nbsp;
			

			<select name="SharedGroupMember" class="common">
						<%
								do while (  Not RsSharedGroupMember.EOF)
						%>
								<option value="<%=RsSharedGroupMember("loginname")%>" <% if Search_SharedGroupMember=RsSharedGroupMember("loginname") then response.write "selected" end if %> > <%=RsSharedGroupMember("loginname")%></option>
						<%
								RsSharedGroupMember.movenext
								Loop
						%>
			</select>&nbsp; | <input type="radio" name="ShareSelection" value="share3" onClick=""  <%if Search_SharedSelection = "share3" then response.write "checked" end if %>   > All AEs in the Sales Team&nbsp;  
			
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
	
	'dim Rs1 as adodb.recordset 
	
	'StrCnn.open myDSN 
	
	set Rs1 = server.createobject("adodb.recordset")
	
	'rs1.CursorLocation=3
	
' Check if it is one digit, for special purpose
    'If len(Trim(Search_ClientTo)) = 1 and Search_ClientTo <> 9 Then
        ' Search_ClientTo = Search_ClientTo + 1
   ' Elseif Search_ClientTo = 9 Then
          'Search_ClientTo = 999999999
    'End if

 'Rs return 2 value
 '1) Total number of matched client
 '2) all records for targeted client

iRecord = iPageCurrent 

 

' Rs1.open (" Exec retrieve_transactionhistory '0','z','701','701','',9,3,2009,9,3,2009,'ALL','','','1','5','CLNT','ASC' ") ,  StrCnn,3,1

 	Select Case  Search_SharedSelection
	case "share2"
		'shared group member
			'response.write  ("Exec retrieve_TransactionHistory '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Transaction_Type&"','"&Search_Market&"','"&Search_Currency&"','"&Search_Instrument&"','"&Search_Amount_type&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") 
 			Rs1.open ("Exec retrieve_TransactionHistory '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Transaction_Type&"','"&Search_Market&"','"&Search_Currency&"','"&Search_Instrument&"','"&Search_Amount_type&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
 			
	case "share3"
			'response.write  ("Exec retrieve_TransactionHistory '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Transaction_Type&"','"&Search_Market&"','"&Search_Currency&"','"&Search_Instrument&"','"&Search_Amount_type&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ")

 			Rs1.open ("Exec retrieve_TransactionHistory '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Transaction_Type&"','"&Search_Market&"','"&Search_Currency&"','"&Search_Instrument&"','"&Search_Amount_type&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
 			
			

	case else
		'normal
			'response.write  ("Exec retrieve_TransactionHistory '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Transaction_Type&"','"&Search_Market&"','"&Search_Currency&"','"&Search_Instrument&"','"&Search_Amount_type&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") 
 
 			Rs1.open ("Exec retrieve_TransactionHistory '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Transaction_Type&"','"&Search_Market&"','"&Search_Currency&"','"&Search_Instrument&"','"&Search_Amount_type&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
			
		
	end select	
	





	'assign total number of pages
	iRecordCount = rs1(0)


  if iRecordCount <= 0 then
		
		If Err.Number <> 0 then
			
			'SQL connection error handler
			response.write  "<table><tr><td class='RedClr'>" & MSG_BUSY & "<br></td></tr></table>"
		response.write  "<table><tr><td class='RedClr'>" & Err.Number & "<br></td></tr></table>"
	else
			'no record found
			response.write ("No record found")
				
		End If
		
	
	else
		'record found
		
		'response.write iRecordCount 
		
		'cal total no of pages
		iPageCount = int(iRecordCount / RECORDPERPAGE)+1
		
		'move to next recordset
  	Set rs1 = rs1.NextRecordset() 
 
%>    
    
<div id="reportbody1" >

<script type="text/javascript">
var somediv=document.getElementById("reportbody1")
disableSelection(somediv) //disable text selection within DIV with id="mydiv"
</script>

</span>
	    
    <DIV align=center>

  <TABLE border=0 cellPadding=0 cellSpacing=0 height=100% width=99%>

 <tr> 
 <td align="right" height="28" class="NavaMenu" >

	<%if PrintAllowed = 1 then %>  
							<a href="javascript:window.print()">Friendly Print</a>&nbsp;
							<% end if %>
							<%if (session("shell_power") = 8 ) then %>
							<a href="javascript:window.doConvert()">&nbsp;Excel</a>
						<% end if %>

&nbsp;&nbsp;

		<%
'**********
' Start of page navigation 
'**********

response.write (iPageCurrent & " Pages " & iPageCount &"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" )

'First button
%>
	<a href=javascript:dosubmit(1) style='cursor:hand'>First</a>

<%
' Prev button
If iPageCurrent > 1 Then
	%>
	<a href=javascript:dosubmit(<%= iPageCurrent-1 %>) style='cursor:hand'>Previous</a>
<% else %>
Previous
	<%
End If


'Next button
If iPageCurrent < iPageCount Then
	%>
	<a href=javascript:dosubmit(<%= iPageCurrent+1 %>) style='cursor:hand'>Next</a>
<% else %>
Next
	<%
End If
%>

<%
'Last button
%>

<a href=javascript:dosubmit(<%= iPageCount %>) style='cursor:hand'>Last</a>

</td></tr></table>


<%
'**********
' End of page navigation 
'**********
%>



    



<table width="99%" border="0" class="normal" cellspacing="1" cellpadding="2">
<tr bgcolor="#FFFFCC">
      <td align="center">帳戶記錄<br><u>Transaction History</u></td>
</tr>
</table>

<br>

<table  width="100%" border="0" class="sortable" style="border-width: 0;FONT-SIZE: 11px;TEXT-ALIGN: Right;FONT-FAMILY: Verdana, 'MS Sans Serif', Arial" bgcolor="#808080" cellspacing="1" cellpadding="2">

<thead>

<tr bgcolor="#ADF3B6">
      <td width="8%"><a href=javascript:ordersubmit(<% call OrderVariable("TRADEDATE")  %>) >Trade Date<br>交易日期</a></td>
      <td width="7%"><a href=javascript:ordersubmit(<% call OrderVariable("CLNT")  %>) >Client No.<br>客戶編號</a></td>
      <td width="10%"><a href=javascript:ordersubmit(<% call OrderVariable("CLNTNAME")  %>) >Client Name<br>客戶名稱</a></td>
   <td width="10%"> <a href=javascript:ordersubmit(<% call OrderVariable("VALUEDATE")  %>) >Value Date<br>評價日期</a></td>
   <td width="10%"><a href=javascript:popup() > (A full list of trans. type 交易類型清單) </a> <p><a href=javascript:ordersubmit(<% call OrderVariable("TRADETYPE")  %>) >Trans Type<br>交易類型</a></td>
   <td width="10%"><a href=javascript:ordersubmit(<% call OrderVariable("MARKET")  %>) >Market<br>市場</a></td>
   <td width="10%"><a href=javascript:ordersubmit(<% call OrderVariable("CCY")  %>) >Curr<br>貨幣</a></td>
   <td width="9%"><a href=javascript:ordersubmit(<% call OrderVariable("INSTRUMENT")  %>) >Instrument<br>股票號碼</a></td> 
   <td width="12%"><a href=javascript:ordersubmit(<% call OrderVariable("INSTRUMENTDESC")  %>) >Instrument Name<br>股票名稱</a></td> 
   <td width="9%"><a href=javascript:ordersubmit(<% call OrderVariable("QUANTITY")  %>) >Quantity<br>數量</a></td> 
   <td width="17%"><a href=javascript:ordersubmit(<% call OrderVariable("PRICE")  %>) >Price<br>價錢</a></td> 
   <td width="17%"><a href=javascript:ordersubmit(<% call OrderVariable("AMOUNT")  %>) >Amount<br>總值</a></td> 
   <td width="17%"><a href=javascript:ordersubmit(<% call OrderVariable("REMARK")  %>) >Remark<br>備註</a></td> 
</tr>
<tr bgcolor="#ADF3B6" align="center">
			<td><% call OrderImage("TRADEDATE")  %></td>
			<td><% call OrderImage("CLNT")  %></td>
			<td><% call OrderImage("CLNTNAME")  %></td>
			<td><% call OrderImage("VALUEDATE")  %></td>
			<td><% call OrderImage("TRADETYPE")  %></td>
			<td><% call OrderImage("MARKET")  %></td>
			<td><% call OrderImage("CCY")  %></td>
			<td><% call OrderImage("INSTRUMENT")  %></td>
			<td><% call OrderImage("INSTRUMENTDESC")  %></td>
			<td><% call OrderImage("QUANTITY")  %></td>
			<td><% call OrderImage("PRICE")  %></td>
			<td><% call OrderImage("AMOUNT")  %></td>
			<td><% call OrderImage("REMARK")  %></td>

</tr>
</thead>
<tbody>

		<%
			dim mystr
			do while (  Not rs1.EOF)
	
				
		%>
<tr bgcolor="#FFFFCC"> 
      <td> <%=rs1("TradeDate") %>　</td>
      <td><%=rs1("Clnt") %><span class="noprint"><img border=0 src='images/tel.gif' onClick="PopupClientContact('<%=rs1("clnt") %>')"></img></span></td>
      <td ><%=rs1("ClntName") %>　</td>
   <td><%=rs1("ValueDate") %>　</td>
   <td><% =rs1("TradeType") %>  </td>
   <td><%=rs1("Market") %>　</td>
   <td><%=rs1("Ccy") %>　</td>
   <td><%=rs1("Instrument") %>　</td> 
   <td><%=rs1("InstrumentDesc") %>　</td> 
   <td><%  if rs1("strquantity") <> "" then
   											mystr = replace(rs1("strquantity"), chr(13), "<br>")
   											response.write	left(mystr, len(mystr)-1) 
   										end if%>　</td> 
   <td><%  if rs1("strPrice") <> "" then
   											mystr = replace(rs1("strPrice"), chr(13), "<br>")
   											response.write	left(mystr, len(mystr)-1) 
   										end if%>　</td> 
   <td>&nbsp;<%=formatnumber(rs1("Amount"),2) %></td> 
   <td><%=rs1("Remark") %>　</td> 
</tr>
<%



				rs1.movenext
				
		loop
		
%>
</tbody>


</table>
<br>
<br>
<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

<tr bgcolor="#FFFFCC"> 
      <td width="166%" height="18" align="center">End of Statement</td>
</tr>

                </table>





<span class="noprint">

<%
'**********
' Start of page navigation 
'**********
%>
  <TABLE border=0 cellPadding=0 cellSpacing=0 height=100% width=99%>

 <tr> 
 <td align="right" height="28" class="NavaMenu" >
<%



response.write (iPageCurrent & " Pages " & iPageCount &"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp" )

'First button
%>
	<a href=javascript:dosubmit(1) style='cursor:hand'>First</a>

<%
' Prev button
If iPageCurrent > 1 Then
	%>
	<a href=javascript:dosubmit(<%= iPageCurrent-1 %>) style='cursor:hand'>Previous</a>
<% else %>
Previous
	<%
End If


'Next button
If iPageCurrent < iPageCount Then
	%>
	<a href=javascript:dosubmit(<%= iPageCurrent+1 %>) style='cursor:hand'>Next</a>
<% else %>
Next
	<%
End If
%>

<%
'Last button
%>

<a href=javascript:dosubmit(<%= iPageCount %>) style='cursor:hand'>Last</a>

</td></tr></table>


<%
'**********
' End of page navigation 
'**********
%>


<%
	end if 'record found if statement
end if   'having client number if statement
%>
</div>              



<%
'*****************************************************************
' End of report body
'*****************************************************************
%>
                
                
                
                </td>
                </tr>
              </table>
</div>


</span>

              </body>

              </html>
              
<%
'*****************************************************************
' Termination
'*****************************************************************


 Conn.Close
 Set Conn = Nothing


%>
<SCRIPT language=JavaScript>
<!--
function doConvert(){
window.open("ConvertTran.asp?Search_Amount_type=<%=Search_Amount_type%>&Search_Currency=<%=Search_Currency%>&Search_Transaction_Type=<%=Search_Transaction_Type%>&Search_Instrument=<%=Search_Instrument%>&Search_Market=<%=Search_Market%>&From_Day=<%=Search_From_Day%>&From_Month=<%=Search_From_Month%>&From_Year=<%=Search_From_Year%>&To_day=<%=Search_To_Day%>&To_Month=<%=Search_To_Month%>&To_Year=<%=Search_To_Year%>&Search_SharedGroupMember=<%=Search_SharedGroupMember%>&Search_SharedGroup=<%=Search_AEGroup%>"); 

}

//-->
</SCRIPT>