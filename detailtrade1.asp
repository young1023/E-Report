<% 
'*********************************************************************************
'NAME       : DetailTrade1.asp           
'DESCRIPTION: Instrument trading with all details info group by Client
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 090401 Gary Yeung   Prototype
'MODIFIED   : 090415 Roger Wong   Record and page control
'					 		090712 Roger Wong		Add Shared Group
'********************************************************************************

%>

<%
On Error resume Next
%>

<!--#include file="include/SessionHandler.inc.asp" -->
<%

Dim MSG_BUSY
MSG_BUSY = "The selection criteria are too broad for the system to process. Please use more specific selection criteria then resubmit."

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
itotalturnover = 0
itotalconsideration = 0
itotalbrokerage = 0
itotalNetAmount = 0
MTDTurnover = 0
MTDBrokerage = 0
MTDNetAmount = 0
YTDTurnover  = 0
YTDBrokerage = 0
YTDNetAmount = 0

strURL = Request.ServerVariables("URL") ' Retreive the URL of this page from Server Variables
%>





<%
if session("shell_power")="" then
  response.redirect "Default.asp"
end if


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
Dim Search_To_Day
Dim Search_To_Month
Dim Search_To_Year
Dim Search_Market
Dim Search_Instrument
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
Search_Instrument       = Request.form("Instrument")
Search_Order            = Request.form("Order")
Search_Direction        = Request.form("Direction")
Search_SharedSelection  = Request.form("ShareSelection")	
Search_SharedGroup      = Request.form("SharedGroup")
Search_SharedGroupMember= Request.form("SharedGroupMember")



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


'pass all selection to session 
session("GroupID")               =  Search_AEGroup	               
session("ClientFrom")            =  Search_ClientFrom              
session("ClientTo")              =  Search_ClientTo                
session("AEFrom")                =  Search_AEFrom                  
session("AETo")                  =  Search_AETo                    


set RsMarket = server.createobject("adodb.recordset")
RsMarket.open ("Exec Retrieve_AvailableMarket ") ,  StrCnn,3,1


set RsGroupID = server.createobject("adodb.recordset")
RsGroupID.open ("Exec Retrieve_AvailableGroupID ") ,  StrCnn,3,1


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
      <td width="20%">Period From:</td> 
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
     
      <td width="20%">Period To:</td> 
      <td width="27%">
      	     
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


Year_starting = Year(DateAdd("yyyy", -1, Now()))
year_ending = Year(Now())

for i=Year_starting to Year_ending
%>			         
			<option value="<%=i%>" <% if clng(i)=clng(Search_To_Year) then response.write "selected"%>><%=i%></option>

<% next %>

			</select>
			</td>
    
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
     <input name="Instrument" type=text value="<%= Search_Instrument %>" size="15">&nbsp;   
 	     
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
         response.write "Shared"
 			
 			Rs1.open ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '',  '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
			
	case "share3"

 			Rs1.open ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
			
	case else
		'normal
 
 			Rs1.open ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1
	'Response.write ("Exec retrieve_DetailTrade_GroupBy_Client_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") 
	
	end select	



		If Err.Number <> 0 then
			
			'SQL connection error handler
			response.write  "<table><tr><td class='RedClr'>" & MSG_BUSY & "<br></td></tr></table>"

		End If



  
  dim itotalCCYcount
	erase itotalturnover
	erase itotalconsideration		
	erase itotalBrokerage		
  itotalCCYcount=0
  

  if Rs1.EoF then

        response.write "No record found"
		
		If Err.Number <> 0 then
			
			'SQL connection error handler
			'response.write  "<table><tr><td class='RedClr'>The server is currently too busy to process your request right now. Please wait a moment and then try again. If the problem persists, please contact systems administrator.<br></td></tr></table>"
			'response.write Err.Number
		else
			'no record found
			response.write "No record found"
				
		End If

	else


    

      
 
%>    

</span>




   
<%
'**********
' Start of page navigation 
'**********
%> 
    

  




<table width="97%" border="0" class="normal"  cellspacing="1" cellpadding="4">
<tr bgcolor="#FFFFCC"> 
<td  width="20%"></td>
      <td align="center">詳細交易紀錄<br><u>Detail Trade Information</u></td> 
      <td align="right" width="20%"><span class="noprint">
							<%if PrintAllowed = 1 then %>  
							<a href="javascript:window.print()">Friendly Print</a><% 'end if %>
<% End If %>
<%if (session("shell_power") = 8) then %>    
<a href="javascript:window.doConvert()">Excel</a>
						<% end if %>      	
			</span></td>
</tr>
</table>
<br>
<table class="sortable"  width="99%" border="0" style="border-width: 0;FONT-SIZE: 11px;TEXT-ALIGN: Right;FONT-FAMILY: Verdana, 'MS Sans Serif', Arial" bgcolor="#808080" cellspacing="1" cellpadding="2">
<tr class="alignright" bgcolor="#ADF3B6" align="center">
   <td width="15%"><span style="cursor:hand">Client Code <br> 客戶編號</span></td>
   <td width="25%"><span style="cursor:hand">Client Name<br>客戶</span></td>
   <td width="15%"><span style="cursor:hand">Turnover (HKD)<br>交易總額</span></td>
   <td width="15%"><span style="cursor:hand">Brokerage (HKD)<br>客戶佣金</span></td> 
   <td width="15%"><span style="cursor:hand">Net Comm (HKD)<br>經紀佣金</sapn></td> 
   <td width="15%"><span style="cursor:hand">Net Amount (HKD)<br>總額</span></td> 
</tr>

		<%
			dim iPageCCYcount
			dim k
			dim iPageUpdate
			
			dim mystr
			do while (Not rs1.EOF)
				k=1
				
		%>
		
<tr class="alignright" bgcolor="#FFFFCC"> 
   <td width="15%"><a href="DetailTrade2.asp?PrintAllowed=<%=PrintAllowed%>&DisplayFirst=<%=Trim(rs1("ClientCode"))%>&ClientFrom=<%=Search_ClientFrom%>&ClientTo=<%=Search_ClientTo%>&AEFrom=<%=Search_AEFrom%>&AETo=<%=Search_AETo%>&Instrument=<%=Search_Instrument%>&Market=<%=Search_Market%>&FromDay=<%=Search_From_Day%>&FromMonth=<%=Search_From_Month%>&FromYear=<%=Search_From_Year%>&Today=<%=Search_To_Day%>&ToMonth=<%=Search_To_Month%>&ToYear=<%=Search_To_Year%>&Search_Order=ClientCode&Search_Direction=ASC&sid=<%=SessionID%>#DisplayFirst" target=_blank><%=rs1("ClientCode") %></a><span class="noprint"><img border=0 src='images/tel.gif' onClick="PopupClientContact('<%=rs1("ClientCode") %>')"></img></span></td>
   <td width="25%"><%=rs1("ClientName") %></td>
   <td width="15%"><%=formatnumber(rs1("totalturnover"),2)  %></td> 

<%
	

%>
   <td width="15%"><%=formatnumber(rs1("totalBrokerage"),2)%></td> 
   <td width="15%"><%=formatnumber(rs1("totalBrokerage"),2) %></td> 
   <td width="15%"><%=formatnumber(rs1("totalNetAmount"),2) %></td> 
</tr>

<%



   
   itotalturnover = itotalturnover + formatnumber(rs1("totalturnover"))

   itotalbrokerage = itotalbrokerage + formatnumber(rs1("totalBrokerage"))

   itotalconsideration = itotalconsideration + formatnumber(rs1("totalconsideration"))

   itotalNetAmount = itotalNetAmount + formatnumber(rs1("totalNetamount"))

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
   <td width="15%"><%=formatnumber(itotalbrokerage,2) %></td>
   <td width="15%"><%=formatnumber(itotalNetAmount,2) %></td>

</tr>

<tr bgcolor="#FFFFFF"> 
   <td colspan="6">&nbsp;</td>
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
 
 			Rs2.open ("Exec retrieve_DetailTrade_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '1', '100', 'ClientCode', 'ASC' ") ,  StrCnn,3,1
		 'Response.write  ("Exec retrieve_DetailTrade_MTD_HKD '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '1', '100', 'ClientCode', 'ASC' ")        
	
	end select	


    Set Rs8 = server.createobject("adodb.recordset")  
   
    Rs8.open ("Exec Retrieve_RebateAmount_HKD '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_AEFrom&"' ") ,  Conn,3,1
    'Response.write ("Exec Retrieve_RebateAmount_HKD '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_AEFrom&"' ")

 


    If Not Rs2.EoF Then

    MTDTurnover  = formatnumber(rs2("totalturnover"),2)
    MTDBrokerage = formatnumber(rs2("totalBrokerage"),2)
    
    MTDRebateAmount = Rs8("AECommFC") + Rs8("IntroducerRebateFC") + Rs8("ClientRebateFC") + Rs8("ResearchCreditFC")
    MTDNetAmount = rs2("totalNetAmount")

    MTDNetAmount = MTDNetAmount-MTDRebateAmount
    MTDNetAmount = formatnumber(rs2("totalNetAmount"),2)

    End If

%>
   
<tr bgcolor="#FFFFCC"> 
	 
   <td colspan="2" align="right">MTD Subtotal (HKD)</td>
   <td ><%=MTDTurnover%></td>
   <td ><%=formatnumber(rs2("totalBrokerage"),2)%></td>
   <td ><%=formatnumber(rs2("totalBrokerage"),2)%></td>
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

        
    Rs8.MoveNext  

If Not Rs3.EoF Then

    YTDTurnover  = formatnumber(rs3("totalturnover"),2)
    YTDBrokerage = formatnumber(rs3("totalBrokerage"),2)

    YTDRebateAmount = Rs8("AECommFC") + Rs8("IntroducerRebateFC") + Rs8("ClientRebateFC") + Rs8("ResearchCreditFC")
    YTDNetAmount = rs3("totalNetAmount")

    YTDNetAmount = YTDNetAmount-YTDRebateAmount
    YTDNetAmount = formatNumber(YTDNetAmount,2)
    
    End If

Rs3.Close
Set Rs3=Nothing

%>

<tr bgcolor="#FFFFCC"> 

   <td colspan="2" align="right">YTD Subtotal (HKD)<BR></td>
   <td ><%=YTDTurnover%></td>
   <td ><%=YTDBrokerage%></td>
   <td ><%=YTDBrokerage%></td>
   <td ><%=YTDNetAmount%></td>
</tr>

</table>
<br>
<br>


</div>



<%
	    
    end if 'record found if statement
end if   'having client number if statement
%>



<%
'*****************************************************************
' End of report body
'*****************************************************************
%>
                
              
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
 PrintAllowed = 0
%>
<SCRIPT language=JavaScript>
<!--
function doConvert(){
window.open("ConvertDetailTrade.asp?Search_Instrument=<%=Search_Instrument%>&Search_Market=<%=Search_Market%>&From_Day=<%=Search_From_Day%>&From_Month=<%=Search_From_Month%>&From_Year=<%=Search_From_Year%>&To_day=<%=Search_To_Day%>&To_Month=<%=Search_To_Month%>&To_Year=<%=Search_To_Year%>"); 

}

//-->
</SCRIPT>