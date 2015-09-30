<% 
'*********************************************************************************
'NAME       : ClientSummary.asp           
'DESCRIPTION: Client Summary
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 090401 Gary Yeung   Prototype
'MODIFIED   : 090420 Roger Wong   Record and page control
'							090712 Roger Wong		Add Shared Group
'********************************************************************************

' Section Order	Section Code									        	
' --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'      1		         CN		Contract Note / Trade Summary 		(Cash & Custodian)        	
'      4		         TM		Contract Note / Trade Summary 		(Margin)		           	
'      2		         CC		Sub-total of Each Trade					           	
'      2		         CL		Total Amount of All Trade				           	
'      4		         CM		Cash Movement							
'      6		         SM		Securities Movement						
'      4		         BM		Cash Opening Balance						
'      5		         CB		Cash Closing Balance						
'      7		         UT		UnSettled and Pending Trade					
'      8		         TL		Total Value of UnSettled and Pending Trade					
'      9		         MP		Stock Portfolio 				(Margin)			
'      10		         ML		Total Market Value of Stock Portfolio	(Margin)			
'      11		         SP		Stock Portfolio 				(Cash & Custodian)	
'      12		         SL		Total Market Value of Stock Portfolio	(Cash & Custodian)	
'      13 		       MS		Summary 				(Margin)				   	
'      14		         NS		Summary 				(Cash & Custodian)		   	


Title = "Client Summary"

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


'**************
'Initialisation
'**************


Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adCmdText = &H0001

Const PAGE_SIZE = 5  ' The size of our pages.

Dim iPageCurrent ' The page we're currently on
Dim iPageCount   ' Number of pages of records
Dim iRecordCount ' Count of the records returned
Dim I            ' Standard looping variable


Dim cnnSearch  ' ADO connection
Dim rstSearch  ' ADO recordset




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
    	}
    }  -->
    
    </style>

<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
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
		if (datevalidate(document.fm1.SDay.value, document.fm1.SMonth.value -1, document.fm1.SYear.value) == false) 
		{
			return false;

		}
		
		//User must enter Client From field
		if (document.fm1.ClientFrom.value == ""){
  			alert("Please enter client number");
        document.fm1.ClientFrom.focus();
        return false;
		}
		
		
		if  (isNaN(document.fm1.LedgerBalance.value) == true){
  			alert("Min. ledger balance should in numeric format");
        document.fm1.LedgerBalance.focus();
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


//-->
</SCRIPT>

<script language="JavaScript">
function disableCtrlKeyCombination(e)
{
        //list all CTRL + key combinations you want to disable
        var forbiddenKeya = 'a';
        var forbiddenKeyc = 'c';
        var forbiddenKeyx = 'x';


        var key;
        var isCtrl;

        if(window.event)
        {
                key = window.event.keyCode;     //IE
                if(window.event.ctrlKey)
                        isCtrl = true;
                else
                        isCtrl = false;
        }
        else
        {
                key = e.which;     //firefox
                if(e.ctrlKey)
                        isCtrl = true;
                else
                        isCtrl = false;
        }

        //if ctrl is pressed check if other key is in forbidenKeys array
        if(isCtrl)
        {
            
                {
                        //case-insensitive comparation
                        if(forbiddenKeya.toLowerCase() == String.fromCharCode(key).toLowerCase())
                        {
                                return false;
                        }
                        if(forbiddenKeyc.toLowerCase() == String.fromCharCode(key).toLowerCase())
                        {
                                return false;
                        }

						if(forbiddenKeyx.toLowerCase() == String.fromCharCode(key).toLowerCase())
                        {
                                return false;
                        }

                }
        }
        return true;
}
</script>

<script language="JavaScript">
<!--
// disable right click
var message="Sorry, The right click function is disable."; // Message for the alert box

function click(e) {
if (document.all) {
if (event.button == 2) {
alert(message);
return false;
}
}
if (document.layers) {
if (e.which == 3) {
alert(message);
return false;
}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;
// --> 
</script>

</head>
<body leftmargin="0" topmargin="0" OnLoad="document.fm1.submitted.value=0;document.fm1.ClientFrom.focus();"  onkeypress="return disableCtrlKeyCombination(event);" onkeydown="return disableCtrlKeyCombination(event);"  >




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
Dim Search_Statement
Dim Search_Daily_Day
Dim Search_Daily_Month
Dim Search_Daily_Year
Dim Search_Market
Dim Search_Instrument
Dim Search_LedgerBalance
Dim Search_LedgerBalanceType
Dim Search_SharedSelection 
Dim Search_SharedGroup
Dim Search_SharedGroupMember

Search_AEGroup	    = Request.form("GroupID")
Search_ClientFrom   = Request.form("ClientFrom")
Search_ClientTo     = Request.form("ClientTo")
Search_AEFrom       = Request.form("AEFrom")
Search_AETo         = Request.form("AETo")
Search_Daily_Day    = Request.form("SDay")
Search_Daily_Month  = Request.form("SMonth")
Search_Daily_Year   = Request.form("SYear")
Search_LedgerBalanceType   = Request.form("LedgerBalanceType")
Search_LedgerBalance   = Request.form("LedgerBalance")

Search_IncludeMarginAccount   = Request.form("IncludeMarginAccount")
Search_Market           = Request.form("Market")
Search_Instrument       = Request.form("Instrument")
Search_SharedSelection      =  Request.form("ShareSelection")	
Search_SharedGroup  = Request.form("SharedGroup")
Search_SharedGroupMember  = Request.form("SharedGroupMember")



set RsMarket = server.createobject("adodb.recordset")
RsMarket.open ("Exec Retrieve_AvailableMarket ") ,  Conn,3,1


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
If Request.form("page") = ""  Then
	iPageCurrent = 1

	Search_AEGroup	    = session("GroupID")
	Search_ClientFrom   = session("ClientFrom")
	Search_ClientTo     = session("ClientTo")
	Search_AEFrom       = session("AEFrom")
	Search_AETo         = session("AETo")
	
	Search_Daily_Day = day(Session("DBLastModifiedDateValue"))
	Search_Daily_Month = month(Session("DBLastModifiedDateValue"))
	Search_Daily_Year = year(Session("DBLastModifiedDateValue"))
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
                                
                                
                                
                                
set RsGroupID = server.createobject("adodb.recordset")
RsGroupID.open ("Exec Retrieve_AvailableGroupID ") ,  Conn,3,1
                                
                                
%>                              
                                
<%                              
'*****************************************************************
' Start of form
'*****************************************************************
%>
         


<form name="fm1" method="post" action="<%= strURL %>">
	
	
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
				<td colspan="2" align=right><font color="red">*</font> Denotes a mandatory field
				</td>
			</tr>
			<% End If %>
			<tr>
					<td width="20%">Client Number: (From)<font color="red">*</font> </td> 
					<td width="30%">
					<input name="ClientFrom" type=text value="<%= Search_ClientFrom %>" size="15">
                    <img align="top" style="cursor:pointer" onClick="PopupWindow()" src="images/search.gif"> </td>
			<td width="20%">Client Number: (To)</td> 
			<td width="22%">
			<input name="ClientTo" type=text value="<% If Search_ClientFrom <> Search_ClientTo Then Response.Write Search_ClientTo End If%>" size="15"></td>
		</tr>
<%
' Show if branch manager or above
if session("shell_power")>=3 then %>
			  
		<tr> 
			<td width="20%">AE Code: (From)</td> 
			<td width="30%">
			<input name="AEFrom" type=text value="<%= Search_AEFrom %>" size="15">
            <img align="top" style="cursor:pointer" onClick="PopupSearchAE()" src="images/search.gif"></td>

			<td width="20%">AE Code: (To)</td> 
			<td>
			<input name="AETo" type=text value="<% If Search_AEFrom <> Search_AETo Then Response.Write Search_AETo End If%>" size="15"></td>
			</tr>   

<% end if %>  





		<tr> 
			<td width="20%">Date of Report:</td> 
			<td width="30%">
			 
			<select name="SDay" class="common">
			<option value="1" <% if Search_Daily_Day=1 then response.write "selected"%>>1</option>
			<option value="2" <% if Search_Daily_Day=2 then response.write "selected"%>>2</option>
			<option value="3" <% if Search_Daily_Day=3 then response.write "selected"%>>3</option>
			<option value="4" <% if Search_Daily_Day=4 then response.write "selected"%>>4</option>
			<option value="5" <% if Search_Daily_Day=5 then response.write "selected"%>>5</option>
			<option value="6" <% if Search_Daily_Day=6 then response.write "selected"%>>6</option>
			<option value="7" <% if Search_Daily_Day=7 then response.write "selected"%>>7</option>
			<option value="8" <% if Search_Daily_Day=8 then response.write "selected"%>>8</option>
			<option value="9" <% if Search_Daily_Day=9 then response.write "selected"%>>9</option>
			<option value="10" <% if Search_Daily_Day=10 then response.write "selected"%>>10</option>
			<option value="11" <% if Search_Daily_Day=11 then response.write "selected"%>>11</option>
			<option value="12" <% if Search_Daily_Day=12 then response.write "selected"%>>12</option>
			<option value="13" <% if Search_Daily_Day=13 then response.write "selected"%>>13</option>
			<option value="14" <% if Search_Daily_Day=14 then response.write "selected"%>>14</option>
			<option value="15" <% if Search_Daily_Day=15 then response.write "selected"%>>15</option>
			<option value="16" <% if Search_Daily_Day=16 then response.write "selected"%>>16</option>
			<option value="17" <% if Search_Daily_Day=17 then response.write "selected"%>>17</option>
			<option value="18" <% if Search_Daily_Day=18 then response.write "selected"%>>18</option>
			<option value="19" <% if Search_Daily_Day=19 then response.write "selected"%>>19</option>
			<option value="20" <% if Search_Daily_Day=20 then response.write "selected"%>>20</option>
			<option value="21" <% if Search_Daily_Day=21 then response.write "selected"%>>21</option>
			<option value="22" <% if Search_Daily_Day=22 then response.write "selected"%>>22</option>
			<option value="23" <% if Search_Daily_Day=23 then response.write "selected"%>>23</option>
			<option value="24" <% if Search_Daily_Day=24 then response.write "selected"%>>24</option>
			<option value="25" <% if Search_Daily_Day=25 then response.write "selected"%>>25</option>
			<option value="26" <% if Search_Daily_Day=26 then response.write "selected"%>>26</option>
			<option value="27" <% if Search_Daily_Day=27 then response.write "selected"%>>27</option>
			<option value="28" <% if Search_Daily_Day=28 then response.write "selected"%>>28</option>
			<option value="29" <% if Search_Daily_Day=29 then response.write "selected"%>>29</option>
			<option value="30" <% if Search_Daily_Day=30 then response.write "selected"%>>30</option>
			<option value="31" <% if Search_Daily_Day=31 then response.write "selected"%>>31</option>
		
			
			</select>


			<select name="SMonth" class="common">            	
					<option value="1" <% if Search_Daily_Month=1 then response.write "selected"%>>Jan</option>
					<option value="2" <% if Search_Daily_Month=2 then response.write "selected"%>>Feb</option>
					<option value="3" <% if Search_Daily_Month=3 then response.write "selected"%>>Mar</option>
					<option value="4" <% if Search_Daily_Month=4 then response.write "selected"%>>Apr</option>
					<option value="5" <% if Search_Daily_Month=5 then response.write "selected"%>>May</option>
					<option value="6" <% if Search_Daily_Month=6 then response.write "selected"%>>Jun</option>
					<option value="7" <% if Search_Daily_Month=7 then response.write "selected"%>>Jul</option>
					<option value="8" <% if Search_Daily_Month=8 then response.write "selected"%>>Aug</option>
					<option value="9" <% if Search_Daily_Month=9 then response.write "selected"%>>Sep</option>
					<option value="10" <% if Search_Daily_Month=10 then response.write "selected"%>>Oct</option>
					<option value="11" <% if Search_Daily_Month=11 then response.write "selected"%>>Nov</option>
					<option value="12" <% if Search_Daily_Month=12 then response.write "selected"%>>Dec</option>
			</select>


			<select name="SYear" class="common">   
<% 
Dim Year_starting
Dim Year_ending

Year_starting = Year(DateAdd("yyyy", -9, Now()))
year_ending = Year(Now())

for i=Year_starting to Year_ending
%>			         
			<option value="<%=i%>" <% if clng(i)=clng(Search_Daily_Year) then response.write "selected"%>><%=i%></option>

<% next %>

			</select> 
			</td>
	<td width="20%">Market:</td> 
	<td width="29%">
	 
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
	<td width="21%">
	</td>
</tr>
    

    
 <tr> 
      <td width="20%">Instrument:</td> 
      <td width="30%">
      	     
<input name="Instrument" type=text value="<%= Search_Instrument %>" size="15">&nbsp;   
 
			<td width="20%">Ledger Balance Type:</td> 
			<td>
			     
			<select size="1" name="LedgerBalanceType" class="common">
				<option value="ALL" <% if Search_LedgerBalanceType="ALL" then response.write "selected"%>>All</option>
				<option value="DEBIT" <% if Search_LedgerBalanceType="DEBIT" then response.write "selected"%>>Debit  </option>
				<option value="CREDIT" <% if Search_LedgerBalanceType="CREDIT" then response.write "selected"%>>Credit  </option>
			</select></td>
		</tr>

 <tr> 
      <td width="20%">Min. Ledger Balance:</td> 
      <td width="30%">
      	     
<input name="LedgerBalance" type=text value="<%= Search_LedgerBalance %>" size="15">&nbsp;   

			<td width="20%">Include Margin A/C</td> 
			<td>
			   
			<select size="1" name="IncludeMarginAccount" class="common">
				<option value="ALL" <% if Search_IncludeMarginAccount="YES" then response.write "selected"%>>All</option>
				<option value="MRGN" <% if Search_IncludeMarginAccount="MRGN" then response.write "selected"%>>Margin Only  </option>
			<option value="NMRGN" <% if Search_IncludeMarginAccount="NMRGN" then response.write "selected"%>>Non Margin Only  </option>
			</select>   
		</tr>
		
<%if session("SharedGroup") > 0 then 


'List Shared Group member
set RsSharedGroupMember = server.createobject("adodb.recordset")
RsSharedGroupMember.open ("Exec List_SharedGroupMember '"&Session("id")&"', '"&Session("shell_power")&"' ") ,  conn,3,1


'List shared group 
set RsSharedGroup = server.createobject("adodb.recordset")
RsSharedGroup.open ("Exec List_SharedGroup '"&Session("id")&"' ") ,  conn,3,1





%>		  
		<tr> 
			<td colspan="4">&nbsp;<input type="radio" name="ShareSelection" value="share1" onClick=""  <%if Search_SharedSelection = "share1" then response.write "checked" end if %>  > Viewing AE only<span class="noprint"> 
			| <input type="radio" name="ShareSelection" value="share2" onClick=""  <%if Search_SharedSelection = "share2" then response.write "checked" end if %>   > Particular AE in the Sales Team

			<select name="SharedGroupMember" class="common">
						<%
								do while (  Not RsSharedGroupMember.EOF)
						%>
								<option value="<%=RsSharedGroupMember("loginname")%>" <% if Search_SharedGroupMember=RsSharedGroupMember("loginname") then response.write "selected" end if %> > <%=RsSharedGroupMember("loginname")%></option>
						<%
								RsSharedGroupMember.movenext
								Loop
						%>
			</select>&nbsp; | <input type="radio" name="ShareSelection" value="share3" onClick=""  <%if Search_SharedSelection = "share3" then response.write "checked" end if %>   > All AEs in the Sales Team
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

			</span>

			</td> 
		</tr>   

<% end if %>  		
		
		<tr> 
			<td></td>
			<td colspan="3">
			<input type=hidden   value="<%=iPageCurrent%>"   name="page"> 
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

</span>

<%
'*****************************************************************
' Start of report body
'*****************************************************************
%>


<%


If Request("submitted") = 0 Then

'**********
' If no argument
'**********

'do nothing
  

Else


       ' Direct from Margin Summary
       If Request("submitted") = 2 Then

          Search_ClientFrom   = Request("ClientFrom")
 
          Search_ClientTo     = Request("ClientFrom")
       
          Search_Daily_Day = day(Session("DBLastModifiedDateValue"))

          Search_Daily_Month =  Month(Session("DBLastModifiedDateValue"))

          Search_Daily_Year = year(Session("DBLastModifiedDateValue"))

          Search_Market = "ALL"

          Search_IncludeMarginAccount ="ALL"

          Search_LedgerBalanceType = "ALL"

          Search_ClientFrom   = Request("ClientFrom")

       End if


'**********
' If passing arguments
'**********
	
	'dim Rs1 as adodb.recordset 
	
	'StrCnn.open myDSN 
	
	set Rs1 = server.createobject("adodb.recordset")
	
	'rs1.CursorLocation=3
	
Search_Statement ="Daily"


 'Rs return 2 value
 '1) Total number of matched client
 '2) all records for targeted client
 'Rs1.open ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"','"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&iPageCurrent&"', '1' ") ,  conn,3,1


	

	Select Case  Search_SharedSelection
	case "share2"
		'shared group member
			'response.write  ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_IncludeMarginAccount&"', '"&iPageCurrent&"', '1' ") 
			Rs1.open ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_LedgerBalance&"','"&Search_IncludeMarginAccount&"', '"&iPageCurrent&"', '1' ") ,  conn,3,1
	case "share3"
			'response.write  ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_IncludeMarginAccount&"', '"&iPageCurrent&"', '1' ")

			Rs1.open ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_LedgerBalance&"','"&Search_IncludeMarginAccount&"', '"&iPageCurrent&"', '1' ") ,  conn,3,1

	case else
		'normal
			'response.write  ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"','"&Search_LedgerBalance&"', '"&Search_IncludeMarginAccount&"', '"&iPageCurrent&"', '1' ")
			Rs1.open ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"','"&Search_LedgerBalance&"', '"&Search_IncludeMarginAccount&"', '"&iPageCurrent&"', '1' ") ,  conn,3,1
		
	end select	
		

	'assign total number of pages
	iPageCount = rs1(0)


  if iPageCount <= 0 then

		If Err.Number <> 0 then
			
			'SQL connection error handler
			response.write  "<table><tr><td class='RedClr'>" & MSG_BUSY & "<br></td></tr></table>"
			
		else
			'no record found
			response.write ("No record found")
				
		End If
	else
		'record found
		
		'move to next recordset
  	Set rs1 = rs1.NextRecordset() 
 
%>


<div id="reportbody1" >

<script type="text/javascript">
var somediv=document.getElementById("reportbody1")
disableSelection(somediv) //disable text selection within DIV with id="mydiv"
</script>	


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

<% if (session("shell_power") = 1 or session("shell_power") = 5) then %>  
		<span class="noprint">
<% end if %>


<table width="99%" border="0" class="normal"  cellspacing="1" cellpadding="2">



<tr bgcolor="#FFFFCC"> 
<td  width="20%">　</td>
      <td align="center">客戶總結<br><u>Client Summary</u></td> 
      <td align="right" width="20%">
						<% If PrintAllowed = 1 then %> 
							<a href="javascript:window.print()">Friendly Print</a>
						<% end if %>      	
			</td>
</tr>
</table>

<br>
<%'.....................%>


<% 



''''''''''''''''''''''''''
' Header (CH)
''''''''''''''''''''''''''

Dim AccountType

if (  Not rs1.EOF)  then
			if rs1("sectioncode") = "CH" then 
						
						'Accounttype: If or not Margin client 
						AccountType = rs1("AcctType")
						%>
			
						<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
								<tr bgcolor="#ADF3B6"> 
											<td width="30%" height="33">Client Name<br>客戶名稱</td>
											<td width="70%" height="33">Client Address<br>客戶聯絡地址</td> 
											</tr>
											<tr bgcolor="#FFFFCC"> 
											<td width="30%"><%=rs1("ClntName") %></td>
											<td width="70%"><%=rs1("Address") %></td>
								</tr>
						</table>
			
						<br>
						
						<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
						
								<tr bgcolor="#ADF3B6"> 
											<td width="14%" height="33">Print Date<br>列印日期</td>
											<td width="16%" height="33">Account No.<br>客戶編號</td>
											<td width="12%" height="33">AE Code<br>經紀編號</td>
											<td width="26%" height="33">Debit Interest Rate (HKD)<br>借貸利率</td>
											
												<% if AccountType = "MRGN" then
														response.write "<td width='10%' height='33'>Margin Limit<br>保證金限額</td>"
													 end if
												%>
											
											
													<% if AccountType = "MRGN" then
														response.write "<td width='17%' height='33'>Margin Authority Expiry Date<br>信貸限額到期日</td> "
													 end if
												%>
												
								</tr>
						
								
								<tr bgcolor="#FFFFCC"> 
											<td width="14%"><%=rs1("StatementDate") %></td>
											<td width="16%"><img border=0 src='images/tel.gif' onClick="PopupClientContact('<%=rs1("clnt") %>')"></img><%=rs1("clnt") %></td>
											<td width="12%"><%=rs1("AE") %></td>
											<td width="26%"><%= rs1("DebitInterestRate") %>%</td>
											
												<% if AccountType = "MRGN" then %>
													 	<td width='10%'> <%=formatnumber(rs1("MarginLimit")) %>
			    									</td>
														<td width='17%'>

														<%=rs1("ExpiryDate") %>
													  </td> 
												<% end if	%>

								</tr>
								
					</table>
					<%
					rs1.movenext
			end if
end if

%>
		<br>




<%
'''''''''
' Loop for next records
'''''''''

do while (Not rs1.EOF)
			
			Select Case rs1("sectioncode")
					'case "IN" exit do
					'case "CM" exit do
					'case "SM" exit do
					case "SP" exit do
					case "CB" exit do
					'case "MG" exit do
					'case "CN" exit do
			
			end select
			rs1.movenext
loop

'''''''''

%>




<br/>


<% 

	''''''''''''''''''''''''''
	'  Portfolio Statement (SP)
	''''''''''''''''''''''''''

dim TotalPortfolioValue
dim TotalMarginValue

TotalPortfolioValue=0
TotalMarginValue=0


if (  Not rs1.EOF)  then








	












			if rs1("sectioncode") = "SP" then %>
			
						<br>
						<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
							
									<tr bgcolor="#FFFFCC"> 
												<td width="95%" colspan="13">投資組合記錄 Portfolio Statement </td>
									</tr>
									
									<tr bgcolor="#ADF3B6">
												<td width="10%">Market<br>市場</td>
												<td width="10%">Stock Code<br>股票代號</td>
												<td width="12%">Stock Name<br>股票名稱</td>
												<td width="10%">Holding B/F<br>承上</td>
												<td width="15%">Movement<br>In/Out<br>出/入</td>
												<td width="10%">Holding C/F<br>轉下</td>
												<td width="10%">Closing Price<br>收市價</td>
												<td width="10%">CCY<br>貨幣</td> 
												<td width="17%">Value<br>市場價值</td> 
												<td width="17%">#Exch Rate<br>匯率</td> 
												<td width="17%">Value (HKD)<br>市場價值(港元)</td> 
																															
													<%if AccountType = "MRGN" then %>
															
														<td width="17%"> %</td>
														<td width="17%">Margin Value<br>按倉價值(港元)</td> 
													<%	end if %>
									</tr>
									
									<% 
									do while (  Not rs1.EOF)
									
												if rs1("sectioncode") = "SP" then
												%>

															<tr bgcolor="#FFFFCC"> 
																		<td width="10%"><%= rs1("Market") %></td>
																		<td width="10%"><%= rs1("Instrument") %></td>
																		<td width="12%"><% response.write rs1("InstrumentDesc") & "<br>" & rs1("InstrumentCDesc") %>　</td>
																		<td width="10%">
																			<%
																			
																			response.write formatnumber(rs1("OpeningBalance"),0,-2,-1) 
																			
																				if (clng(rs1("orfee1")) <> 0  ) then
																					response.write "<br>*" & formatnumber(clng(rs1("orfee1")),0,-2,-1)
																				end if																			
																			%>
																			</td>
																		<td width="10%">
																			<%
																				response.write formatnumber(rs1("Netbalance"),0,-2,-1) 

																			%>　
																		</td>
																		<td width="10%">
																			<%
																				response.write formatnumber(rs1("EndingBalance"),0,-2,-1) 
																				'stock on hold positive
																				'response.write clng(rs1("orfee1"))
																				if (clng(rs1("orfee1")) <> 0  ) then
																					response.write "<br>*" & formatnumber(clng(rs1("orfee2")),0,-2,-1)
																				end if
																				
														
																			%>　
																		</td>

																		<td width="10%"><%= rs1("Price") %>　</td>
																		<td width="10%"><%= rs1("CCY") %>　</td> 
																		<td width="17%"><%= formatnumber(cDbl(rs1("StockPortValue")) / cDbl(rs1("MarginLimit")),2,-2,-1) %>　</td> 
																		<td width="17%"><%= formatnumber(rs1("MarginLimit"),4) %>　</td> 
																		<td width="17%"><%= formatnumber( int(cDbl(rs1("StockPortValue"))*100)/100  ,2,-2,-1) %>　</td> 


																		<%if AccountType = "MRGN" then %>
																				
																			<td width="17%"><%=rs1("MarginPercent")%></td>
																			<td width="17%"><%=formatnumber(TotalMarginValue,2,-2,-1)%></td> 
																		<%	end if %>
																		
																		<%
																				TotalPortfolioValue  = TotalPortfolioValue + cDbl(rs1("StockPortValue"))
																				TotalMarginValue = TotalMarginValue + cDbl(rs1("MarginValue"))
																		%>
															</tr>
															

															<%
												end if	
												rs1.movenext
												if not rs1.eof then
														if rs1("sectioncode") <> "SP" then 
																exit do
														end if
												end if
									loop
							
									
									%>
									<tr bgcolor="#FFFFCC">
												<td colspan="<%if AccountType = "MRGN" then response.write "10" else response.write "8"%>" align="right">
													<b>股份組合市值 Portfolio Value: </b></td>
												<td><b><%=formatnumber(TotalPortfolioValue) %></b></td>
												<td></td>
												<td><b>
													<%
													if AccountType = "MRGN" then 
														response.write formatnumber(TotalMarginValue) 
													end if %>
													</b>
												</td>
									</tr>
									
						</table>
			<%
			
			end if 
end if

%>

		</table>
		<br>

				<% do while (  Not rs1.EOF)
		
				if rs1("sectioncode") = "CB" then
					exit do
				end if
				rs1.movenext
			loop
			
			if not rs1.eof then
				if rs1("sectioncode") = "CB" then 
		%>



<br>
<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				

<tr bgcolor="#ADF3B6">
   <td width="16%" rowspan="2">ccy<br>貨幣</td>
   <td width="16%" rowspan="2">Ledger balance<br>帳面結餘</td>
   <td width="30%" colspan="4" align="center">Available Balance<br>可用結餘</td>
   <td width="24%" colspan="4">Settlement Amount<br>結算總額</td> 
      <td width="14%" rowspan="2">Int Accrued<br>應計利息</td>
</tr>

<tr bgcolor="#ADF3B6">
   <td width="7%">T Day<br>交易當天</td>
   <td width="7%">T + 1<br>交易次天</td>
   <td width="7%">T + 2<br>交易之後第二天</td>
   <td width="9%">T + 3<br>交易之後第三天</td> 
   <td width="7%">T Day<br>交易當天</td>
   <td width="4%">T + 1<br>交易次天</td>
   <td width="3%">T + 2<br>交易之後第二天</td>
   <td width="7%">T + 3<br>交易之後第三天</td> 
</tr>

		<% do while (  Not rs1.EOF)
			
				if rs1("sectioncode") = "CB" then
		%>
		

<tr bgcolor="#FFFFCC"> 
   <td width="16%"><%= rs1("ccy") %></td>
   <td width="7%"><%= formatnumber(rs1("EndingBalance"),2,-2,-1) %>　</td>
   <td width="7%"><%= formatnumber(rs1("ORFee1"),2,-2,-1) %>　</td>
   <td width="7%"><%= formatnumber(rs1("ORFee2"),2,-2,-1) %>　</td>
   <td width="9%"><%= formatnumber(rs1("orfee3"),2,-2,-1) %>　</td> 
   <td width="7%"><%= formatnumber(rs1("orfee4"),2,-2,-1) %>　</td> 
   <td width="4%"><%= formatnumber(rs1("orfee7"),2,-2,-1)  %> 　</td> 
   <td width="3%"><%= formatnumber(rs1("orfee8"),2,-2,-1) %>　</td> 
   <td width="7%"><%= formatnumber(rs1("orfee9"),2,-2,-1) %>　</td> 
      <td width="14%"><%= formatnumber(rs1("orfee10"),2,-2,-1) %></td>
      <td width="14%"><%= formatnumber(rs1("MTDDebitInterest"),2,-2,-1) %></td>

</tr>
<%
					end if	
				rs1.movenext
				
			if not rs1.eof then
				Select Case rs1("sectioncode")
						case "IN" exit do
						case "CM" exit do
						case "SM" exit do
						case "SP" exit do
						'case "CB" exit do
						case "MG" exit do
						case "CN" exit do
				
				end select
			end if

		loop
		

	end if 
	end if

%>

</table>
<br>



<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

<tr bgcolor="#FFFFCC"> 
      <td width="166%" height="18" align="center">End of Statement</td>


<% if (session("shell_power") = 1 or session("shell_power") = 5) then %>  
		</span>
<% end if %>


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



<%
'*****************************************************************
' End of report body
'*****************************************************************
%>

</td></tr></table>


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
 Conn.Close
 Set Conn = Nothing
%>