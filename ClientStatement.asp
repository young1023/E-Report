
<% 
'*********************************************************************************
'NAME       : ClientStatement.asp           
'DESCRIPTION: Client Statement (Monthly & Daily)
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 090401 Gary Yeung   Prototype
'MODIFIED   : 090403 Roger Wong   Record and page control
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
%>

<%
Title = "Client Statement"

' **************
' Authorisation
' **************
if session("shell_power")="" then
  response.redirect "logout.asp?r=-1"
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" media="screen"/>
<!-- Include Print CSS -->
<link rel="stylesheet" type="text/css" media="print" href="include/print.css" />


<SCRIPT language=JavaScript>
<!--

function PopupClientContact(clientnumber) {
	 
		var str='ListClientContact.asp?sid=<%=SessionID%>&clientnumber=' + clientnumber
		
		newwindow=window.open(str , "myWindow", 
									"status = 1, height = 300, width = 600, resizable = 1'"  )
		 if (window.focus) {
           newwindow.focus();
       }
 			
}

function validateUserEntry(){

		
		//User must enter Client From field
		if (document.fm1.ClientFrom.value == ""){
  			alert("Please enter client number");
        document.fm1.ClientFrom.focus();
        return false;
		}
		
		if  (isNaN(document.fm1.NetValue.value) == true ){
  			alert("Net value should in numeric format");
        document.fm1.NetValue.focus();
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

function dosubmit(what){
			if (validateUserEntry() == false)
			{
				return false

			}  
			// Date validation for daily statement
			if (datevalidate(document.fm1.SDay.value, document.fm1.SMonth.value -1, document.fm1.SYear.value) == false) 
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
<body leftmargin="0" topmargin="0" OnLoad="document.fm1.submitted.value=0;document.fm1.ClientFrom.focus();" onkeypress="return disableCtrlKeyCombination(event);" onkeydown="return disableCtrlKeyCombination(event);" >




<%
'response.write "<div><table><tr><td>___" & session("shell_power") & "___</td></tr></table></div>"



if session("shell_power")="" then
  response.redirect "logout.asp?r=-1"
end if
%>

         

<%




'**************
'Argument handler
'**************

Dim Search_GroupID
Dim Search_ClientFrom
Dim Search_ClientTo
Dim Search_AEFrom
Dim Search_AETo
Dim Search_AEGroup
Dim Search_Statement
Dim Search_Monthly_Month
Dim Search_Monthly_Year
Dim Search_Daily_Day
Dim Search_Daily_Month
Dim Search_Daily_Year
Dim Search_SharedSelection 
Dim Search_SharedGroup
Dim Search_SharedGroupMember
Dim Search_NetValue

Search_AEGroup	    = Request.form("GroupID")
Search_ClientFrom   = Request.form("ClientFrom")
Search_ClientTo     = Request.form("ClientTo")
Search_AEFrom       = Request.form("AEFrom")
Search_AETo         = Request.form("AETo")
'Search_AEGroup 			= Request.form("AEGroup") 
Search_Statement    = Request.form("Statement")
Search_Monthly_Month= Request.form("SMMonth")
Search_Monthly_Year = Request.form("SMYear")
Search_Daily_Day    = Request.form("SDay")
Search_Daily_Month  = Request.form("SMonth")
Search_Daily_Year   = Request.form("SYear")
Search_NetValue = Request.form("NetValue")	

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

'Shared group handler
'If Session("SharedGroup") > 0 then
'		Search_AEFrom = Request.form("AEFrom")
'		Search_AETo = Request.form("AETo")
'		
'end if


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
	'Search_GroupID ="ALL"
	Search_Statement="Daily"
	
	Search_AEGroup	    = session("GroupID")
	Search_ClientFrom   = session("ClientFrom")
	Search_ClientTo     = session("ClientTo")
	Search_AEFrom       = session("AEFrom")
	Search_AETo         = session("AETo")
	
	Search_Daily_Day = day(Session("DBLastModifiedDateValue"))
	Search_Daily_Month = month(Session("DBLastModifiedDateValue"))
	Search_Daily_Year = year(Session("DBLastModifiedDateValue"))
	Search_Monthly_Month = month(Session("DBLastModifiedDateValue"))
	Search_Monthly_Year = year(Session("DBLastModifiedDateValue"))
	Search_SharedSelection = "share1"



Else
	iPageCurrent = Clng(Request.form("page"))
End If

session("GroupID")               =  Search_AEGroup	               
session("ClientFrom")            =  Search_ClientFrom              
session("ClientTo")              =  Search_ClientTo                
session("AEFrom")                =  Search_AEFrom                  
session("AETo")                  =  Search_AETo                    


set RsGroupID = server.createobject("adodb.recordset")
RsGroupID.open ("Exec Retrieve_AvailableGroupID ") ,  StrCnn,3,1




%>

 
<span class="noprint">

<!-- #include file ="include/Master.inc.asp" -->

<div id="Content">


<%
'*****************************************************************
' Start of form
'*****************************************************************
%>



<form name="fm1" method="post" action="<%= strURL %>?sid=<%=SessionID%>">
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
                    <img align="top"  onClick="PopupWindow()" src="images/search.gif" style="cursor:pointer"></td>
	<td width="20%">Client Number: (To)</td> 
	<td width="30%">
	 
	<input name="ClientTo" type=text value="<% If Search_ClientFrom <> Search_ClientTo Then Response.Write Search_ClientTo End If%>" size="15"></td>
	</tr>
	
	
	<% if session("shell_power") >=3 then %>
	
	<tr>
	<td width="20%" >AE Code: (From)</td> 
	<td width="30%" >
	 
	<input name="AEFrom" type=text value="<%= Search_AEFrom %>" size="15">
    <img align="top" style="cursor:pointer" onClick="PopupSearchAE()" src="images/search.gif"></td>
	<td width="20%" >AE Code: (To)</td> 
	<td width="30%" >
	 
	<input name="AETo" type=text value="<% If Search_AEFrom <> Search_AETo Then Response.Write Search_AETo End If%>" size="15"></td>
	</tr>
	<% End If %>
		

	 <tr> 
      <td width="20%">Min. Net Value:</td> 
      <td colspan="3">
      	     
<input name="NetValue" type=text value="<%= Search_NetValue %>" size="15">&nbsp;   
 
    </tr>

	
		
		
	<tr> 
	
	<td width="20%" > 
	<input type="radio" name="Statement" value="Monthly" onClick="" <% if Search_Statement="Monthly" then response.write "checked"%>>&nbsp;Monthly Statement:</td> 
	<td >
			<select name="SMMonth" class="common">            	
					<option value="1"  <% if Search_Monthly_Month=1 then response.write "selected"%>>Jan</option>
					<option value="2"  <% if Search_Monthly_Month=2 then response.write "selected"%>>Feb</option>
					<option value="3"  <% if Search_Monthly_Month=3 then response.write "selected"%>>Mar</option>
					<option value="4"  <% if Search_Monthly_Month=4 then response.write "selected"%>>Apr</option>
					<option value="5"  <% if Search_Monthly_Month=5 then response.write "selected"%>>May</option>
					<option value="6"  <% if Search_Monthly_Month=6 then response.write "selected"%>>Jun</option>
					<option value="7"  <% if Search_Monthly_Month=7 then response.write "selected"%>>Jul</option>
					<option value="8"  <% if Search_Monthly_Month=8 then response.write "selected"%>>Aug</option>
					<option value="9"  <% if Search_Monthly_Month=9 then response.write "selected"%>>Sep</option>
					<option value="10" <% if Search_Monthly_Month=10 then response.write "selected"%>>Oct</option>
					<option value="11" <% if Search_Monthly_Month=11 then response.write "selected"%>>Nov</option>
					<option value="12" <% if Search_Monthly_Month=12 then response.write "selected"%>>Dec</option>
			</select>
	
	<select name="SMYear" class="common">            
			<% 
			Dim Year_starting
			Dim Year_ending
			
			Year_starting = Year(DateAdd("yyyy", -8, Now()))
			year_ending = Year(Now())
			
			for i=Year_starting to Year_ending
			%>			         
			<option value="<%=i%>" <% if clng(i)=clng(Search_Monthly_Year) then response.write "selected"%>><%=i%></option>
			
			<% next %>
	</select>
	</td>
	<td >
	<input type="radio" name="Statement" value="Daily" onClick="" <% if Search_Statement="Daily" then response.write "checked"%>> Daily Statement :</td> 
	<td >
		
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


Year_starting = Year(DateAdd("yyyy", -8, Now()))
year_ending = Year(Now())

for i=Year_starting to Year_ending
%>			         
			<option value="<%=i%>" <% if clng(i)=clng(Search_Daily_Year) then response.write "selected"%>><%=i%></option>

<% next %>

			</select>

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
			<td colspan="4"><input type="radio" name="ShareSelection" value="share1" onClick=""  <%if Search_SharedSelection = "share1" then response.write "checked" end if %>  >Viewing AE only &nbsp;|&nbsp;
			
			<input type="radio" name="ShareSelection" value="share2" onClick=""  <%if Search_SharedSelection = "share2" then response.write "checked" end if %>   > Particular AE in the Sales Team
			&nbsp;

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
			&nbsp;

          	|

          <input type="radio" name="ShareSelection" value="share3" onClick=""  <%if Search_SharedSelection = "share3" then response.write "checked" end if %>   > All AEs in the Sales Team 
			&nbsp;
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
 	<input type=hidden   name="submitted"> 
			<input id="Submit1" type="button" value="Submit" onClick="dosubmit(1);"></td>

		</tr>
	


	</table>
	<br>
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
	

 'Rs return 2 value
 '1) Total number of matched client
 '2) all records for targeted client

	
'response.write ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '"&session("shell_power")&"', '"&Search_SharedGroup&"','"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&iPageCurrent&"', '1' ")
 
	Select Case  Search_SharedSelection
	case "share2"
		'shared group member

	'		response.write ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&iPageCurrent&"', '1' ")
			Rs1.open ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_NetValue&"','"&iPageCurrent&"', '1' ") ,  StrCnn,3,1

	case "share3"

		'	response.write ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"','"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&iPageCurrent&"', '1' ")
			
			Rs1.open ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"','"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_NetValue&"','"&iPageCurrent&"', '1' ") ,  StrCnn,3,1
	case else
		'normal
			'response.write ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_NetValue&"','"&iPageCurrent&"', '1' ")
			Rs1.open ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_NetValue&"','"&iPageCurrent&"', '1' ") ,  StrCnn,3,1		
	end select	 
 'Rs1.open ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&iPageCurrent&"', '1' ") ,  StrCnn,3,1
 

 
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

</span>


<div id="reportbody1" >
<script type="text/javascript">
var somediv=document.getElementById("reportbody1")
disableSelection(somediv) //disable text selection within DIV with id="mydiv"
</script>	

<DIV align=center >
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




<table width="99%" border="0" class="normal"  cellspacing="1" cellpadding="2">
<tr bgcolor="#FFFFCC">
		<td  width="30%">　</td> 
      <% if Search_statement="Monthly" then %>
      	<td align="center" width="40%">月結單<br><u>Monthly Statement</u></td>
      <%else %>
				<td align="center" width="40%">
					<% if cint(rs1("Quantity")) > 0 then %>
							綜合日結單/買賣合約<br><u>Daily Combined Statement/Contract Note</u></td>
					<% else %>
							綜合日結單<br><u>Daily Combined Statement</u></td>
					
					<% end if%>
			<% end if %>
      <td align="right" valign="bottom" width="30%">　</td>
</tr>


<tr bgcolor="#FFFFCC">
			<td  width="30%"></td> 
			<td  width="40%"></td>      
			<td align="right" valign="bottom" width="30%" id="noprint">

						<% If Session("PrintAllowed") = 1 then %>  
							<a href="javascript:window.print()">Friendly Print</a>
						<% end if %>
			</td>
		
</tr>
</table>


<br>

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
						<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
					
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
						
						<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
						
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
											<td width="14%"><% = FormatDateTime(now(),2) %></td>
											<td width="16%">  <img border=0 src='images/tel.gif' onClick="PopupClientContact('<%=rs1("clnt") %>')"></img><%=rs1("clnt") %></td>
<% 

       Client_Fun = Rs1("clnt") 
       
%>
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


<%
'''''''''
' Loop for next records
'''''''''

do while (Not rs1.EOF)
			
			Select Case rs1("sectioncode")
					case "IN" exit do
					case "CM" exit do
					case "SM" exit do
					case "SP" exit do
					case "CB" exit do
					'case "MG" exit do
					case "CN" exit do
			
			end select
			rs1.movenext
loop

'''''''''
%>

<% 

''''''''''''''''''''''''''
' Cash statement (CM)
''''''''''''''''''''''''''

Dim TotalBalance
Dim LastCurrency
LastCurrency = ""
if (  Not rs1.EOF)  then
			
			' Cash statement section is including Interest section
			if (rs1("sectioncode") = "CM" or rs1("sectioncode") = "IN" ) then %>
			
						
						<br>
						<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
									
									<tr bgcolor="#FFFFCC"> 
												<td width="95%" colspan="12" ><span lang="zh-tw">現金帳戶記錄 Cash Statement</td>
									</tr>
									<tr bgcolor="#ADF3B6"> 
												<td width="5%">Date<br>交易日期</td>
												<td width="6%">Sdate<br>交收日期</td>
												<td width="6%">Ref No<br>參考編號</td>
												<td width="10%">Stock Code<br>股票代號</td>
												<td width="6%">Bought/In<br>買 / 入</td>
												<td width="4%">Sold/Out<br>沽 / 出</td>
												<td width="21%">Description<br>摘要</td>
												<td width="9%">Price<br>單價</td>
												<td width="7%">CCY<br>貨幣</td>
												<td width="7%">Debit<br>支帳/結欠</td>
												<td width="7%">Credit<br>存帳/結存</td>
												<td width="10%">Balance<br>結餘</td> 
									</tr>


									
									
									<% 
									''''''''''''''''''''''''''
									' Cash statement (CM)

									'TotalBalance = cdbl(rs1("Openingbalance"))
									if rs1("sectioncode") = "CM" then
											do while (  Not rs1.EOF)
											
														if rs1("sectioncode") = "CM" then
														
																' if currency change, display the balance B/F statement
																if rs1("ccy") <> LastCurrency then
																TotalBalance = cdbl(rs1("Openingbalance"))
		
																	%>
																	<tr bgcolor="#FFFFCC">
																				<td colspan="5" align="right">承上結餘Balance B/F </td>
																				<td colspan="3" align="right"> </td>
																				<td><%=rs1("CCY")%> </td>
																				<td colspan="2" align="right"> </td>
																			  <td><%=formatnumber(rs1("Openingbalance"),2,-2,-1)%> </td>
								
																	</tr>									
										            <% end if 
																	TotalBalance = TotalBalance + cdbl(rs1("Amount"))
										            
										            %>

																	
																	
																	
																	<tr bgcolor="#FFFFCC"> 
																	<td width="5%" height="19"><%=rs1("TradeDate") %>　</td>
																	<td width="6%" height="19"><%=rs1("SettleDate") %>　</td>
																	<td width="6%" height="19"><%=rs1("RefNo") %>　</td>
																	<td width="10%" height="19"><%=rs1("Instrument") %>　</td>
																	<td width="6%" height="19"><% if rs1("Quantity") > "0" then response.write formatnumber(rs1("Quantity"),0,-2,-1) %>　</td>
																	<td width="4%" height="19"><% if rs1("Quantity") < "0" then response.write formatnumber(rs1("Quantity"),0,-2,-1) %>　</td>
																	<td width="21%" height="19"><%=rs1("comment") %>　</td> 
																	<td width="9%" height="19"><% if rs1("Price") <> "0" then response.write formatnumber(rs1("Price"),4,-2,-1) %>　</td> 
																	<td width="7%" height="19"><%=rs1("CCY") %>　</td> 
																	<td width="7%" height="19"><% if rs1("Amount") < "0" then response.write formatnumber(rs1("Amount"),2,-2,-1) %>　</td> 
																	<td width="7%" height="19"><% if rs1("Amount") > "0" then response.write  formatnumber(rs1("Amount"),2,-2,-1) %>　</td> 
																	<td width="10%" height="19"><%=formatnumber(TotalBalance, 2,-2,-1)%></td> 
																	</tr>
																	<%
		
																	LastCurrency = rs1("ccy") 
																	rs1.movenext
																	
																	if not rs1.eof then

																			'display  Balance C/F if currency changed
																			if ( (rs1("ccy") <> LastCurrency and rs1("sectioncode") = "CM") or (rs1("sectioncode") <> "CM") ) then
																				%>
																						<tr bgcolor="#FFFFCC">
																									<td colspan="5" align="right">結餘轉下Balance C/F </td>
																									<td colspan="3" align="right"> </td>
																									<td><%=LastCurrency%> </td>
																									<td colspan="2" align="right"> </td>
																								  <td><%=formatnumber(TotalBalance,2,-2,-1)%> </td>
													
																						</tr>								
													            <% 
													            end if 

																			if (rs1("sectioncode") <> "CM")  then 
																					exit do
																			end if
																	end if
														end if
																				
												loop
									end if		
									%>
									
								
									<%

									''''''''''''''''''''''''''
									' Daily Interest (IN)

									if (  Not rs1.EOF)  then
											if rs1("sectioncode") = "IN" then
													 do while (  Not rs1.EOF)
													
																if rs1("sectioncode") = "IN"  then
																%>
																
																		<tr bgcolor="#FFFFCC">
																					<td width="5%"><%=rs1("ConfirmationDate")%></td>
																					<td colspan="7" align="right">Daily Interest Accrued </td>
																					<td width="5%"><%=rs1("CCY")%></td>
																					<td width="21%" colspan="2" align="right"> </td>
																				  <td width="10%"><%=formatnumber(rs1("MTDDebitInterest"),2,-2,-1)%> </td>
									
																		</tr>		
																<%
																end if	
																rs1.movenext
																
																if not rs1.eof then
																		if rs1("sectioncode") <> "IN" then 
																				exit do
																		end if
																end if
													loop

													'rs1.movenext
												end if
									end if 
								%>
						</table>

						<%
			
			end if 
end if

%>


<%
'''''''''
' Loop for next records
'''''''''

do while (Not rs1.EOF)
			
			Select Case rs1("sectioncode")
					case "IN" exit do
					case "CM" exit do
					case "SM" exit do
					case "SP" exit do
					case "CB" exit do
					'case "MG" exit do
					case "CN" exit do
			
			end select
			rs1.movenext
loop

'''''''''
%>



<% 
	''''''''''''''''''''''''''
	' Stock Journal (SM)
	''''''''''''''''''''''''''


if (  Not rs1.EOF)  then
			if rs1("sectioncode") = "SM" then %>
						<br>			
						<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
								
									<tr bgcolor="#FFFFCC"> 
												<td width="95%" colspan="9" >股票帳戶記錄 Stock Journal for </td>
									</tr>
									
									
									<tr bgcolor="#ADF3B6"> 
												<td width="5%">Date<br>交易日期</td>
												<td width="6%">Sdate<br>交收日期</td>
												<td width="6%">Ref No<br>參考編號</td>
												<td width="10%">Market<br>市場</td>
												<td width="10%">Stock Code<br>股票代號</td>
												<td width="6%">Bought/In<br>買/入</td>
												<td width="4%">Sold/Out<br>沽/出</td>
												<td width="15%">Description<br>摘要</td>
												<td width="20%">Stock Name<br>股票名稱</td> 
									</tr>
									
									<% do while (  Not rs1.EOF)
									
												if rs1("sectioncode") = "SM"  then
												%>
												
															<tr bgcolor="#FFFFCC"> 
																		<td width="5%"><%= rs1("tradedate") %></td>
																		<td width="6%"><%= rs1("settledate") %>　</td>
																		<td width="6%"><%= rs1("RefNo") %>　</td>
																		<td width="10%"><%= rs1("Market") %>　</td>
																		<td width="10%"><%= rs1("Instrument") %>　</td>
																		<td width="6%"><%if clng(rs1("Quantity")) > 0 then response.write formatnumber(rs1("Quantity"),0) end if%>　</td>
																		<td width="4%"><%if clng(rs1("Quantity")) < 0 then response.write formatnumber(abs(cDbl(rs1("Quantity"))),0) end if%>　</td>
																		<td width="15%"><%= rs1("Comment") %>　</td> 
																		<td width="20%"><%= rs1("InstrumentDesc") %>　</td> 
															</tr>
															<%
												end if	
												rs1.movenext
												
												if not rs1.eof then
														if rs1("sectioncode") <> "SM" then 
																exit do
														end if
												end if
									loop
									
									
									
									%>
						</table>
			<%
			end if 

end if
%>

<%
'''''''''
' Loop for next records
'''''''''

do while (Not rs1.EOF)
			
			Select Case rs1("sectioncode")
					case "IN" exit do
					case "CM" exit do
					case "SM" exit do
					case "SP" exit do
					case "CB" exit do
					'case "MG" exit do
					case "CN" exit do
			
			end select
			rs1.movenext
loop

'''''''''
%>


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
						<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
								
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
																			<td width="17%"><%=formatnumber(cDbl(rs1("MarginValue")),2,-2,-1)%></td> 
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
												<td colspan="10" align="right">
													<b>股份組合市值( Portfolio Value (HKD): </b></td>
												<td><b><%=formatnumber(TotalPortfolioValue,2,-2,-1) %></b></td>

													<%
													if AccountType = "MRGN" then 
														response.write "<td></td><td><b>"
														response.write formatnumber(TotalMarginValue,2,-2,-1) 
														response.write "</b></td>"
													end if %>
													
												</td>
									</tr>
									
						</table>
			<%
			
			end if 
end if

%>

<%
'''''''''
' Loop for next records
'''''''''

do while (Not rs1.EOF)
			
			Select Case rs1("sectioncode")
					case "IN" exit do
					case "CM" exit do
					case "SM" exit do
					case "SP" exit do
					case "CB" exit do
					'case "MG" exit do
					case "CN" exit do
			
			end select
			rs1.movenext
loop

'''''''''
%>


<%
	''''''''''''''''''''''''''
	'  Cash Balance (CB)
	''''''''''''''''''''''''''

dim TotalCashValue
dim TotalNetValue
dim TotalMarginNetValue

TotalCashValue=0
TotalNetValue=0
TotalMarginNetValue=0

'Response.write rs1("sectioncode")

if (  Not rs1.EOF)  then
			if rs1("sectioncode") = "CB" then %>

			
						<br>
						<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
								
						
									<tr bgcolor="#FFFFCC"> 
												<td width="115%" colspan="8">現金結餘 Cash Balance</td>
									</tr>
									
									<tr bgcolor="#ADF3B6">
												<td width="14%">CCY<br>貨幣</td>
												<td width="16%">Available Balance<br>可用結餘</td>
												<td width="12%">Unavailable Balance<br>未可用結餘</td>
												<td width="26%">Ledger Balance<br>帳面結餘</td>
												<td width="10%">## Interest Accrued<br>## 應計利息</td>
												<td width="10%">#Exchange Rate<br># 匯率</td>
												<td width="10%">Balance (HKD)<br>結餘</td>
												<% if AccountType = "MRGN" then %>
														<td width="17%"></td> 
												<% end if %>
									</tr>
									
									
									<%  
									do while ( Not rs1.EOF)
									
												if rs1("sectioncode") = "CB" then
												%>
												
												<tr bgcolor="#FFFFCC"> 
															<td width="14%"><%= rs1("CCY") %></td>
															<td width="16%"><%= formatnumber(rs1("OpeningBalance"),2,-2,-1) %>　</td>
															<td width="12%"><%= formatnumber(rs1("Netbalance"),2,-2,-1) %>　</td>
															<td width="26%"><%= formatnumber(rs1("EndingBalance"),2,-2,-1) %>　</td>
															<td width="10%"><%= formatnumber(rs1("MTDDebitInterest"),2,-2,-1)  %>　</td>
															<td width="10%"><%= formatnumber(rs1("MarginLimit"),4,-2,-1) %>　</td>
															<td width="10%"><%= formatnumber(rs1("StockPortValue"),2,-2,-1) %>　</td>
															<% if AccountType = "MRGN" then %>
																	<td width="17%"></td> 
															<% end if %>
															
												</tr>
												<%
												TotalCashValue=TotalCashValue+cdbl(rs1("StockPortValue"))
												end if	
												rs1.movenext
												
												
												if not rs1.eof then
														if rs1("sectioncode") <> "CB" then 
																exit do
														end if
												end if
									loop
									%>
									
									<%
										TotalNetValue = TotalCashValue  + TotalPortfolioValue
										TotalMarginNetValue = TotalBalance + TotalMarginValue
									%>
									<tr bgcolor="#FFFFCC">
												<td colspan="6" align="right"><b>現金值 Cash Value: </b></td>
												<td><%=formatnumber(TotalCashValue,2,-2,-1) %></td>
															<% if AccountType = "MRGN" then %>
																	<td width="17%"></td> 
															<% end if %>
								  </tr>
								  <tr bgcolor="#FFFFCC">
												<td colspan="6" align="right"><b>總值 Net Value: </b></td>
												<td><b><%=formatnumber(TotalNetValue,2,-2,-1) %></b></td>
												
													<%
													if AccountType = "MRGN" then %>
														<td><b> <%= formatnumber(TotalMarginNetValue,2,-2,-1) %>
															</b>
														</td>
													<% end if %>												
									</tr>
									
									<% if AccountType = "MRGN" then %>
									<tr bgcolor="#FFFFCC">
												<td colspan="7" align="right"></td>
												
													<%
													
														if TotalMarginNetValue >= 0 then
																	response.write "<td colspan=1><b>Margin Available</b></td>"
														else
																	response.write "<td colspan=1><b>Margin Call</b></td>"
														end if
													%>
													
								  </tr>
									<% end if 
									
									'foot note
									%>
									<tr bgcolor="#FFFFCC">
											<td colspan="4">
												# Exchange Rate used are for indication only 匯率只作指引用途 <br>
												## Interest Accrued  will be credited/debited to your account on the last day of each month
												<br> 應計利息將會於每月的最後一天從閣下之戶口存入\提取
											<td>
											<td><td>
											<% if AccountType = "MRGN" then %>
													<td width="17%"></td> 
											<% end if %>									</tr>	
						</table>
						
						<%
						' Foot Note
						
						
						
			end if
					
			 
'end if
else
%>
<br>
			<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">			 <tr bgcolor="#FFFFCC">
					<td align="right"><b>總值 Net Value: </b></td>
					<td width="20%"><b><%=formatnumber(TotalPortfolioValue,2,-2,-1) %></b></td>
												
																						
			  </tr>
			</table>
<%
			End If	

%>

<%
'''''''''
' Loop for next records
'''''''''

do while (Not rs1.EOF)
			
			Select Case rs1("sectioncode")
					case "IN" exit do
					case "CM" exit do
					case "SM" exit do
					case "SP" exit do
					case "CB" exit do
					'case "MG" exit do
					case "CN" exit do
			
			end select
			rs1.movenext
loop

'''''''''
%>

<% 
	''''''''''''''''''''''''''
	'  Trade Detail (CN)
	''''''''''''''''''''''''''

if (  Not rs1.EOF)  then
	if rs1("sectioncode") = "CN" then 
	
	
	'temp instrument datastore
	dim j,k

  dim Market
  dim RefNo
  dim TradeDate
  dim SettleDate
  dim Instrument
  dim BuySell
  dim CCY  
  dim NetBalance
  dim Amount 
  dim InstrumentDesc	
  dim InstrumentCDesc	
	dim orfee(10)
	dim feename(10)
	dim quantity()
	dim price()
 
	%>


<br>

<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

						
<tr bgcolor="#FFFFCC"> 
      <td width="166%" height="18" colspan="12">交易詳情 <span lang="en-us">Trade Details</span></td>
</tr>

<tr bgcolor="#ADF3B6">
   <td width="8%" height="18">Market<br>市場</td>
   <td width="5%" height="18">Trade No<br>交易編號</td>
   <td width="5%" height="18">T/D <br>交易</td>
   <td width="5%" height="18">S/D<br>交收</td>
   <td width="18%" height="18">Stock<br>股票</td>
   <td width="10%" height="18">B/S<br>買沽</td>
   <td width="5%" height="18">Quanlity<br>數量</td>
   <td width="5%" height="18">Price<br>成交價</td>
   <td width="5%" height="18">CCY<br>貨幣</td> 
   <td width="10%" height="18">Gross Amount<br>總額</td> 
   <td width="17%" height="18">Transaction Cost<br>交易費用</td> 
   <td width="10%" height="18">Net Amount<br>淨額</td> 
</tr>
		<% do while (  Not rs1.EOF)

				
				if rs1("sectioncode") = "CN" and rs1("comment") = "1" then
					j=0


					'destroy array
					erase price
					erase quantity		

				  Market          = rs1("Market")
				  RefNo           = rs1("RefNo")
				  TradeDate       = rs1("TradeDate")
				  SettleDate      = rs1("SettleDate")
				  Instrument      = rs1("Instrument")
				  BuySell         = rs1("BuySell")
				  CCY             = rs1("CCY")
				  NetBalance      = rs1("NetBalance")
				  InstrumentDesc	= rs1("InstrumentDesc")
				  InstrumentCDesc	= rs1("InstrumentCDesc")
					Amount          = rs1("Amount")
					
					feename(0) = rs1("feename1")
					feename(1) = rs1("feename2")
					feename(2) = rs1("feename3")
					feename(3) = rs1("feename4")
					feename(4) = rs1("feename5")
					feename(5) = rs1("feename6")
					feename(6) = rs1("feename7")
					feename(7) = rs1("feename8")
					feename(8) = rs1("feename9")
					feename(9) = rs1("feename10")

					orfee(0) = rs1("ORFee1")
					orfee(1) = rs1("ORFee2")
					orfee(2) = rs1("ORFee3")
					orfee(3) = rs1("ORFee4")
					orfee(4) = rs1("ORFee5")
					orfee(5) = rs1("ORFee6")
					orfee(6) = rs1("ORFee7")
					orfee(7) = rs1("ORFee8")
					orfee(8) = rs1("ORFee9")
					orfee(9) = rs1("ORFee10")
							
					'store orfee and quantity into array
					
					'loop to next new ref no
					do while (Not rs1.EOF) 'find all quantity and price of trading

						
						if  rs1.eof or rs1("RefNo") <> RefNo  then 
						%>

		
						<%  'write table content %>
						<tr bgcolor="#FFFFCC"> 
						   <td><%= Market %>　</td>
						   <td><%= RefNo %>　</td>
						   <td><%= TradeDate %>　</td>
						   <td><%= SettleDate %>　</td>
						   <td><%= Instrument %>&nbsp;<%= InstrumentDesc %>&nbsp;<%= InstrumentCDesc %>　</td>
						   <td><%= BuySell %>　   </td>
						   <td><br>
						   <%
						   For Each item In quantity
								if item <> "" then
									Response.Write(formatnumber(item,0) & "<br>")
								end if
								Next
							%>　  
						   </td>
						   <td><br>
						   <%
						   For Each item In price
								if item <> "" then
									Response.Write(formatnumber(item,4) & "<br>")
								end if
								Next
							%>　      
						   </td>
						   
						   <td><%= CCY %>　       </td> 
						   <td><%= formatnumber(NetBalance,2,-2,-1) %>　</td> 
						   <td>
						   <% 
						   For k=0 to 10 
						   	if feename(k) <> ""  and cDbl(orfee(k)) <> 0  then 
						   		Response.Write (feename(k) & ": " & formatnumber(cDbl(orfee(k)),2,-2,-1) & "<br>")
						   	end if
						   next %>
			
						   </td> 
							 <td width="17%" height="18"><%= formatnumber(cDbl(Amount),2,-2,-1) %>　</td> 
						</tr>							
					<%
						'move to next new ref no
						exit do
					else 
				
							' if not the first record, add price and quantity into array
							ReDim Preserve price(j+1)
							ReDim Preserve quantity(j+1)	
							price(j) = rs1("Price")
							quantity(j) = rs1("Quantity")
							j=j+1
							rs1.movenext
							
							
					end if
					

				loop



				
			else
							rs1.movenext
			end if	' end of if rs1("sectioncode") = "CN" and rs1("comment") = "1" then
			

			
			if not rs1.eof then
				if rs1("sectioncode") = "SP" then
					exit do
				end if
			end if
			
		loop
	          %>
		<%  'write table content 
				'write buffered record
		%>
		
						<tr bgcolor="#FFFFCC"> 
						   <td><%= Market %>　</td>
						   <td><%= RefNo %>　</td>
						   <td><%= TradeDate %>　</td>
						   <td><%= SettleDate %>　</td>
						   <td><%= Instrument %>&nbsp;<%= InstrumentDesc %>&nbsp;<%= InstrumentCDesc %>　</td>
						   <td><%= BuySell %>　   </td>
						   <td><br>
						   <%
						   For Each item In quantity
								if item <> "" then
									Response.Write(formatnumber(item,0) & "<br>")
								end if
								Next
							%>　  
						   </td>
						   <td><br>
						   <%
						   For Each item In price
								if item <> "" then
									Response.Write(formatnumber(item,4) & "<br>")
								end if
								Next
							%>　      
						   </td>
						   
						   <td><%= CCY %>　       </td> 
						   <td><%= formatnumber(NetBalance,2,-2,-1) %>　</td> 
						   <td>
						   <% 
						   For k=0 to 10 
						   	if feename(k) <> ""  and cDbl(orfee(k)) <> 0  then 
						   		Response.Write (feename(k) & ": " & formatnumber(cDbl(orfee(k)),2,-2,-1) & "<br>")
						   	end if
						   next %>
			
						   </td> 
							 <td width="17%" height="18"><%= formatnumber(cDbl(Amount),2,-2,-1) %>　</td> 
						</tr>				



 </table>
 
 	<%

      end if 

	end if
	
		

		
       


	''''''''''''''''''''''''''
	'  Fund
	''''''''''''''''''''''''''


    set Rs2 = server.createobject("adodb.recordset")
         Rs2.open ("Exec Retrieve_ClientStatement_fun '"&Client_Fun&"', '"&Client_Fun&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_NetValue&"','"&iPageCurrent&"', '1' ") ,  StrCnn,3,1		
   response.write ("Exec Retrieve_ClientStatement_fun '"&Client_Fun&"', '"&Client_Fun&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_NetValue&"','"&iPageCurrent&"', '1' ")


if Not Rs2.EoF then

   

	%>


<br>

<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

						
<tr bgcolor="#FFFFCC"> 
      <td width="166%" height="18" colspan="12">在途交易基金詳情 Unsettled Detail Fund for <span lang="en-us"><% = FormatDateTime(Rs2("StatementDate"),1) %></span></td>
</tr>

<tr bgcolor="#ADF3B6">
   <td width="8%" height="18">Market<br>市場</td>
   <td width="5%" height="18">Order No<br>交易編號</td>
   <td width="5%" height="18">O/D<br>交易</td>
   <td width="18%" height="18">Fund<br>股票</td>
   <td width="10%" height="18">B/S<br>買沽</td>
   
   <td width="5%" height="18">Unit<br>成交價</td>
   <td width="5%" height="18">CCY<br>貨幣</td> 
   <td width="10%" height="18">Gross Amount<br>總額</td> 
   <td width="17%" height="18">Transaction Cost<br>交易費用</td> 
   <td width="10%" height="18">Net Amount<br>淨額</td> 
</tr>
		<% 
               
                     Rs2.MoveFirst

                  Do While Not Rs2.EoF
					
						%>

	<tr bgcolor="#FFFFCC"> 
						   <td><%= Rs2("Market") %></td>
						   <td><%= Rs2("TradeNo") %></td>
						   <td><%= Rs2("TradeDate") %></td>
						   <td><%= Rs2("Instrument") %>&nbsp;<%= Rs2("InstrumentDesc") %>&nbsp;<%= Rs2("InstrumentCDesc") %>　</td>
						   <td><%
                                   
                                   If Trim(Rs2("BuySell")) = "BUY" Then

                                      Response.Write "Subscription"

                                   Elseif Trim(Rs2("BuySell")) = "SELL" Then

                                      Response.write "Redemption"

                                   End If                                     
                                 

                                %>　   </td> 
						   
						   <td>
						   <% = Rs2("LotSize") %></td> 
							 <td width="17%" height="18"><%= Rs2("ccy") %>　</td> 
<td width="17%" height="18"><%= formatnumber(cDbl(rs2("consideration")),2,-2,-1) %>　</td>
<td width="17%" height="18"><%= formatnumber(cDbl(Rs2("ORFee1")),2,-2,-1) %>　</td>
<td width="17%" height="18"><%= formatnumber(cDbl(Rs2("NetAmount")),2,-2,-1) %>　</td>
						</tr>							
					<%
									
                          Rs2.MoveNext
         
				   loop

                          Set rs2 = rs2.NextRecordset() 
%>

                       <tr bgcolor="#FFFFCC">
                           <td colspan="9" align="right">
                           <% = Rs2("CCY") %>
                           </td>
                           <td colspan="3" align="right">
                           <% = formatnumber(Rs2("TotalBalance"),2,-2,-1) %>
                           </td>
                       </tr>

	


 </table>
 
 	<% End if %>






		<br>
		<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
		<tr bgcolor="#FFFFCC"> 
					<td width="166%" height="18" colspan="12">
						1. For Margin Accounts, the margin limit is subject to change from time to time
						at our sole discretion. Your current margin limit is stated in the first page of 
						this statement
						 <br>
						 2. Stock Borrowing Services for Hong Kong Market is now available in UOB KayHian 
						(Hong Kong) Limi ted, for more details, please contact your Account Manager or customer
						service hotline at 2826 4868. 
						<br>
						3. Effective from 11 Aug 2009, we require all clients to sign the W-8BEN certificate to 
						certify their foreign status of US before trading US Securities. W-8BEN certificate is a 
						Certificate of Foreign Status of US. With W-8BEN certificate, the client wil l be exempted 
						from US Capital Gains Tax. For further information, kindly contact your responsible salespersons.
			</td>
		</tr>		
	</table>
	<br>
		<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
		<tr bgcolor="#FFFFCC"> 
					<td width="166%" height="18" colspan="12" align="center">結單完結<br>End of Statement</td>
		</tr>
		
	</table>
	</td>
	</tr>
</table>

</div>      



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
</span>


</td></tr></table>
</div>
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