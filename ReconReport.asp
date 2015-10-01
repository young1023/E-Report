<% 
'*********************************************************************************
'NAME       : ReconReport.asp           
'DESCRIPTION: Recon File Convertion Report
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 150822 Gary Yeung   Prototype
'MODIFIED   : 
'			
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

Title = "Stock Reconciliation Report"




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

Dim itotalturnover(), itotalconsideration(), itotalbrokerage(), itotalCCY()
Dim iPageturnover(), iPageconsideration(), iPagebrokerage(), iPageCCY()


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
		if ((datevalidate(document.fm1.FromDay.value, document.fm1.FromMonth.value -1, document.fm1.FromYear.value) == false) || 
				(datevalidate(document.fm1.ToDay.value, document.fm1.ToMonth.value -1, document.fm1.ToYear.value) == false))
		{
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



Search_From_Day         = Request.form("FromDay")
Search_From_Month       = Request.form("FromMonth")
Search_From_Year        = Request.form("FromYear")
Search_To_Day           = Request.form("ToDay")
Search_To_Month         = Request.form("ToMonth")
Search_To_Year          = Request.form("ToYear")
Search_Market           = Request.form("Market")
Search_Instrument       = Request.form("Instrument")
Search_Order            = Request.form("Order")
Search_Direction        = Request.form("Direction")




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


Else
	iPageCurrent = Clng(Request.form("page"))
End If



set RsMarket = server.createobject("adodb.recordset")
RsMarket.open ("Exec Retrieve_AvailableMarket ") ,  StrCnn,3,1



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

<form name="fm1" method="post" action="">
  <table width="97%" border="0" class="normal">


 
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


Year_starting = Year(DateAdd("yyyy", -2, Now()))
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


Year_starting = Year(DateAdd("yyyy", -2, Now()))
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
			<option value="" <% if Search_Market="" then response.write "selected" %> >All</option>
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
	

 'Rs return 2 value
 '1) Total number of matched client
 '2) all records for targeted client

'iRecord = (iPageCurrent -1) * RECORDPERPAGE  +1

	'Response.write 	("Exec Retrieve_StockReconciliation  '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '1' ") 

 	Rs1.open ("Exec Retrieve_StockReconciliation  '"&Search_From_Day&"', '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_To_Day&"', '"&Search_To_Month&"', '"&Search_To_Year&"', '"&Search_Market&"','"&Search_Instrument&"', '"&iPageCurrent&"', '1' ") ,  StrCnn,3,1
			

	'Rs1.open ("Exec Retrieve_StockReconciliation  '1', '2', '2009', '1', '2', '2009', '','', '1', '1') ,  StrCnn,3,1


		If Err.Number <> 0 then
			
			'SQL connection error handler
			response.write  "<table><tr><td class='RedClr'>" & MSG_BUSY & "<br></td></tr></table>"

			End If



  
  dim itotalCCYcount
	erase itotalturnover
	erase itotalconsideration		
	erase itotalBrokerage		
  itotalCCYcount=0
  
	do while (  Not rs1.EOF)

			itotalCCYcount=itotalCCYcount+1


			ReDim Preserve itotalturnover(itotalCCYcount+1)
			ReDim Preserve itotalconsideration(itotalCCYcount+1)
			ReDim Preserve itotalBrokerage(itotalCCYcount+1)
			ReDim Preserve itotalCCY(itotalCCYcount+1)


			iRecordCount = iRecordCount + rs1("totalrecordcount") 'total number of records
			itotalturnover(itotalCCYcount) = rs1("totalturnover")
			itotalconsideration(itotalCCYcount) = rs1("totalconsideration")
			itotalBrokerage(itotalCCYcount) = rs1("totalBrokerage")
			itotalCCY(itotalCCYcount) = rs1("CCY")
			
			'			response.write "B" & j	
			rs1.movenext
			
	loop
		
		'response.write iRecordCount & itotalturnover(1) & "AAA"  & itotalconsideration(1)
	
'iRecordCount = 0

  if iRecordCount <= 0 then
		
		
		
		If Err.Number <> 0 then
			
            'response.write Err.Number
			'SQL connection error handler
			response.write  "<table><tr><td class='RedClr'>The server is currently too busy to process your request right now. Please wait a moment and then try again. If the problem persists, please contact systems administrator.<br></td></tr></table>"
			
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


   
<%
'**********
' Start of page navigation 
'**********
%> 
    <DIV align=center>

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

</span>
<%
'**********
' End of page navigation 
'**********
%>



<% if (session("shell_power") = 1 or session("shell_power") = 5) then %>  
		<span class="noprint">
<% end if %>

     
<br>

<table width="99%" border="0" class="normal"  cellspacing="1" cellpadding="2">
<tr bgcolor="#FFFFCC"> 
<td  width="20%">¡@</td>
      <td align="center">¸Ô²Ó¬ö¿ý<br><u>Stock Reconciliation Report</u></td> 
      <td align="right" width="20%">
						<% 'If Session("PrintAllowed") = 1 then %>   
						<%if (session("shell_power") <> 1 and session("shell_power") <> 5) then %>  
							<a href="javascript:window.print()">Friendly Print</a><% 'end if %> &nbsp;&nbsp;
<% End If %>
<%if (session("shell_power") = 8) then %>    
<a href="javascript:window.doConvert()">Excel</a>
						<% end if %>      	
			</td>
</tr>
</table>
<br>
<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

<tr bgcolor="#ADF3B6" align="center">

<%

    for each x in Rs1.Fields

     
%>

      <td><% = x.name %></td>
 
<% Next %>

</tr>

		<%
			dim iPageCCYcount
			dim k
			dim iPageUpdate
			
			ReDim Preserve iPageCCY(1)
			ReDim Preserve iPageturnover(1)
			ReDim Preserve iPageconsideration(1)
			ReDim Preserve iPagebrokerage(1)
			
			iPageCCYcount = 0
			iPageturnover(0) = 0
			iPageconsideration(0)= 0
			iPagebrokerage(0)= 0
			'iPageCCY(0) = ""
			
			dim mystr
			do while (Not rs1.EOF)
				k=1
				
		%>

<tr bgcolor="#FFFFCC"> 

<%

    for each x in Rs1.Fields

     
%>

      <td><% = Rs1(x.name) %></td>
 
<% Next %>
</tr>


<%

				
					
				rs1.movenext
				
		loop
		

%>


<% 
	dim l
	For l=1 to iPageCCYcount 
'Dim itotalturnover(), itotalconsideration(), itotalbrokerage(), itotalCCY()

%>

<tr bgcolor="#FFFFCC"> 

   <td colspan="9" align="right"><% if l=1 then response.write "Subtotal<BR>"%> </td>
   <td ><%=iPageCCY(l)%>&nbsp;<%=formatnumber(iPageconsideration(l)) %></td>

   <td colspan="2" align="right"><% if l=1 then response.write "Subtotal<BR>"%> </td>
   <td ><%=iPageCCY(l)%>&nbsp;<%=formatnumber(iPagebrokerage(l)) %></td>

   <td colspan="8" align="right"><% if l=1 then response.write "Subtotal<BR>"%> </td>
   <td ><%=iPageCCY(l)%>&nbsp;<%=formatnumber(iPageturnover(l)) %></td>

 
</tr>

<% Next %>

<% 
	dim j
	For j=1 to itotalCCYcount 
'Dim itotalturnover(), itotalconsideration(), itotalbrokerage(), itotalCCY()

%>


<tr bgcolor="#FFFFCC"> 
	 
   <td colspan="9" align="right"><% if j=1 then response.write "Total<BR>"%>  </td>
   <td ><%=itotalCCY(j)%>&nbsp;<%=formatnumber(itotalconsideration(j)) %></td>

   <td colspan="2" align="right"><% if j=1 then response.write "Total<BR>"%>  </td>
   <td ><%=itotalCCY(j)%>&nbsp;<%=formatnumber(itotalbrokerage(j)) %></td>

   <td colspan="7" align="right"><% if j=1 then response.write "Total<BR>"%>  </td>
   <td ><%=itotalCCY(j)%>&nbsp;<%=formatnumber(itotalturnover(j)) %></td>

   
</tr>

<% Next %>
</table>
<br>
<br>

<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
	<tr bgcolor="#FFFFCC"> 
      <td width="166%" height="18" align="center">End of Statement</td>
	</tr>
</table>
                
</div>
              </center>



<% if (session("shell_power") = 1 or session("shell_power") = 5) then %>  
		</span>
<% end if %>

<span class="noprint">
<%
'**********
' Start of page navigation 
'**********
%> 
    <DIV align=center>

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

 'Rs1.Close
 set Rs1 = Nothing
 'Rs2.close
 Set Rs2 = Nothing
 Conn.Close
 Set Conn = Nothing
%>
<SCRIPT language=JavaScript>
<!--
function doConvert(){
window.open("ConvertStockReconciliation.asp?Search_Instrument=<%=Search_Instrument%>&Search_Market=<%=Search_Market%>&From_Day=<%=Search_From_Day%>&From_Month=<%=Search_From_Month%>&From_Year=<%=Search_From_Year%>&To_day=<%=Search_To_Day%>&To_Month=<%=Search_To_Month%>&To_Year=<%=Search_To_Year%>"); 

}

//-->
</SCRIPT>