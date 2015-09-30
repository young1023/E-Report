<% 
'*********************************************************************************
'NAME       : margincall.asp           
'DESCRIPTION: Margin Call 
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 090401 Gary Yeung   Prototype
'MODIFIED   : 090425 Roger Wong   Record and page control
'			:  090712 Roger Wong	    Add Shared Group
'MODIFIED	: 090929	  Gary Yeung       Add Excel Function  
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



Title = "Margin Control Summary"




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




function validateUserEntry(){

		
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

function PopupClientContact(clientnumber) {
	 
		var str='ListClientContact.asp?sid=<%=SessionID%>&clientnumber=' + clientnumber
		
		newwindow=window.open(str , "myWindow", 
									"status = 1, height = 300, width = 600, resizable = 1'"  )
		 if (window.focus) {
           newwindow.focus();
       }
 			
}

function PopupClientSummary(clientnumber) {
	 
		var str='PrintClientSummary.asp?sid=<%=SessionID%>&clientnumber=' + clientnumber
		
		newwindow=window.open(str , "myWindow", 
									"status = 1, height = 300, width = 600, resizable = 1'"  )
		 if (window.focus) {
           newwindow.focus();
       }
 			
}

function pagesubmit(what){
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
				if (validateUserEntry() == false)
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
Dim Search_MinDebitBalance
Dim Search_MarginExceedPercent
Dim Search_Order 
Dim Search_AccountType
Dim Search_Direction
Dim Search_SharedSelection 
Dim Search_SharedGroup
Dim Search_SharedGroupMember

Search_AEGroup	    = Request.form("GroupID")
Search_ClientFrom                  = Request.form("ClientFrom")
Search_ClientTo                    = Request.form("ClientTo")
Search_AEFrom                      = Request.form("AEFrom")
Search_AETo                        = Request.form("AETo")
Search_MinDebitBalance             = Request.form("MinDebitBalance")
Search_MarginExceedPercent         = Request.form("MarginExceedPercent")
Search_AccountType                 = Request.form("AccountType")
Search_Order                       = Request.form("Order")
Search_Direction                   = Request.form("Direction")
Search_SharedSelection      =  Request.form("ShareSelection")	
Search_SharedGroup  = Request.form("SharedGroup")
Search_SharedGroupMember  = Request.form("SharedGroupMember")
Search_From_Day         = Day(Session("DBLastModifiedDate"))
Search_From_Month       = Month(Session("DBLastModifiedDate"))
Search_From_Year        = Year(Session("DBLastModifiedDate"))




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
	Search_Order = "CLNTCODE"
	Search_Direction = "ASC"
	Search_MinDebitBalance = 0
	Search_MarginExceedPercent = 1.05

	Search_AEGroup	    = session("GroupID")
	Search_ClientFrom   = session("ClientFrom")
	Search_ClientTo     = session("ClientTo")
	Search_AEFrom       = session("AEFrom")
	Search_AETo         = session("AETo")
	Search_SharedSelection = "share1"


Else
	iPageCurrent = Clng(Request.form("page"))
End If


session("GroupID")               =  Search_AEGroup	               
session("ClientFrom")            =  Search_ClientFrom              
session("ClientTo")              =  Search_ClientTo                
session("AEFrom")                =  Search_AEFrom                  
session("AETo")                  =  Search_AETo
session("MinDebitBalance")       =  Search_MinDebitBalance 
session("MarginExceedPercent")   =  Search_MarginExceedPercent  
session("AccountType")           =  Search_AccountType  
                    


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
set RsGroupID = server.createobject("adodb.recordset")
RsGroupID.open ("Exec Retrieve_AvailableGroupID ") ,  StrCnn,3,1


%>

<%
'*****************************************************************
' Start of form
'*****************************************************************
%>
<form name="fm1" method="post" action="">

  <table width="97%" border="0" class="normal">
		<% if (  session("shell_power") = 3   or session("shell_power") = 4   or session("shell_power") = 8  )	 then %>
		<tr > 
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
			<td colspan="2"  align="right"><font color="red">*</font> Denotes a mandatory field
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
				<input name="ClientTo" type=text value="<% If Search_ClientFrom <> Search_ClientTo Then Response.Write Search_ClientTo End If%>" size="15"></td>
		</tr>
    
    
<%
' Show if branch manager or above
if session("shell_power")>=3 then %>
    
	<tr>
		<td>AE Code: (From)</td> 
		<td>
		<input name="AEFrom" type=text value="<%= Search_AEFrom %>" size="15">
      <img align="top" style="cursor:pointer" onClick="PopupSearchAE()" src="images/search.gif"> </td>
	
		<td width="22%">AE Code: (To)</td> 
		<td width="28%">
		<input name="AETo" type=text value="<% If Search_AEFrom <> Search_AETo Then Response.Write Search_AETo End If%>" size="15"></td>
	</tr>
   
<% end if %>


	<%if (session("shell_power") <> 1 and session("shell_power") <> 5) then %>    
				<tr> 
				<td>Min. Debit Balance:</td> 
				<td>
				   
				<input name="MinDebitBalance" type=text value="<%= Search_MinDebitBalance %>" size="15"></td>
				<td>Margin % exceed:</td> 
				<td>
				<input name="MarginExceedPercent" type=text value="<%= Search_MarginExceedPercent %>" size="15"></td>
				</tr>
    
	<% else %>
	
				<tr> 
					<td>
					   
					<input type=hidden   name="MinDebitBalance" type=text value="<%= Search_MinDebitBalance %>" size="30"></td>
					<input  type=hidden  name="MarginExceedPercent" type=text value="<%= Search_MarginExceedPercent %>" size="30"></td>
				
				
					</td>
				</tr>
    	    
 	<% end if %>  
  
  

    		<tr> 
 	<input type=hidden   value="<%=Search_order%>"   name="Order"> 
 	<input type=hidden   value="<%=Search_Direction%>"   name="Direction"> 
 	<input type=hidden   name="submitted"> 
			<td>Account Type</td> 
			<td colspan="3">
			     
			<select size="1" name="AccountType" class="common">
			<option value="ALL" <% if Search_AccountType="ALL" then response.write "selected" %> >All</option>
			<option value="MRGN" <% if Search_AccountType="MRGN" then response.write "selected" %> >Margin</option>
			<option value="NON-MRGN" <% if Search_AccountType="NON-MRGN" then response.write "selected" %>>Non-Margin</option>
			</select>

      
		
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
			<td>
			<input type=hidden   value="<%=iPageCurrent%>"   name="page"> 
			<input id="Submit1" type="button" value="Submit" onClick="pagesubmit(1);"></td>

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



	'response.write ("exec Retrieve_MarginCall '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '"&Search_MinDebitBalance&"','"&Search_MarginExceedPercent&"', '"&Search_AccountType&"','"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ")
	
'	Rs1.open  ("exec Retrieve_MarginCall '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '"&Search_MinDebitBalance&"','"&Search_MarginExceedPercent&"', '"&Search_AccountType&"','"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1


	Select Case  Search_SharedSelection
	case "share2"
		'shared group member
		'	response.write  ("exec Retrieve_MarginCall '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '',  '"&Search_MinDebitBalance&"','"&Search_MarginExceedPercent&"', '"&Search_AccountType&"','"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ")

			Rs1.open  ("exec Retrieve_MarginCall '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '',  '"&Search_MinDebitBalance&"','"&Search_MarginExceedPercent&"', '"&Search_AccountType&"','"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1

	case "share3"
		'	response.write  ("exec Retrieve_MarginCall '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_MinDebitBalance&"','"&Search_MarginExceedPercent&"', '"&Search_AccountType&"','"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") 

		Rs1.open  ("exec Retrieve_MarginCall '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_MinDebitBalance&"','"&Search_MarginExceedPercent&"', '"&Search_AccountType&"','"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1

	case else
		'normal
		'	response.write  ("exec Retrieve_MarginCall '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_MinDebitBalance&"','"&Search_MarginExceedPercent&"', '"&Search_AccountType&"','"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ")

			Rs1.open  ("exec Retrieve_MarginCall '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_MinDebitBalance&"','"&Search_MarginExceedPercent&"', '"&Search_AccountType&"','"&iPageCurrent&"', '"&RECORDPERPAGE&"', '"&Search_Order&"', '"&Search_Direction&"' ") ,  StrCnn,3,1

	end select	


	'assign total number of pages
	iRecordCount = rs1(0)


  if iRecordCount <= 0 then
		
		If Err.Number <> 0 then
			
			'SQL connection error handler
			response.write  "<table><tr><td class='RedClr'>" & MSG_BUSY & "<br></td></tr></table>"
			
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
    
    <DIV align=center>

  <TABLE border=0 cellPadding=0 cellSpacing=0 width=97% >

 <tr > 
 <td class="NavaMenu" align="right" height="28" >
<%if PrintAllowed = 1 then %>  
							<a href="javascript:window.print()">Friendly Print</a>&nbsp;
							<% end if %>
						<%if (session("shell_power") = 8) then %>  

							&nbsp;<a href="javascript:window.doConvert()">Excel</a>
	<% end if %>
&nbsp;&nbsp;
		<%
'**********
' Start of page navigation 
'**********

response.write (iPageCurrent & " Pages " & iPageCount &"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" )

'First button
%>
	<a href=javascript:pagesubmit(1) style='cursor:hand'>First</a>

<%
' Prev button
If iPageCurrent > 1 Then
	%>
	<a href=javascript:pagesubmit(<%= iPageCurrent-1 %>) style='cursor:hand'>Previous</a>
<% else %>
Previous
	<%
End If


'Next button
If iPageCurrent < iPageCount Then
	%>
	<a href=javascript:pagesubmit(<%= iPageCurrent+1 %>) style='cursor:hand'>Next</a>
<% else %>
Next
	<%
End If
%>

<%
'Last button
%>

<a href=javascript:pagesubmit(<%= iPageCount %>) style='cursor:hand'>Last</a>

</td></tr></table>


<%
'**********
' End of page navigation 
'**********
%>
</span>
<br>

		<table width="99%" border="0"  class="normal" cellspacing="1" cellpadding="2">
			<tr bgcolor="#FFFFCC">
					<td align="center">按倉資料<br><u><% = Title %></u></td>
			</tr>
		
		</table>

<br>
		
		<table  width="99%" border="0" class="normal" style="border-width: 0; text-align:right" bgcolor="#808080" cellspacing="1" cellpadding="2">


			<tr bgcolor="#ADF3B6">
			   <td width="10%"><a href=javascript:ordersubmit(<% call OrderVariable("CLNTCODE")  %>) >Client No.<br>客戶編號</a></td>
			   <td width="18%"><a href=javascript:ordersubmit(<% call OrderVariable("CLNTNAME")  %>) >Client Name<br>客戶名稱</a></td>
			   <td width="5%"><a href=javascript:ordersubmit(<% call OrderVariable("CCY")  %>) >Currency<br>貨幣</a></td>
			   <td width="10%"><a href=javascript:ordersubmit(<% call OrderVariable("CURRENCY")  %>) >Cash Balance to be settled (HKD)</a></td>
			   <td width="5%"><a href=javascript:ordersubmit(<% call OrderVariable("CLNTAECODE")  %>) >AE Code<br>經紀編號</a></td>
			   <td width="12%"><a href=javascript:ordersubmit(<% call OrderVariable("CLNTAENAME")  %>) >AE Name<br>經紀名稱</td>
			   <td width="12%"><a href=javascript:ordersubmit(<% call OrderVariable("PORTFILIO")  %>) >Portfilio Mkt Value (HKD)<br>組合總值(港元)</a></td> 
			   <td width="12%"><a href=javascript:ordersubmit(<% call OrderVariable("BALANCE")  %>) >A/C Bal (Due Amt)</a></td> 
			   <td width="12%"><a href=javascript:ordersubmit(<% call OrderVariable("MARGINPERCENT")  %>) >Margin %<br>按倉比率</a></td> 
			   <td width="12%"><a href=javascript:ordersubmit(<% call OrderVariable("LOSS")  %>) >Loss<br>損失</a></td> 
			   <td width="12%"><a href=javascript:ordersubmit(<% call OrderVariable("MARGINVALUE")  %>) >Margin Value<br>按倉價值</a></td> 
			   <td width="12%"><a href=javascript:ordersubmit(<% call OrderVariable("ACCUREDINT")  %>) >Accrued Int</a></td> 
			</tr>
			
			
<tr bgcolor="#ADF3B6" align="center">
			<td><% call OrderImage("CLNTCODE")  %></td>
			<td><% call OrderImage("CLNTNAME")  %></td>
			<td><% call OrderImage("CCY")  %></td>
			<td><% call OrderImage("CURRENCY")  %></td>
			<td><% call OrderImage("CLNTAECODE")  %></td>
			<td><% call OrderImage("CLNTAENAME")  %></td>
			<td><% call OrderImage("PORTFILIO")  %></td>
			<td><% call OrderImage("BALANCE")  %></td>
			<td><% call OrderImage("MARGINPERCENT")  %></td>
			<td><% call OrderImage("LOSS")  %></td>
			<td><% call OrderImage("MARGINVALUE")  %></td>
			<td><% call OrderImage("ACCUREDINT")  %></td>

</tr>
	<%
		do while (  Not rs1.EOF)

			
	%>
		
		
		<tr bgcolor="#FFFFCC"> 
			  <td width="10%"><a href="ClientSummaryListDetail.asp?DisplayFirst=<%=rs1("ClntCode") %>&ClientFrom=<%=Search_ClientFrom%>&ClientTo=<%=Search_ClientTo%>&AEFrom=<%=Search_AEFrom%>&AETo=<%=Search_AETo%>&SDay=<%=Search_From_Day%>&SMonth=<%=Search_From_Month%>&SYear=<%=Search_From_Year%>&sid=<%=SessionID%>" target=_blank><%=rs1("ClntCode") %></a><img border=0 src='images/tel.gif' onClick="PopupClientContact('<%=rs1("ClntCode") %>')"></img></td>
			  <td width="18%"><%=rs1("ClntName") %></td>
		   <td width="5%"><%=rs1("CCY") %></td>
		   <td width="10%" style="white-space: nowrap"><%=formatnumber(rs1("Currency"),2) %></td>
		   <td width="5%"><%=rs1("ClntAECode") %></td>
		   <td width="12%"><%=rs1("ClntAEName") %></td>
		   <td width="12%" style="white-space: nowrap"><%=formatnumber(rs1("Portfilio"),2) %></td> 
		   <td width="12%" style="white-space: nowrap"><%=formatnumber(rs1("Balance"),2) %></td> 
		   <td width="12%" style="white-space: nowrap"><%=formatnumber(rs1("Marginpercent"),2) %></td> 
		   <td width="12%">　</td> 
		   <td width="12%" style="white-space: nowrap"><%=formatnumber(rs1("MarginValue"),2) %></td> 
		   <td width="12%" style="white-space: nowrap"><%=formatnumber(rs1("AccuredInt"),2) %></td> 
		</tr>
	<%

		rs1.movenext
	loop

	%>
		</table>
		<br>
		<br>
		


</div>
		
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
		<a href=javascript:pagesubmit(1) style='cursor:hand'>First</a>

	<%
	' Prev button
	If iPageCurrent > 1 Then
		%>
		<a href=javascript:pagesubmit(<%= iPageCurrent-1 %>) style='cursor:hand'>Previous</a>
	<% else %>
	Previous
		<%
	End If


	'Next button
	If iPageCurrent < iPageCount Then
		%>
		<a href=javascript:pagesubmit(<%= iPageCurrent+1 %>) style='cursor:hand'>Next</a>
	<% else %>
	Next
		<%
	End If
	%>

	<%
	'Last button
	%>

	<a href=javascript:pagesubmit(<%= iPageCount %>) style='cursor:hand'>Last</a>

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
window.open("ConvertMargin.asp?Search_MinDebitBalance=<%=Search_MinDebitBalance%>&Search_MarginExceedPercent=<%=Search_MarginExceedPercent%>&Search_AccountType=<%=Search_AccountType%>&From_Day=<%=Search_From_Day%>&From_Month=<%=Search_From_Month%>&From_Year=<%=Search_From_Year%>&To_day=<%=Search_To_Day%>&To_Month=<%=Search_To_Month%>&To_Year=<%=Search_To_Year%>&Search_SharedGroupMember=<%=Search_SharedGroupMember%>&Search_SharedGroup=<%=Search_AEGroup%>"); 

}

//-->
</SCRIPT>