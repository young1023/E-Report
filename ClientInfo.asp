<% 
'*********************************************************************************
'NAME       : ClientInfo.asp           
'DESCRIPTION: Client Info
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 090401 Gary Yeung   Prototype
'MODIFIED   : 090421 Roger Wong   Record and page control
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

Title = "Client Information"




'**************
'Initialisation
'**************


Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adCmdText = &H0001

Const adCmdStoredProc    = 4
Const adInteger          = 3
Const adCurrency         = 6
Const adParamInput       = 1
Const adParamOutput      = 2
Const adExecuteNoRecords = 128
Const adVarChar = 200



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
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
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
Dim Search_SharedSelection 
Dim Search_SharedGroup
Dim Search_SharedGroupMember

Search_AEGroup	    = Request.form("GroupID")
Search_ClientFrom   = Request.form("ClientFrom")
Search_ClientTo     = Request.form("ClientTo")
Search_AEFrom       = Request.form("AEFrom")
Search_AETo         = Request.form("AETo")
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


set RsGroupID = server.createobject("adodb.recordset")
RsGroupID.open ("Exec Retrieve_AvailableGroupID ") ,  StrCnn,3,1


%>

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
			<td colspan="2">
<span class="noprint">
			<font color="red">*</font> Denotes a mandatory field</span>
			</td>
		</tr>
		<% End If %>

			<tr>
					<td width="20%">Client Number: (From)<font color="red">*</font> </td> 
					<td width="30%">
					<input name="ClientFrom" type=text value="<%= Search_ClientFrom %>" size="15">
            <img align="top" style="cursor:pointer" onClick="PopupWindow()" src="images/search.gif"> </td>
					
			<td width="20%">Client Number: (To)</td> 
			<td>
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
			| <span class="noprint">
			<input type="radio" name="ShareSelection" value="share2" onClick=""  <%if Search_SharedSelection = "share2" then response.write "checked" end if %>   > Particular AE in the Sales Team

			<select name="SharedGroupMember" class="common">
						<%
								do while (  Not RsSharedGroupMember.EOF)
						%>
								<option value="<%=RsSharedGroupMember("loginname")%>" <% if Search_SharedGroupMember=RsSharedGroupMember("loginname") then response.write "selected" end if %> > <%=RsSharedGroupMember("loginname")%></option>
						<%
								RsSharedGroupMember.movenext
								Loop
						%>
			</select>&nbsp; |
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
			</select>&nbsp;				

			</span>

			</td> 
		</tr>   

<% end if %>  		
	
	
  	<tr> 
      <td>　</td> 
      <td colspan="3"><input type=hidden   value="<%=iPageCurrent%>"   name="page"> 
 	<input type=hidden   name="submitted"> 
	     
&nbsp;<input id="Submit1" type="button" value="Submit" onClick="dosubmit(1);"></td>
    </tr> 
    		

    </table>
 </form>   
    
<br>

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

	set Rs1 = server.createobject("adodb.Recordset")

	'rs1.CursorLocation=3
	
	Search_Statement ="Daily"

 'Rs return 2 value
 '1) Total number of matched client
 '2) all records for targeted client
 'response.write ("Exec Retrieve_ClientInformation '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '"&iPageCurrent&"', '1' ")
 'Rs1.open ("Exec Retrieve_ClientInformation '0', '9', '', '', '', '"&iPageCurrent&"', '1' "),  StrCnn,3,1
' Rs1.open ("Exec Retrieve_ClientInformation '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '"&iPageCurrent&"', '1' "),  StrCnn,3,1

 'Rs1.open ("Exec AA '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '"&iPageCurrent&"', '1' "),  StrCnn,3,1
  
Select Case  Search_SharedSelection
	case "share2"
		'shared group member
		'	response.write  ("Exec Retrieve_ClientInformation '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '',  '"&iPageCurrent&"', '1' ")
			Rs1.open ("Exec Retrieve_ClientInformation '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '',  '"&iPageCurrent&"', '1' "),  StrCnn,3,1



	case "share3"
			'response.write  ("Exec Retrieve_ClientInformation '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&iPageCurrent&"', '1' ")

			Rs1.open ("Exec Retrieve_ClientInformation '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&iPageCurrent&"', '1' "),  StrCnn,3,1
			

	
	case else
		'normal
			'response.write   ("Exec Retrieve_ClientInformation '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&iPageCurrent&"', '1' ")
			Rs1.open ("Exec Retrieve_ClientInformation '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&iPageCurrent&"', '1' "),  StrCnn,3,1
	
	
	'Rs1.open ("Retrieve_ClientInformation '0', '9', '', '', '1', '', '', '1', '1' "),  StrCnn,3,1		

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
</span>
<DIV align=center>
		<%
'**********
' Start of page navigation 
'**********
%>
  <TABLE border=0 cellPadding=0 cellSpacing=0 height=100% width=99%>

 <tr> 
 <td align="right" height="28" class="NavaMenu" >

<% If Trim(PrintAllowed) = 1 then %> 
							<a href="javascript:window.print()">Friendly Print</a>
						<% end if %>
&nbsp;&nbsp;
<%
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


      <table width="99%" border="0" class="normal" cellspacing="1" cellpadding="4">
<tr bgcolor="#FFFFCC">
      <td align="center">客戶資料<br><u>Client Information</u></td>

</tr>

</table>
<br>


<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">


			<tr>
			<td width="34%" bgcolor="#ADF3B6">Client Number<br>客戶號碼</td>
			<td width="65%" bgcolor="#FFFFFF">
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "1000" then
					response.write rs1("ClntCode")
					response.write "<BR>"
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "1000") then
						exit do
					end if
				end if
			loop
		%>
			
			</td>
			</tr>

<tr>
      <td width="34%" bgcolor="#ADF3B6">Client Full Name<br>客戶名稱</td>
      <td width="65%" bgcolor="#FFFFFF">
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "1100" then
					response.write rs1("FieldValue")
					response.write "<BR>"
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "1100") then
						exit do
					end if
				end if
			loop
		%>
      </td>
   </tr>
<tr>
      <td width="34%" bgcolor="#ADF3B6">Other Name<br>其他名稱</td>
      <td width="65%" bgcolor="#FFFFFF">		
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "1200" then
					response.write rs1("FieldValue")
					response.write "<BR>"
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "1200") then
						exit do
					end if
				end if
			loop
		%>
				  </td>
   </tr>

<tr bgcolor="#ADF3B6">
      <td width="34%">Cheque Name<br>支票名稱</td>
      <td width="65%" bgcolor="#FFFFFF">		
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "1300" then
					response.write rs1("FieldValue")
					response.write "<BR>"
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "1300") then
						exit do
					end if
				end if
			loop
		%>
				  </td>
</tr>

<tr bgcolor="#ADF3B6">
      <td width="34%">Account Type<br>客戶類型</td>
      <td width="65%" bgcolor="#FFFFFF">		
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "1400" then
					response.write rs1("FieldValue")
					response.write "<BR>"
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "1400") then
						exit do
					end if
				end if
			loop
		%>
				  </td>
</tr>

<tr bgcolor="#ADF3B6">
      <td width="34%">Margin Limit<br>按倉限額</td>
      <td width="65%" bgcolor="#FFFFFF">		
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "1500" then
					if isnumeric(rs1("FieldValue")) then
						response.write formatnumber(rs1("FieldValue"),2)
						response.write "<BR>"
					end if
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "1500") then
						exit do
					end if
				end if
			loop
		%>
				  </td>
</tr>

<tr bgcolor="#ADF3B6">
      <td width="34%">Transaction Limit<br>交易限額</td>
      <td width="65%" bgcolor="#FFFFFF">		
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "1600" then
					if isnumeric(rs1("FieldValue")) then
						response.write formatnumber(rs1("FieldValue"),2)
						response.write "<BR>"
					end if
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "1600") then
						exit do
					end if
				end if
			loop
		%>
				  </td>
</tr>

<tr bgcolor="#ADF3B6">
      <td width="34%">Daily Limit<br>每日限額</td>
      <td width="65%" bgcolor="#FFFFFF">		
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "1700" then
					if isnumeric(rs1("FieldValue")) then
						response.write rs1("FieldValue")
						response.write "<BR>"
					end if
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "1700") then
						exit do
					end if
				end if
			loop
		%>
				  </td>
</tr>

<tr bgcolor="#FFFFCC"> 
      <td width="34%" bgcolor="#ADF3B6">Brokerage rate setup for different 
		markets<br>不同市場之經紀佣金設定</td>
      <td width="65%" bgcolor="#FFFFFF">		
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "1800" then
					response.write CDbl(rs1("FieldValue"))
					response.write "<BR>"
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "1800") then
						exit do
					end if
				end if
			loop
		%>
				  </td>
</tr>

<tr bgcolor="#FFFFCC"> 
      <td width="34%" bgcolor="#ADF3B6">Client contact<br>客戶聯絡資料</td>
      <td width="65%" bgcolor="#FFFFFF">		
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "1900" then
					response.write rs1("FieldValue")
					response.write "<BR>"
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "1900") then
						exit do
					end if
				end if
			loop
		%>
				  </td>
</tr>

<tr bgcolor="#FFFFCC"> 
      <td width="34%" bgcolor="#ADF3B6">Client rebate setup<br>客戶回扣設定</td>
      <td width="65%" bgcolor="#FFFFFF">		
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "2000" then
					response.write CDbl(rs1("FieldValue"))
					response.write "<BR>"
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "2000") then
						exit do
					end if
				end if
			loop
		%>
				  </td>
</tr>

<tr bgcolor="#FFFFCC"> 
      <td width="34%" bgcolor="#ADF3B6">Client's debit and credit interest rate 
		setup<br>客戶結欠及結存之設定</td>
      <td width="65%" bgcolor="#FFFFFF">		
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "2100" then
					response.write CDbl(rs1("FieldValue"))
					response.write "<BR>"
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "2100") then
						exit do
					end if
				end if
			loop
		%>
				  </td>
</tr>

<tr bgcolor="#FFFFCC"> 
      <td width="34%" bgcolor="#ADF3B6">Client's Settlement Method<br>客戶結算方法</td>
      <td width="65%" bgcolor="#FFFFFF">		
      <% 
			do while (  Not rs1.EOF) 
				if rs1("sectioncode") = "2200" then
					response.write rs1("FieldValue")
					response.write "<BR>"
					rs1.movenext
				end if	
				if (rs1.eof) then
					exit do
				else
					if (rs1("sectioncode") <> "2200") then
						exit do
					end if
				end if
			loop
		%>
				  </td>
</tr>

</table>
<br>





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

</td></tr></table>
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