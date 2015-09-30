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

Server.ScriptTimeout = 1800000

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

set RsSettle = server.createobject("adodb.recordset")



strURL = Request.ServerVariables("URL") ' Retreive the URL of this page from Server Variables
%>





<html>
<head>


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
  
			document.fm1.submitted.value=what;
			document.fm1.action="<%= strURL %>?sid=<%=SessionID%>";
		    document.fm1.submit();
	
}



function hidediv() { 
document.getElementById('hideShow').style.visibility = 'hidden'; 
} 

function showdiv() { 
document.getElementById('hideShow').style.visibility = 'visible'; 
} 


function RemoveContent(d) {

document.getElementById(d).style.display = "none";

}

function InsertContent(d) {

document.getElementById(d).style.display = "";

}

//-->
</script>

</head>
<body leftmargin="0" topmargin="0" OnLoad="document.fm1.submitted.value=0;"  onkeypress="return disableCtrlKeyCombination(event);" onkeydown="return disableCtrlKeyCombination(event);"  >




<span class="noprint">


  


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
Dim PrintAllowed

SharedNumber            = Trim(Request("SharedNumber"))
SharedName              = Trim(Request("SharedName"))
Search_ClientFrom1      = session("ClientFrom")                           
Search_ClientTo1        = session("ClientTo") 


Search_AEGroup	    = Request("GroupID")
Search_ClientFrom   = Trim(Request("DisplayFirst"))
Search_ClientTo     = Trim(Request("DisplayFirst"))
Search_AEFrom       = Request("AEFrom")
Search_AETo         = Request("AETo")
Search_Daily_Day    = Request("SDay")
Search_Daily_Month  = Request("SMonth")
Search_Daily_Year   = Request("SYear")
Search_LedgerBalanceType   = "ALL"
Search_LedgerBalance   = Request("LedgerBalance")

Search_IncludeMarginAccount   = "ALL"
Search_Market           = "ALL"
Search_Instrument       = ""
Search_SharedSelection      =  Request("ShareSelection")	
Search_SharedGroup  = Request("SharedGroup")
Search_SharedGroupMember  = Request("SharedGroupMember")

UserCurrentClient = Request("CurrentClient")



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


If request.form("submitted")  = 1 Then


    Search_ClientFrom    = SharedNumber
    Search_ClientTo      = SharedNumber
    SharedName           = SharedNumber
	Search_Daily_Day      = session("Search_From_Day")  
	Search_Daily_Month    = session("Search_From_Month")
	Search_Daily_Year     = session("Search_From_Year")
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
	Search_Daily_Day      = session("Search_From_Day")  
	Search_Daily_Month    = session("Search_From_Month")
	Search_Daily_Year     = session("Search_From_Year")
	Search_AEGroup	     = session("GroupID")
	Search_AEFrom        = session("AEFrom")
	Search_AETo          = session("AETo")
	Search_Market        = session("Search_Market")
    Search_Instrument    = session("Search_Instrument") 
   
End If

If request("submitted")  = 3 Then


    Search_ClientFrom    = request("ClientFrom")
    Search_ClientTo      = request("ClientTo")
	Search_Daily_Day     = session("Search_From_Day")  
	Search_Daily_Month   = session("Search_From_Month")
	Search_Daily_Year    = session("Search_From_Year")
	Search_AEGroup	     = session("GroupID")
	Search_AEFrom        = session("AEFrom")
	Search_AETo          = session("AETo")
	Search_Market        = session("Search_Market")
    Search_Instrument    = session("Search_Instrument") 
   
End If
                                
                                
                                
set RsGroupID = server.createobject("adodb.recordset")
RsGroupID.open ("Exec Retrieve_AvailableGroupID ") ,  Conn,3,1
                                

  '  Query for Client Number and Client Name

    set Rs = server.createobject("adodb.recordset")
     
    'response.write  ("Exec Retrieve_ClientSummary_ClientCode '"&Search_ClientFrom1&"', '"&Search_ClientTo1&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") 
 
    Rs.open ("Exec Retrieve_ClientSummary_ClientCode '"&Search_ClientFrom1&"', '"&Search_ClientTo1&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") ,  Conn,3,1

    set RsN = server.createobject("adodb.recordset")
     
    'response.write  ("Exec Retrieve_ClientSummary_ClientName '"&Search_ClientFrom1&"', '"&Search_ClientTo1&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") 
 
    RsN.open ("Exec Retrieve_ClientSummary_ClientName '"&Search_ClientFrom1&"', '"&Search_ClientTo1&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_Market&"','"&Search_Instrument&"' ") ,  Conn,3,1

    
    

                                
%>                              
                                
<%                              
'*****************************************************************
' Start of form
'*****************************************************************
%>
         


<form name="fm1" method="post" action="<%= strURL %>">
	
<table width="99%" border="0" class="normal"  cellspacing="2" cellpadding="4">
		<tr> 
			<td colspan="4">
			&nbsp;
			<select name="SharedNumber" class="common" onChange="dosubmit(1)">
              <option value=""> -- Please select the client code -- </option>
    		   
                       <%      Rs.MoveFirst 

                               Do While Not Rs.EoF %>

               <option value="<%=Rs("Clnt")%>" <%If Trim(SharedNumber)=Trim(Rs("Clnt")) Then%>Selected<%End If%>>  <%=Rs("Clnt")%> |  <%=Rs("ClntName")%>  </option>
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

               <option value="<%=RsN("Clnt")%>" <%If Trim(SharedName)=Trim(RsN("Clnt")) Then%>Selected<%End If%>>    <%=RsN("ClntName")%> | <%=RsN("Clnt")%>  </option>
                       <%
								RsN.Movenext
								Loop
						%>
              
			</select>	
					</td>
             <td align="left"><a href="ClientSummaryListDetail.asp?submitted=3&ClientFrom=<%=Search_ClientFrom1%>&ClientTo=<%=Search_ClientTo1%>&AEFrom=<%=Search_AEFrom%>&AETo=<%=Search_AETo%>&Instrument=<%=Search_Instrument%>&Market=<%=Search_Market%>&FromDay=<%=Search_From_Day%>&FromMonth=<%=Search_From_Month%>&FromYear=<%=Search_From_Year%>&Today=<%=Search_To_Day%>&ToMonth=<%=Search_To_Month%>&ToYear=<%=Search_To_Year%>&Search_Order=ClientCode&Search_Direction=ASC&PrintAllowed=1&sid=<%=SessionID%>">Show All</a>
             </td>
		</tr> 
</table>

<br>	

		


			<input type=hidden   value="<%=iPageCurrent%>"   name="page"> 
 	<input type=hidden   name="submitted"> 


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
 
            PrintAllowed =  0  

          ' Permission for Printing 
          '************************
         set pRs = server.createobject("adodb.recordset")

		 pRs.open ("Exec Check_PrintPermission '"&Session("MemberID")&"', '"&Title&"' ") , Conn,3,1

  
           iRecordCount = pRs(0)

           If iRecordCount > 0 then
 
           PrintAllowed =  1     

           End if    
        
%>	
	
<table width="99%" border="0" class="normal"  cellspacing="1" cellpadding="2">



<tr bgcolor=<%=TableBGColor %>> 
<td  width="20%"> </td>
      <td align="center">«È¤áÁ`µ²<br><u>Client Detail Information</u></td> 
        <td align="right" width="20%">
           <span class="noprint">
							<%
                             if PrintAllowed = 1 then %>  
							<a href="javascript:window.print()">Friendly Print</a>&nbsp;&nbsp;
							<% End If %>
			
          </span>      	
			</td>
</tr>
</table>
<br>
<%



'**********
' If passing arguments
'**********
	
	'dim Rs1 as adodb.recordset 
	
	'Conn.open myDSN 
	
	set Rs0 = server.createobject("adodb.recordset")
	set Rs1 = server.createobject("adodb.recordset")
	set Rs2 = server.createobject("adodb.recordset")
	set Rs3 = server.createobject("adodb.recordset")

	
	'rs1.CursorLocation=3
	
Search_Statement ="Daily"
dim Search_CurrentRecord
Search_CurrentRecord=1


 'Rs return 2 value
 '1) Total number of matched client
 '2) all records for targeted client
	

	Select Case  Search_SharedSelection
	case "share2"
		'shared group member
			'response.write  ("Exec Retrieve_ClientSummaryListDetail '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1'  ") 
response.write "This is 1" & session("shell_power") & "," & Search_CurrentRecord
			Rs0.open ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_LedgerBalance&"','"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1'  ") ,  Conn,3,1
	case "share3"
response.write "This is 2" & session("shell_power") & "," & Search_CurrentRecord
			'response.write  ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_LedgerBalance&"','"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1' ")

			Rs0.open ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_LedgerBalance&"','"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1' ") ,  Conn,3,1

	case else
response.write "This is 3" & session("shell_power") & "," & Search_CurrentRecord
		'normal
		response.write  ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"','"&Search_LedgerBalance&"', '"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1'  ")
			Rs0.open ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"','"&Search_LedgerBalance&"', '"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1'  ") ,  Conn,3,1
		
	end select	
	
		

	'assign total number of pages
	iPageCount = rs0(0)


  if iPageCount <= 0 then

		If Err.Number <> 0 then
			
			'SQL connection error handler
			'response.write  "<table><tr><td class='RedClr'>" & MSG_BUSY & "<br></td></tr></table>"
			
		else
			'no record found
			response.write ("No record found")
				
		End If
	else
		'record found
		rs0 = nothing
		

	End If

'--------------------------


'for loop for every client code


for Search_CurrentRecord = 1 to iPageCount 
	

	Select Case  Search_SharedSelection
	case "share2"
		'shared group member
			'response.write  ("Exec Retrieve_ClientSummaryListDetail '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1'  ") 
			Rs1.open ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_LedgerBalance&"','"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1'  ") ,  Conn,3,1
	case "share3"
			'response.write  ("Exec Retrieve_ClientSummaryListDetail '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1'  ")

			Rs1.open ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"', '"&Search_LedgerBalance&"','"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1' ") ,  Conn,3,1

	case else
		'normal
			response.write  ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"','"&Search_LedgerBalance&"', '"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1'  ")
			Rs1.open ("Exec Retrieve_ClientSummary '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"','"&Search_Market&"','"&Search_Instrument&"','"&Search_LedgerBalanceType&"','"&Search_LedgerBalance&"', '"&Search_IncludeMarginAccount&"', '"&Search_CurrentRecord&"', '1'  ") ,  Conn,3,1
		
	end select	
	
		

	'assign total number of pages
	iPageCount = rs1(0)


  if iPageCount <= 0 then

		If Err.Number <> 0 then
			
			'SQL connection error handler
			'response.write  "<table><tr><td class='RedClr'>" & MSG_BUSY & "<br></td></tr></table>"
			
		else
			'no record found
			'response.write ("No record found")
				
		End If
	else
		'record found
		
		'move to next recordset
  	Set rs1 = rs1.NextRecordset() 
  	
  	
		 CurrentClientCode = rs1("clnt")
		If UserCurrentClient = CurrentClientCode then
		      TableBGColor = "#FFFFAA" 
		 Else
		      TableBGColor = "#FFFFCC"
		
		end if



%>

 




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

   
				<% If Trim(rs1("Clnt")) = Trim(Request("DisplayFirst")) Then %>
				       <a name = "DisplayFirst">
				<% End If %>  


                                          			
						<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
						
								<tr >
										<td bgcolor="#ADF3B6" width="10%" height="33"><b>Client: </b></td> 
										<td bgcolor=<%=TableBGColor %> width="20%" height="33"><b><%=rs1("clnt") %></b></td>
										<td bgcolor=<%=TableBGColor %> width="40%" height="33"><b><%=rs1("ClntName")%></b></td>
										<td bgcolor=<%=TableBGColor %> width="30%" height="33"><b>
										<%
												RsSettle.open ("Exec Retrieve_ClientSummaryDetail_SettleMethod '"&CurrentClientCode&"'  ") ,  Conn,3,1
												
												do while (  Not RsSettle.EOF)
												response.write RsSettle("FieldValue") & "<br>"
												RsSettle.movenext
												Loop
												
												RsSettle.Close
										%>										
										</b></td> 
										
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
					case "MG" exit do
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
			if (rs1("sectioncode") = "SM" or rs1("sectioncode") = "CM") then %>
						<br>			
						<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
								
									<tr bgcolor=<%=TableBGColor %>> 
												<td width="95%" colspan="9" >Movement</td>
									</tr>
									
									
									<tr bgcolor="#ADF3B6"> 
												<td width="5%">Date</td>
												<td width="6%">Value</td>
												<td width="6%">Ref No</td>
												<td width="15%">Description</td>
												<td width="10%" align="right">Quantity</td>
												<td width="6%">CCY</td>
												<td width="4%" align="right">Price</td>
												<td width="20%" align="right">Amount</td> 
									</tr>
									
									<% do while (  Not rs1.EOF)
									
												if (rs1("sectioncode") = "SM" or rs1("sectioncode") = "CM")  then
												%>
												
															<tr bgcolor=<%=TableBGColor %>> 
																		<td width="5%"><%= rs1("tradedate") %></td>
																		<td width="6%"><%= rs1("settledate") %> </td>
																		<td width="6%"><%= rs1("RefNo") %> </td>
																		<td width="50%">
																			<%= rs1("buysell") + " "%> 
																			<%= rs1("Instrument") + " "%> 
																			<%= rs1("comment") + " "%> 
																		</td>
																		<td width="5%" align="right"><% if formatnumber(rs1("Quantity"),0) <> 0 then response.write formatnumber(rs1("Quantity"),0) end if %> </td> 
																		<td width="5%"><%= rs1("ccy") %> </td> 
																		<td width="5%" align="right"><% if rs1("price") <> 0 then response.write rs1("price") end if %> </td> 
																		<td width="5%" align="right"><% if formatnumber(rs1("Amount"),2) <> 0 then response.write formatnumber(rs1("Amount"),2) end if %> </td> 

															</tr>
															<%
												end if	
												rs1.movenext
												
												if not rs1.eof then
														if (rs1("sectioncode") <> "SM" AND rs1("sectioncode") <> "CM") then 
																exit do
														end if
												end if
									loop
									
									
									
									%>
						</table>
						<br>
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
				
							
									<tr bgcolor=<%=TableBGColor %>> 
												<td width="95%" colspan="11">Stock </td>
									</tr>
									
									<tr bgcolor="#ADF3B6">
												<td width="5%">Mkt</td>
												<td width="5%">StkCode</td>
												<td width="20%">StkName</td>
												<td width="5%" align="right">Ledger Balance</td>
		
												<td width="5%" align="right">Bonus & UnderReg</td>
												<td width="5%">CCY</td>
												<td width="5%" align="right">Price</td> 
												<td width="5%" align="right">ExRate</td> 
												<td width="5%" align="right">Portfolio value (HKD)</td> 
																															
													<%if AccountType = "MRGN" then %>
															
														<td width="5%">Mrgn %</td>
														<td width="5%">Mrgnable Value</td> 
													<%	end if %>
									</tr>
									
									<% 
									do while (  Not rs1.EOF)
									
												if rs1("sectioncode") = "SP" then
												%>

															<tr bgcolor=<%=TableBGColor %>> 
																		<td ><%= rs1("Market") %></td>
																		<td><%= rs1("Instrument") %></td>
																		<td><% response.write rs1("InstrumentDesc")  %> </td>

																		<td align="right">
																			<%
																				response.write formatnumber(rs1("EndingBalance"),0) 
																				'stock on hold positive
																				'response.write clng(rs1("orfee1"))
																				if (clng(rs1("orfee1")) <> 0  ) then
																					response.write "<br>*" & formatnumber(clng(rs1("orfee2")),0)
																				end if
																				
														
																			%> 
																		</td>

																		<td align="right">0 </td>
																		<td ><%= rs1("CCY") %> </td> 
																		<td align="right"><%= rs1("price") %> </td> 
																		<td align="right"><%= formatnumber(rs1("MarginLimit"),2) %> </td> 
																		<td align="right"><%= formatnumber( int(cDbl(rs1("StockPortValue"))*100)/100  ,2) %> </td> 


																		<%if AccountType = "MRGN" then %>
																				
																			<td align="right"><%=rs1("MarginPercent")%></td>
																			<td align="right"><%=formatnumber(rs1("MarginValue"),2)%></td> 
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
									<tr bgcolor=<%=TableBGColor %>>
												<td colspan="<% if AccountType = "MRGN" then response.write "8" else response.write "8" %> " align="right">
													<b>Portfolio Value: </b></td>
												<td align="right"><b><%=formatnumber(TotalPortfolioValue) %></b></td>

													<% if AccountType = "MRGN" then %>
														<td></td>
														<TD align="right"><b><% =formatnumber(TotalMarginValue) %> </b></TD>
														
													<% end if %>
													
												
												
									</tr>
									
						</table>

		</table>
			<%
			
			end if 
end if

%>


<%
	Dim Total_CashValue
	Dim Total_NetValue
	dim shownetvalue

shownetvalue=0
	Total_CashValue =0
	Total_NetValue = 0
Total_NetValue = TotalPortfolioValue
%>				


<% if (  Not rs1.EOF)  then
			if rs1("sectioncode") = "CB" then %>

<br>

<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
<tr bgcolor=<%=TableBGColor %>> 
			<td width="95%" colspan="12" >Cash Balance</td>
</tr>
<tr bgcolor="#ADF3B6">
   <td rowspan="2">CCY</td>
   <td colspan="4" align="center">Available Balance</td>
   <td colspan="4" align="center">Settlement Amount</td> 
   <td rowspan="2" align="right">Ledger balance</td>
   <td rowspan="2" align="right">Balance (HKD)</td>
   <td rowspan="2" align="right">Int Accrued</td>
</tr>

<tr bgcolor="#ADF3B6">
   <td align="right">T Day</td>
   <td align="right">T + 1</td>
   <td align="right">T + 2</td>
   <td align="right">T + 3</td> 
   <td align="right">T Day</td>
   <td align="right">T + 1</td>
   <td align="right">T + 2</td>
   <td align="right">T + 3</td> 
</tr>


				<% do while (  Not rs1.EOF)
		
				if rs1("sectioncode") = "CB" then
					exit do
				end if
				rs1.movenext
			loop
			
			if not rs1.eof then
				if rs1("sectioncode") = "CB" then 
		%>




		<% 

				TotalPortfolioValue
				do while (  Not rs1.EOF)
			
				if rs1("sectioncode") = "CB" then
		%>
		

<tr bgcolor=<%=TableBGColor %>> 
   <td ><%= rs1("ccy") %></td>
   <td align="right"><%= formatnumber(rs1("ORFee1"),2) %> </td>
   <td align="right"><%= formatnumber(rs1("ORFee2"),2) %> </td>
   <td align="right"><%= formatnumber(rs1("orfee3"),2) %> </td> 
   <td align="right"><%= formatnumber(rs1("orfee4"),2) %> </td> 
   <td align="right"><%= formatnumber(rs1("orfee7"),2)  %>  </td> 
   <td align="right"><%= formatnumber(rs1("orfee8"),2) %> </td> 
   <td align="right"><%= formatnumber(rs1("orfee9"),2) %> </td> 
   <td align="right"><%= formatnumber(rs1("orfee10"),2) %></td>
   <td align="right"><%= formatnumber(rs1("EndingBalance"),2) %> </td>
   <td align="right"><%= formatnumber(rs1("StockPortValue"),2) %> </td>
   <td align="right"><%= formatnumber(rs1("MTDDebitInterest"),2) %></td>

</tr>
<%
					Total_CashValue = Total_CashValue + cDbl(rs1("EndingBalance"))*cDbl(rs1("MarginLimit"))
					end if	
				rs1.movenext
				Total_NetValue = TotalPortfolioValue + Total_CashValue
				
				
				
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
	end if %>
	

	<tr bgcolor=<%=TableBGColor %>> 
		<td colspan="10" align=right> <b>Cash Value</b></td>
		<td align="right"><b><% = formatnumber(Total_CashValue,2)%></b> </td>
		<td  > </td>
	</tr>
	
	<tr bgcolor=<%=TableBGColor %>> 
		<td colspan="10" align=right ><b>Net Value</b></td>
		<td align=right><b><u><% = formatnumber(Total_NetValue,2)%></u></b> </td>
		<td width = "10%"></td>

	</tr>
	
</table>

<% 
	shownetvalue = 1
	end if 
		end if
%>

<% 
	if shownetvalue <> 1 then %>
<br>
<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

	<tr bgcolor=<%=TableBGColor %>> 
		<td align=right width = "75%"><b>Net Value</b></td>
		<td align=right><b><u><% = formatnumber(Total_NetValue,2)%></u></b> </td>
		<td width = "10%"></td>

	</tr>
</table>

<% end if %>

<%



     Rs3.open ("Exec Retrieve_ClientSummaryListDetail_clientinformation '"&CurrentClientCode&"',  '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"' ") ,  Conn,3,1
     'response.write ("Exec Retrieve_ClientSummaryListDetail_clientinformation '"&CurrentClientCode&"',  '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"' ") 


if Not Rs3.EoF  then %>




	<br>
				<% if AccountType = "MRGN" then %>
			

								
									<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
										<tr bgcolor=<%=TableBGColor %>> 
													<td width="95%" colspan="11" >Margin Client Only</td>
										</tr>								
										<tr >
											<td bgcolor="#ADF3B6" width="33%" align=right>  Margin Limit  </td> 

											<td bgcolor="#ADF3B6" width="33%" align=right>  Available / (Margin Call) </td>
								
											<td bgcolor="#ADF3B6" width="33%" align=right>%used </td> 

										</tr>
										
										<tr >
											<td align=right bgcolor=<%=TableBGColor %> >  <%=formatnumber(rs3("MarginLimit"),0)  %>  </td> 					
											<td align=right bgcolor=<%=TableBGColor %> >  <%=formatnumber(rs3("MarginCall"),2) %></td>
											<td align=right bgcolor=<%=TableBGColor %>> 
											<% 
                                                                          'response.write Total_CashValue & "<br>"
                                                                          'response.write rs3("MarginLimit") & "<br>"



												if (Total_CashValue < 0) then
													if (TotalMarginValue > cdbl(rs3("MarginLimit"))) then
														response.write formatnumber(abs(Total_CashValue / cdbl(rs3("MarginLimit"))*100),2) &"<br>"

													else
														response.write formatnumber(abs(Total_CashValue / TotalMarginValue*100),2)  

													end if
												else
														response.write formatnumber(0,2)  
					
												end if
											%> %</td> 

										</tr>		
								
									</table>
			
				<% else	%>
									<table  width="20%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
										<tr bgcolor=<%=TableBGColor %>> 
													<td width="95%" colspan="11" >Cash or Custodian Client</td>
										</tr>				
														
										<tr >

											<td bgcolor="#ADF3B6" width="33%" align=right>  Trading Limit  </td> 


										</tr>
										
										<tr >
											<td align=right bgcolor=<%=TableBGColor %> > &nbsp  <%=formatnumber(rs3("TransactionLimit"),0) %></td> 					
										</tr>										
									</table>


				<% end if %>


	<% end if %>

<br>
<hr color="red" width="97%">
<br><br>

<%
	end if 'record found if statement
'end if   'having client number if statement
%>

<% If Trim(rs1("Clnt")) = Trim(Request("DisplayFirst")) Then %>
    </a>
<!-- <div id="lower" style="font-size: 24px; display: none;"> -->
<% End If %> 

<%

rs1.close
rs3.close
'for loop for every client code
next

%>

</div>



	

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
 Set Rs2 = Nothing
 Set Rs3 = Nothing
 set RsSettle = Nothing
 Conn.Close

 'pass all selection to session 
session("GroupID")                 =  Search_AEGroup	               
session("ClientFrom")              =  Search_ClientFrom1              
session("ClientTo")                =  Search_ClientTo1                
session("AEFrom")                  =  Search_AEFrom                  
session("AETo")                    =  Search_AETo                    
session("Search_From_Day")         =  Search_Daily_Day
session("Search_From_Month")       =  Search_Daily_Month
session("Search_From_Year")        =  Search_Daily_Year
session("Search_Market")           =  Search_Market
session("Search_Instrument")       =  Search_Instrument

 Server.ScriptTimeout = 180


 Set Conn = Nothing
%>