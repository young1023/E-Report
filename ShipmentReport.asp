<!--#include file="include/SessionHandler.inc.asp" -->

<%

Search_From_Month       = Request.form("FromMonth")
Search_From_Year        = Request.form("FromYear")
Search_Market           = Request.form("Market")
Search_Instrument       = Request.form("Instrument")
Search_DepotCode        = Request.form("DepotCode")
Search_ISIN             = Request.form("ISIN")
Search_Sedol            = Request.form("Sedol")
Search_Match            = Request.form("Match")


' Retrieve page to show or default to the first
If Request.form("page") = "" Then
	iPageCurrent = 1
Else
	iPageCurrent = Clng(Request.form("page"))
End If


If Search_From_Month = "" Then
     Search_From_Month = month(now()) 
 End If
 
      If len(Search_From_Month) = 1 Then
           Search_From_Month = "0" & Search_From_Month
      End if

If Search_From_Year = "" Then      
	Search_From_Year = year(Now())  - 1
End If

' Market pull down menu
set RsMarket = server.createobject("adodb.recordset")
RsMarket.open ("Exec Retrieve_AvailableMarket ") ,  StrCnn,3,1


On Error resume Next



if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

Title = "Stock Reconciliation Report"

if session("shell_power")="" then
  response.redirect "Default.asp"
end if

%>

<html>
<head>
	
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>Report</title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<link rel="stylesheet" type="text/css" media="print" href="include/print.css" />

<SCRIPT language=JavaScript>
<!--

function dosubmit(what){
 
 document.fm1.action="ShipmentReport.asp?sid=<%=SessionID%>";
 document.fm1.page.value=what;
 document.fm1.submit();
	
}


function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="ShipmentReport.asp?sid=<%=SessionID%>"
document.fm1.submit();
}

function findenum()
{
document.fm1.pageid.value=1;
document.fm1.action="ShipmentReport.asp?sid=<%=SessionID%>"
document.fm1.submit();
}
//-->
</script>

</head>

<body leftmargin="0" topmargin="0">


<span class="noprint">
<!-- #include file ="include/Master.inc.asp" -->
</span>

<div id="Content">


<form name="fm1" method="post" action="">
  <table width="97%" border="0" class="normal">
 <tr> 
      <td width="20%">Date:</td> 
      <td>
			<select name="FromMonth" class="common">            	
					<option value="01" <% if Search_From_Month=01 then response.write "selected"%>>Jan</option>
					<option value="02" <% if Search_From_Month=02 then response.write "selected"%>>Feb</option>
					<option value="03" <% if Search_From_Month=03 then response.write "selected"%>>Mar</option>
					<option value="04" <% if Search_From_Month=04 then response.write "selected"%>>Apr</option>
					<option value="05" <% if Search_From_Month=05 then response.write "selected"%>>May</option>
					<option value="06" <% if Search_From_Month=06 then response.write "selected"%>>Jun</option>
					<option value="07" <% if Search_From_Month=07 then response.write "selected"%>>Jul</option>
					<option value="08" <% if Search_From_Month=08 then response.write "selected"%>>Aug</option>
					<option value="09" <% if Search_From_Month=09 then response.write "selected"%>>Sep</option>
					<option value="10" <% if Search_From_Month=10 then response.write "selected"%>>Oct</option>
					<option value="11" <% if Search_From_Month=11 then response.write "selected"%>>Nov</option>
					<option value="12" <% if Search_From_Month=12 then response.write "selected"%>>Dec</option>
			</select>


			<select name="FromYear" class="common">   
<% 


Year_starting = Year(DateAdd("yyyy", -1 , Now()))
year_ending = Year(Now())

for i=Year_starting to Year_ending
%>			         
			<option value="<%=i%>" <% if clng(i)=clng(Search_From_Year) then response.write "selected"%>><%=i%></option>

<% next %>

			</select> </td>
     
     <td width="20%"></td>
      <td></td>

 	     
	
    
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
  

<%

   

' Start the Queries
' *****************

       set Rs1 = server.createobject("adodb.recordset")
               
    response.write ("Exec Retrieve_InOutReport '"&Search_From_Month&"', '"&Search_From_Year&"' , '"&iPageCurrent&"' ") 
              
	Rs1.open ("Exec Retrieve_InOutReport '"&Search_From_Month&"', '"&Search_From_Year&"' , '"&iPageCurrent&"' ") ,  conn,3,1


	'assign total number of pages
	
	
	
	iRecordCount = Rs1("Tcount")
	
 
  'response.write Rs1("Tcount")

%>   
  
<div id="reportbody1" >

   
<br>


<table width="99%" border="0" class="normal"  cellspacing="1" cellpadding="2">
<span class="noprint">
<tr bgcolor="#FFFFCC"> 
<td  width="20%">�@</td>
      <td align="center">In Out Report</td> 
      <td align="right" width="20%">
      

						
<a href="javascript:window.doConvert()">CSV</a>


					   	
			</td>
</tr>

<tr>

   <td>

<%




  if iRecordCount <= 0 then


		'no record found
		response.write ("No record found")
						
	else

		
        'cal total no of pages
		iPageCount = int(iRecordCount / 1000) + 1
		


		'move to next recordset
  	Set Rs1 = Rs1.NextRecordset() 


%>
</td>
 <td  colspan="2" align="right" >

<%
response.write ( iPageCurrent & " Page of " & iPageCount &"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp" )

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



</td>
</tr>
</table>
<br/>

<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

<tr bgcolor="#ADF3B6" align="center">
      
      <td></td>
      <td>Material Code</td>
      <td>Product Name</td>
      <td>Retailer</td>
      <td>Month/Year</td>
      <td>Sale In Volume</td>
      <td>Sale In Amount</td>
      <td>Sale Out Volume</td>
      <td>Sale Out Amount</td>
      <td>Difference in Volume</td> 
     
        
</tr>
<%		
    i=0
    
    
  do while Not Rs1.EoF

   if Rs1.eof then exit do
 

   i=i+1
		
%>
<tr bgcolor="#FFFFCC"> 
<td>
<%
        Response.write i 
 
%>
</td>

<td><%  = Rs1("Material")  %></td> 

<td>
<%  = Rs1("ProductName") %>
</td>


<td><%  = Rs1("Retailer")  %></td> 

<td><%  = Search_From_Month & "/" & Search_From_Year %></td> 


<td>
<% = FormatNumber(Rs1("SaleInQTY"),0) %>
</td>

<td>

<% = FormatNumber(Rs1("SaleInAmount"),2) %>

</td>

<td>
<% = FormatNumber(Rs1("SaleOutQTY"),0) %>
</td>


<td><% = FormatNumber(Rs1("SaleOutAmount"),2) %></td>

<td ><% = FormatNumber(Rs1("SaleInQTY"),0) - FormatNumber(Rs1("SaleOutQTY"),2) %></td>



</tr>


<%

				
					
				Rs1.movenext
				
		loop
	
End If

%>

                             
</table>
</form>  

<br/>
<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
	<tr bgcolor="#FFFFCC"> 
      <td width="99%" height="18" align="center">End of Report</td>
	</tr>
</table>
                
</div>
              </center>




              </body>

              </html>
              
<%
'*****************************************************************
' Termination
'*****************************************************************

 Rs1.Close
 set Rs1 = Nothing

 Conn.Close
 Set Conn = Nothing

 ' function
  Sub countpage(PageCount,pageid)
  response.write pagecount&"</font> Pages "
	   if PageCount>=1 and PageCount<=10 then
		 for i=1 to PageCount
		   if (pageid-i =0) then
              response.write "<font color=green> "&i&"</font> "
		   else
             response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		   end if
		 next
	   elseif PageCount>11 then
	      if pageid<=5 then
		     for i=1 to 10
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       else
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
		     next
		  else
		    for i=(pageid-4) to (pageid+5)
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       elseif i=<pagecount then
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
			next
		  end if
	   end if
  end sub

%>
<SCRIPT language=JavaScript>
<!--
function doConvert(){
window.open("LubeExcelReport.asp?From_Month=<%=Search_From_Month%>&From_Year=<%=Search_From_Year%>"); 

}

//-->
</SCRIPT>