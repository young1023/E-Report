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
     Search_From_Month = month(Session("DBLastModifiedDateValue")) - 1
 End If
 
      If len(Search_From_Month) = 1 Then
           Search_From_Month = "0" & Search_From_Month
      End if

If Search_From_Year = "" Then      
	Search_From_Year = year(Session("DBLastModifiedDateValue"))  
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
 
 document.fm1.action="ReconReport.asp?sid=<%=SessionID%>";
 document.fm1.page.value=what;
 document.fm1.submit();
	
}

function doretrieve(){

 //   window.open('Retrieve_ABC.asp?sid=<%=SessionID%>', 'winname', 'directories=no,titlebar=no,toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400,height=350');
 document.fm1.action="Retrieve_ABC.asp?sid=<%=SessionID%>";
document.fm1.submit();
	
}


function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="ReconReport.asp?sid=<%=SessionID%>"
document.fm1.submit();
}

function findenum()
{
document.fm1.pageid.value=1;
document.fm1.action="ReconReport.asp?sid=<%=SessionID%>"
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
     
     <td width="20%"> <input id="Button1" type="button" value="Retrieve ABC Position" onClick="doretrieve();"></td>
      <td></td>

 	     
	
    
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
	
      <td width="20%">Depot</td> 
      <td>
<%
      
       Sql_d = "Select DepotCode from ReconDepotfolder order by DepotCode asc"

       Set Rs_d = Conn.Execute(Sql_d)

%>

<select size="1" name="DepotCode" class="common">
			<option value="" <% if Search_Depot="" then response.write "selected" %> >All</option>
			<%
					do while (  Not Rs_d.EOF)
			%>
					<option value="<%=Rs_d("DepotCode")%>" <% if Trim(Search_DepotCode)=Trim(Rs_d("DepotCode")) then response.write "selected" %> ><%=Rs_d("DepotCode")%></option>
			
			<%
					    Rs_d.movenext
					Loop
			%>
	</select>


</td>

</tr>
    
 
    

 <tr> 

    <td width="20%">Status:</td> 
	<td width="30%">	
    
            <select size="1" name="Match" class="common">
	        <option value="" <% if Search_Match="" then response.write "selected" %> >ALL</option>
		    <option value="1" <% if Search_Match="1" then response.write "selected" %> >Match</option>
            <option value="2" <% if Search_Match="2" then response.write "selected" %> >Not Match</option>
	       </select>    
 	</td>
	<td width="20%"></td> 
	<td width="30%">   
 	
</td>
	
     
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

    response.write ("Exec Retrieve_MonthReport '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_Market&"', '"&Search_DepotCode&"', '"&Search_Match&"' , '"&iPageCurrent&"' ") 
              
	Rs1.open ("Exec Retrieve_MonthReport '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_Market&"', '"&Search_DepotCode&"', '"&Search_Match&"' , '"&iPageCurrent&"' ") ,  conn,3,1


	'assign total number of pages
	
	
	
	iRecordCount = Rs1("Tcount")
	
  

%>   
  
<div id="reportbody1" >

   
<br>

<table width="99%" border="0" class="normal"  cellspacing="1" cellpadding="2">
<span class="noprint">
<tr bgcolor="#FFFFCC"> 
<td  width="20%">�@</td>
      <td align="center">Reconciliation Exception Report</td> 
      <td align="right" width="20%">
      
 <%
 
 ' show excel and print button only when there is record
 
 if Search_DepotCode <> "" then
 
 %>
						
<a href="javascript:window.doConvert()">Excel</a>&nbsp;<a href="javascript:window.print()">Friendly Print</a>

<%

 end if

 %>
					   	
			</td>
</tr>

<tr>
</span>
   <td>

<%




  if iRecordCount <= 0 then


		'no record found
		response.write ("No record found")
						
	else

		if Search_DepotCode <> "" then
		
        'cal total no of pages
		iPageCount = int(iRecordCount / 1000) + 1
		
		else
		
		 'cal total no of pages
		iPageCount = int(iRecordCount / 10) + 1
		
		end if

		'move to next recordset
  	Set Rs1 = Rs1.NextRecordset() 

%>
</td>
 <td  colspan="2" align="right" >

<%
response.write (iPageCurrent & " Page of " & iPageCount &"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp" )

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
      
      <td width="20%" >Depot</td>
      <td width="7%" >Depot Code</td>
      <td width="7%" >STOCK Code</td>
      <td width="30%">Instrument Name</td>
      <td>Custodian Position</td>
      <td>ABC Position</td>
      <td>Difference</td> 
      
        
</tr>
<%		
     ' define varable to check instrument code and qty
    QTY0 = 0 
    QTY1 = 0
    Instrument0 = ""
    Instrument1 = ""
    i=0
    
    
  do while Not Rs1.EoF

   if Rs1.eof then exit do
   
   Instrument0 = Instrument1
   QTY0        = QTY1

   i=i+1
		
%>
<tr bgcolor="#FFFFCC"> 
<td>
<%
        Response.write i & ".&nbsp;"
        Response.Write Rs1("DepotName")


 
%>
</td>
<td>
<%
   
        Response.Write Rs1("DepotCode")

        'Response.write Instrument0
 
%>
</td>

<td><%
       If Rs1("Instrument") <> "" Then
       
        Response.Write Rs1("Instrument")
        
        Instrument1 = Rs1("Instrument")
        
       Elseif Rs1("ISIN") <> "" Then
       
        Response.Write Rs1("ISIN")
        
        Instrument1 =  Rs1("ISIN")
        
       Else
       
        Response.Write Rs1("Sedol")
        
         Instrument1 = Rs1("Sedol")
        
       End If

   
%></td>
<td>
<% 



        Response.Write Rs1("InstrumentName")

%>
</td>


<td><% = formatnumber(Rs1("UnitHeld"),0) %></td>
<td ><%
         QTY1 = formatnumber(Rs1("TotalQTY"),0)

         If trim(Instrument0) = Trim(Instrument1) and trim(QTY0) = trim(QTY1) then
         
          Response.Write "0"
          
          Else

          Response.Write formatnumber(Rs1("TotalQTY"),0)
          
          End if
 
    %>

</td>
<td>
<%  

            If trim(Instrument0) = Trim(Instrument1) and trim(QTY0) = trim(QTY1) then  

            Response.Write formatnumber(Rs1("UnitHeld"),0)
            
    Else

 %>

<% = formatnumber((formatnumber(Rs1("UnitHeld"),0) - formatnumber(Rs1("TotalQTY"),0)),0)   %>

<% End If  %>

</td> 



</tr>


<%

				
					
				Rs1.movenext
				
		loop
	


%>

<%		

    If Search_Match => 2 Then

   Set rs1 = rs1.NextRecordset() 
    

  do while Not Rs1.EoF

   if Rs1.eof then exit do

   i=i+1
		
%>
<tr bgcolor="#FFFFCC"> 
<td>
<%
        Response.write i & ".&nbsp;"
        Response.Write Rs1("DepotName")


 
%>
</td>
<td>
<%
   
        Response.Write Rs1("DepotCode")


 
%>
</td>

<td><%
       If Rs1("Instrument") <> "" Then
       
        Response.Write Rs1("Instrument")
        
       Elseif Rs1("ISIN") <> "" Then
       
        Response.Write Rs1("ISIN")
        
       Else
       
        Response.Write Rs1("Sedol")
        
       End If


 
%></td>
<td>
<% 

        

        sql2 = "Select distinct ShortName from UOBKHHKEQPRO.dbo.Instrument where "
        
        If Trim(Rs1("Instrument")) <> "" then

        sql2 = sql2 & " Instrument = '" & Trim(Rs1("Instrument")) & "'"
        
        Check
        
        Elseif Trim(Rs1("ISIN")) <> "" then
        
        sql2 = sql2 & " ISIN = '" & Trim(Rs1("ISIN")) & "'"
        
        Else 
        
        sql2 = sql2 & " Sedol = '" & Trim(Rs1("Sedol")) &"'"
        
        End if
             
        'response.write sql2

        Set Rs2 = Conn.execute(sql2)
        
        If not Rs2.Eof then
        
           Do While Not Rs2.EoF

        Response.Write Rs2("ShortName") & "<br/>"
        
           Rs2.Movenext
           
           Loop
        
        Else
        
        Response.write "Instrument name cannot be found."
        
        End if

%>
</td>


<td><% = formatnumber(Rs1("UnitHeld"),0) %></td>
<td ><% = formatnumber(Rs1("TotalQTY"),0)%></td>
<td><% = formatnumber((formatnumber(Rs1("UnitHeld"),0) - formatnumber(Rs1("TotalQTY"),0)),0)   %></td> 



</tr>


<%

				
					
				Rs1.movenext
				
		loop
	


 Set rs1 = rs1.NextRecordset() 

    

  do while Not Rs1.EoF

   if Rs1.eof then exit do

   i=i+1
		
%>
<tr bgcolor="#FFFFCC"> 
<td>
<%
        Response.write i & ".&nbsp;"
        Response.Write Rs1("DepotName")


 
%>
</td>
<td>
<%
   
        Response.Write Rs1("DepotCode")


 
%>
</td>

<td><%
      
       
        Response.Write Rs1("StockCode")
        
 


 
%></td>
<td>
<% 



        Response.Write Rs1("InstrumentName")

%>
</td>


<td><% = formatnumber(Rs1("UnitHeld"),0) %></td>
<td ><% = formatnumber(Rs1("TotalQTY"),0)%></td>
<td><% = formatnumber((formatnumber(Rs1("UnitHeld"),0) - formatnumber(Rs1("TotalQTY"),0)),0)   %></td> 



</tr>


<%

				
					
				Rs1.movenext
				
		loop
	
	

End if

End If

%>

                              <tr bgcolor="#FFFFCC"> 
                                <td align="left" colspan="10" height="28"> 




<span class="noprint">
 <%

     
  
 
 ' show excel and print button only when there is record
  
	 if   iRecordCount > 0 then
               
             response.write "Total <font color=red>"& i &"</font> Records "
   
	 end if

if   findrecord>0 then
  response.write "<input type=hidden value='' name=whatdo>"
  response.write "<input type=hidden value="&pageid&" name=pageid>"
end if
			  Rs1.close
			  set Rs1=nothing
			  conn.close
			  set conn=nothing
%>

</span>
   </td>
     </tr>                             
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
window.open("ReconExcelReport.asp?DepotCode=<%=Search_DepotCode%>&Search_Match=<%=Search_Match%>&Search_Market=<%=Search_Market%>&From_Month=<%=Search_From_Month%>&From_Year=<%=Search_From_Year%>"); 

}

//-->
</SCRIPT>