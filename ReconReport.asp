<!--#include file="include/SessionHandler.inc.asp" -->

<%

Search_From_Month       = Request.form("FromMonth")
Search_From_Year        = Request.form("FromYear")
Search_Market           = Request.form("Market")
Search_Instrument       = Request.form("Instrument")
Search_ISIN             = Request.form("ISIN")
Search_Sedol            = Request.form("Sedol")
Search_Match            = Request.form("Match")
whatdo                  = Request.form("whatdo")

If whatdo = "1" then


   set RsABC = server.createobject("adodb.recordset")
   RsABC.open ("Exec Process_ReconMonthly") ,  StrCnn,3,1


End if




' Retrieve page to show or default to the first
pageid=trim(request.form("pageid"))
	
If Request.form("pageid") = "" Then
	Pageid = 1
End if

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
<SCRIPT language=JavaScript>
<!--

function dosubmit(){
 
 document.fm1.action="ReconReport.asp?sid=<%=SessionID%>";
 document.fm1.submit();
	
}

function doRetrieve(){
 
 document.fm1.action="ReconReport.asp?sid=<%=SessionID%>";
 document.fm1.whatdo.value=1;
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

<body leftmargin="0" topmargin="0" OnLoad="document.fm1.submitted.value=0;document.fm1.Instrument.focus();">

<!-- #include file ="include/Master.inc.asp" -->

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
      <td><input id="Retrieve" type="button" value=" Retrieve last month data from ABC " onClick="doRetrieve();"></td>
  
 	     
	
    
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

    <td width="20%">ISIN Code:</td> 
	<td width="30%"> <input name="ISIN" type=text value="<%= Search_ISIN %>" size="15">&nbsp;   
 	</td>
	<td width="20%">Sedol:</td> 
	<td width="30%"> <input name="Sedol" type=text value="<%= Search_Sedol %>" size="15">&nbsp;   
 	
</td>
	
     
    </tr>
    

 <tr> 

    <td width="20%">Status:</td> 
	<td width="30%">	
    
            <select size="1" name="Match" class="common">
			<option value="" <% if Search_Match ="" then response.write "selected" %> >ALL</option>
		    <option value="1" <% if Search_Match="1" then response.write "selected" %> >Match</option>
            <option value="2" <% if Search_Match="2" then response.write "selected" %> >No Match</option>
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

          <input id="Submit1" type="button" value="Submit" onClick="dosubmit();"></td>

		</tr>    

  

    </table>
  

<%


' Start the Queries
' *****************
                set frs = server.createobject("adodb.recordset")
                
                 fsql = "select r.DepotName as 'DepotName', I.Market as 'Market', M.Location, r.DepotCode as 'DepotCode', "

                 fsql = fsql & "I.Instrument as 'Local Code', S.Instrument as 'Instrument', S.ISINCode as 'ISIN', S.Sedol as 'Sedol', InstrumentName, UnitHeld, TotalQTY "
                 
                 fsql = fsql & "  from ReconMonthly M Join UOBKHHKEQPRO.dbo.Instrument I on M.Instrument = I.Instrument  "
                 
                 fsql = fsql & " join ReconDepotFolder r on M.depot = Cast(r.depotCode as varchar)  Join StockReconciliation S "
                 
                 fsql = fsql & " on  r.DepotID = S.DepotID "
                 
                 fsql = fsql & " and ( I.Instrument = S.Instrument or  I.ISIN = S.ISINCode "
                 
                 fsql = fsql & " or   I.Sedol = S.Sedol ) "

                 If Search_Match = "1" Then

                 fsql = fsql & " and Cast(Cast(UnitHeld as float) as decimal) = Cast(Cast(totalQTY as float) as decimal)"

                 ElseIf Search_Match = "2" then

                 fsql = fsql & " and cast(Cast(UnitHeld as float) as decimal) <> Cast(Cast(totalQTY as float) as decimal)"

                 End If
                 
                 If Search_Market <> "" Then
                 
                 fsql = fsql & " and I.Market = '"& Search_Market &"'"
                 
                 End If

                 If Search_Market <> "" Then
                 
                 fsql = fsql & " and I.Market = '"& Search_Market &"'"
                 
                 End If

                 If Search_Market <> "" Then
                 
                 fsql = fsql & " and I.Market = '"& Search_Market &"'"
                 
                 End If

                 If Search_Market <> "" Then
                 
                 fsql = fsql & " and I.Market = '"& Search_Market &"'"
                 
                 End If

             
                'response.write fsql
                set frs=createobject("adodb.recordset")
		        frs.cursortype=1
		        frs.locktype=1
                frs.open fsql,conn

%>   
  
<div id="reportbody1" >

   
<br>

<table width="99%" border="0" class="normal"  cellspacing="1" cellpadding="2">
<tr bgcolor="#FFFFCC"> 
<td  width="20%">¡@</td>
      <td align="center">Reconciliation Exception Report</td> 
      <td align="right" width="20%">
						
<a href="javascript:window.doConvert()">Excel</a>
					   	
			</td>
</tr>

<tr>
   <td bgcolor="#FFFFCC" colspan="3">

<%

        if frs.RecordCount=0 then

           'response.write "<tr bgcolor=#ffffff align=center><td colspan=7><font color=red>No Record</font></td></tr>"
 
        else


          findrecord=frs.recordcount

          response.write "Total <font color=red>"&findrecord&"</font> Records ;"
  
         frs.PageSize = 100

         call countpage(frs.PageCount,pageid)

         end if

%>
</td>
</tr>
</table>
<br>

<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

<tr bgcolor="#ADF3B6" align="center">

      <td width="10%" >Market: </td>
      <td width="10%" bgcolor="#FFFFCC">
      <% 
           If Search_Market = "" Then
               
               Response.Write "ALL"

           Else

               Response.Write frs("Market") 

           End If
       %></td>
      <td width="30%" bgcolor="#FFFFCC"></td>
      <td bgcolor="#FFFFCC">
           
      </td>
       
</tr>
</table>
<br>


<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

<tr bgcolor="#ADF3B6" align="center">
      
      <td width="20%" >Depot</td>
      <td width="7%" >Depot Code</td>
      <td width="7%" >Location</td>
      <td width="7%" >STOCK Code</td>
      <td width="7%">Local Code</td>
      <td width="30%">Instrument Name</td>
      <td>Status</td>
      <td>ABC Position</td>
      <td>Custodian Position</td>
      <td>Difference</td>    
</tr>
<%		
    i=0
 if frs.recordcount>0 then
  frs.AbsolutePage = pageid
  do while (frs.PageSize-i)
   if frs.eof then exit do
   i=i+1
		
%>
<tr bgcolor="#FFFFCC"> 
<td>
<%
   
        Response.Write frs("DepotName")


 
%>
</td>
<td>
<%
   
        Response.Write frs("DepotCode")


 
%>
</td>

<td>
<%
   
        Response.Write frs("Location")


 
%>
</td>
<td>
<%
       If frs("Instrument") <> "" Then
       
        Response.Write frs("Instrument")
        
       Elseif frs("ISIN") <> "" Then
       
        Response.Write frs("ISIN")
        
       Else
       
        Response.Write frs("Sedol")
        
       End If


 
%>
</td>
<td><% = frs("Local Code") %></td>
<td>
<% 



        Response.Write frs("InstrumentName")

%>
</td>

<td></td>
<td><% = formatnumber(frs("UnitHeld"),0) %></td>
<td ><% = formatnumber(frs("TotalQTY"),0)%></td>
<td><% = formatnumber((formatnumber(frs("UnitHeld"),0) - formatnumber(frs("TotalQTY"),0)),0)   %></td>
 



</tr>


<%

				
					
				frs.movenext
				
		loop
	
end if	

%>

                              <tr bgcolor="#FFFFCC"> 
                                <td align="right" colspan="10" height="28"> 




<span class="noprint">
 <%
	 if   findrecord >0 then
             call countpage(PageSize ,pageid)
			 response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			 if Clng(pageid)<>1 then
                 response.write " <a href=javascript:gtpage('1') style='cursor:hand' >First</a> "
                 response.write " <a href=javascript:gtpage('"&(pageid-1)&"') style='cursor:hand' >Previous</a> "
			 else
                 response.write " First "
                 response.write " Previous "
			 end if
	         if Clng(pageid)<>Clng(frs.PageCount) then
                 response.write " <a href=javascript:gtpage('"&(pageid+1)&"') style='cursor:hand' >Next</a> "
                 response.write " <a href=javascript:gtpage('"&frs.PageCount&"') style='cursor:hand' >Last</a> "
             else
                 response.write " Next "
                 response.write " Last "
			 end if
	         response.write "&nbsp;&nbsp;"
	 end if

if   findrecord>0 then
  response.write "<input type=hidden value='' name=whatdo>"
  response.write "<input type=hidden value="&pageid&" name=pageid>"
end if
			  frs.close
			  set frs=nothing
			  conn.close
			  set conn=nothing
%>

</span>
   </td>
     </tr>                             
</table>
</form>  


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

 frs.Close
 set frs = Nothing

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
window.open("ReconExcelReport.asp?Search_Instrument=<%=Search_Instrument%>&Search_Market=<%=Search_Market%>&From_Month=<%=Search_From_Month%>&From_Year=<%=Search_From_Year%>&To_Month=<%=Search_To_Month%>&To_Year=<%=Search_To_Year%>"); 

}

//-->
</SCRIPT>