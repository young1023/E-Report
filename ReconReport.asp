<!--#include file="include/SessionHandler.inc.asp" -->
<%
On Error resume Next

if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

Title = "Stock Reconciliation Report"

strURL = Request.ServerVariables("URL") ' Retreive the URL of this page from Server Variables

if session("shell_power")="" then
  response.redirect "Default.asp"
end if

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



function dosubmit(){
  

			  document.fm1.action="<%= strURL %>?sid=<%=SessionID%>";
			  document.fm1.submit();
	
}


//-->
</SCRIPT>

</head>

<body leftmargin="0" topmargin="0" OnLoad="document.fm1.submitted.value=0;document.fm1.ClientFrom.focus();">

<!-- #include file ="include/Master.inc.asp" -->

<div id="Content">


<%

'**************
'Argument handler
'**************

Dim Search_From_Day
Dim Search_From_Month
Dim Search_From_Year
Dim Search_To_Day
Dim Search_To_Month
Dim Search_To_Year
Dim Search_Market
Dim Search_Instrument


Search_From_Month       = Request.form("FromMonth")
Search_From_Year        = Request.form("FromYear")
Search_To_Month         = Request.form("ToMonth")
Search_To_Year          = Request.form("ToYear")
Search_Market           = Request.form("Market")
Search_Instrument       = Request.form("Instrument")





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


Search_Date = "01/" & Search_From_Month & "/" & Search_From_Year

Search_NDate = "01/" & Search_To_Month + 1 & "/" & Search_To_Year


set RsMarket = server.createobject("adodb.recordset")
RsMarket.open ("Exec Retrieve_AvailableMarket ") ,  StrCnn,3,1


%>

<form name="fm1" method="post" action="">
  <table width="97%" border="0" class="normal">
 <tr> 
      <td width="20%">Period From:</td> 
      <td>
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

          <input id="Submit1" type="button" value="Submit" onClick="dosubmit();"></td>

		</tr>    

    </table>
</form>    

<%




' Start the Queries
    ' Start the queries
' *****************
      
       fsql = "select * from StockReconciliation s left join ReconDepotFolder r on s.depotid = "

       fsql = fsql & "r.depotid" 

       'fsql = fsql & "left join UOBKHHKEQPRO.dbo.Instrument I on"

       'fsql = fsql & " S.ISINCode = I.ISIN "

       fsql = fsql & " Where 1 = 1 "

       'fsql = fsql & " and Coupon_Type = '"& Coupon_Type &"'"

       'fsql = fsql & " and Coupon_Batch = '"& Coupon_Batch &"'"

       
       'fsql = fsql & " and Coupon_Number LIKE '%"&Barcode&"%' " 

       

   

 
  ' Search by Date
  ' **************


      'fsql = fsql & " and  TradeDate >=   Convert(datetime, '" & Search_Date &"', 105) " 

  
      'fsql = fsql & " and  TradeDate < DATEADD(dd,DATEDIFF(dd,0, Convert(datetime, '" & Search_NDate &"', 105)),0) + 1 " 


      fsql = fsql & " order by r.DepotID desc"

        response.write fsql
        'response.end

        set frs=createobject("adodb.recordset")
		frs.cursortype=1
		frs.locktype=1
        frs.open fsql,conn


 
%>   

 
   
<div id="reportbody1" >

   
<br>

<table width="99%" border="0" class="normal"  cellspacing="1" cellpadding="2">
<tr bgcolor="#FFFFCC"> 
<td  width="20%">　</td>
      <td align="center">詳細紀錄<br><u>Stock Reconciliation Report</u></td> 
      <td align="right" width="20%">
						
<a href="javascript:window.doConvert()">Excel</a>
					   	
			</td>
</tr>

<tr>
   <td bgcolor="#FFFFCC" colspan="3">

<%

        if frs.RecordCount=0 then

           response.write "<tr bgcolor=#ffffff align=center><td colspan=7><font color=red>No Record</font></td></tr>"
 
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

      <td width="10">Depot</td>
      <td width="10">Market</td>
      <td>Custodian ID</td>
      <td>Trade Date</td>
      <td>ISINCode</td>
      <td>Common Code</td>
      <td>Security Name</td>
      <td>Description</td>
      <td>Unit Held</td>
      <td>Total Amount</td>
 




</tr>


		<%
			
Total = 0

i=0

 if frs.recordcount>0 then
  frs.AbsolutePage = pageid
  do while (frs.PageSize-i)
   if frs.eof then exit do
   i=i+1
		
		%>

<tr bgcolor="#FFFFCC"> 


<td><% = frs("DepotName") %></td>
<td><% = frs("Market") %></td>
<td><% = frs("CustodianID") %></td>
<td><% = frs("TradeDate") %></td>
<td><% = frs("ISINCode") %></td>
<td><% = frs("CommonCode") %></td>
<td><% = frs("SecurityName") %></td>
<td><% = frs("Description") %></td>
<td><% = frs("UnitHeld") %></td>
<td><% = frs("TotalAmount") %></td>
 



</tr>


<%

				
					
				frs.movenext
				
		loop
	
end if	

%>

                              <tr bgcolor="#FFFFCC"> 
                                <td align="right" colspan="10" height="28"> 
<script language=JavaScript>
<!--

function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="ReconReport.asp"
document.fm1.submit();
}

//-->
</script>



<span class="noprint">
<%
	 if frs.recordcount>0 then
             call countpage(frs.PageCount,pageid)
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
%>
</span>
                                </td>
                              </tr>
</table>



<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
	<tr bgcolor="#FFFFCC"> 
      <td width="99%" height="18" align="center">End of Statement</td>
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
window.open("ConvertStockReconciliation.asp?Search_Instrument=<%=Search_Instrument%>&Search_Market=<%=Search_Market%>&From_Day=<%=Search_From_Day%>&From_Month=<%=Search_From_Month%>&From_Year=<%=Search_From_Year%>&To_day=<%=Search_To_Day%>&To_Month=<%=Search_To_Month%>&To_Year=<%=Search_To_Year%>"); 

}

//-->
</SCRIPT>