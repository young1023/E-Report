<!--#include file="include/SessionHandler.inc.asp" -->

<%

Search_From_Month       = Request.form("FromMonth")
Search_From_Year        = Request.form("FromYear")
Search_DepotID          = trim(request.form("Depot"))
Search_FileName         = request("FileName")

If Search_DepotID = "" Then

     Search_DepotID = Trim(Request("DepotID"))

End If




'response.write Search_DepotID

' Retrieve page to show or default to the first
pageid=trim(request.form("pageid"))
	
If Request.form("pageid") = "" Then
	Pageid = 1
End If
If Search_From_Month = "" Then
	Search_From_Month = month(Session("DBLastModifiedDateValue")) -1
End If
If Search_From_Year = "" Then
  	Search_From_Year = year(Session("DBLastModifiedDateValue"))
End If

If len(Search_From_Month) = 1 Then
           Search_From_Month = "0" & Search_From_Month
End if
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
 
 document.fm1.action="ReconCheckList.asp?sid=<%=SessionID%>";
 document.fm1.submit();
	
}


function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="ReconCheckList.asp?sid=<%=SessionID%>"
document.fm1.submit();
}

function findenum()
{
document.fm1.pageid.value=1;
document.fm1.action="ReconCheckList.asp?sid=<%=SessionID%>"
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
     
     <td ></td> 
      <td></td>   
 	     
	
    
    </tr>

     <%

            sql = "Select DepotID, DepotCode, DepotName From ReconDepotFolder order by depotID "

            set rs = Conn.Execute(sql)

     %>

    
<tr> 
	<td width="20%">Depot:</td> 
	<td width="50%">
	 
	<select size="1" name="Depot" class="common">
			<option value="">All</option>
			<%
                    If Not rs.Eof then

                        rs.MoveFirst



					do while (Not rs.EOF)
			%>
	
<option value="<%=rs("DepotID")%>" <% if trim(Search_DepotID)=trim(rs("DepotID")) then response.write "selected" end if %> ><% =rs("DepotCode")%>.<%=rs("DepotName")%></option>
			
			<%

					rs.movenext
					Loop

                   End If
			%>
	</select></td>
	
      <td >File Name:</td> 
      <td>
     <input name="FileName" type=text value="<%= Search_FileName %>" size="30">   
 	     </td>
    </tr> 

		
		<tr> 
			<td></td>
			<td colspan="3">
 	<input type=hidden   name="submitted"> 

          <input id="Submit1" type="button" value="Submit" onClick="dosubmit();"></td>

		</tr>    

    </table>
  

<%


' Start the Queries
' *****************
                set frs = server.createobject("adodb.recordset")
                
      
       fsql = "select  * "

       fsql = fsql & " from StockReconciliation s left join ReconDepotFolder r on s.depotid = r.depotid "

       fsql = fsql & " Where 1 = 1 "

       If Search_DepotID <> "" then

       fsql = fsql & " and s.DepotID = "&Search_DepotID

       end if

      If Search_FileName <> "" then

       fsql = fsql & " and s.ImportFileName like '%"&Search_FileName&"%'"

       end if

   
  ' Search by Date
  ' **************


        fsql = fsql & " and left(Importfilename,4) =   '" &Search_From_Month & Right(Search_From_Year,2)& "' " 

        fsql = fsql & " order by r.DepotCode"

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
      <td align="center">Broker Statement Crosscheck List</td> 
      <td align="right" width="20%">	   	
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
  
         frs.PageSize = 300

         call countpage(frs.PageCount,pageid)

         end if

%>
</td>
</tr>
</table>
<br>


<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

<tr bgcolor="#ADF3B6" align="center">

      <td width="40%">Depot</td>
      <td>Market</td>
      <td>Month</td>
      <td>Local Exchange Symbol</td>
      <td>Instrument Name</td>
      <td>Total Holding</td>
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
<td ><% = i & ". "%><% =frs("DepotName")%>
</td>
<td><% = frs("Market") %></td>
<td>
<% = MonthName(Left(frs("ImportFileName"),2)) & " " & Mid(frs("ImportFileName"),3,2) %>
</td>

<td>
<%
    Sql_I = "Select ShortName from UOBKHHKEQPRO.dbo.Instrument where 1=1 "

    If frs("ISINCode") <> "" Then

        Response.write frs("ISINCode")

    Sql_I = Sql_I & "and ISIN = '"&frs("ISINCode")&"'"

    End If

    If frs("CommonCode") <> "" Then

        Response.write  frs("CommonCode")

    Sql_I = Sql_I & "and Sedol= '"&frs("CommonCode")&"'"


    End If

    If frs("Instrument") <> "" Then

        Response.write frs("Instrument")

        Sql_I = Sql_I & "and  Instrument = '"&frs("Instrument")&"'"
   

    End If

     

    Set Rs_I = Conn.Execute(Sql_I)

%></td>
<td><% = Rs_I("ShortName") %></td>

<td><% = formatnumber(frs("UnitHeld"),0) %></td>
 



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

if frs.recordcount>0 then
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