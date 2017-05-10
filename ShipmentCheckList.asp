<!--#include file="include/SessionHandler.inc.asp" -->

<%

Search_From_Month       = Request.form("FromMonth")
Search_From_Year        = Request.form("FromYear")

whatdo                  = Request.form("whatdo")



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

Title = "Shipment Check List"

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
 
 document.fm1.action="ShipmentCheckList.asp?sid=<%=SessionID%>";
 document.fm1.submit();
	
}


function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="ShipmentCheckList.asp?sid=<%=SessionID%>"
document.fm1.submit();
}

function findenum()
{
document.fm1.pageid.value=1;
document.fm1.action="ShipmentCheckList.asp?sid=<%=SessionID%>"
document.fm1.submit();
}

function doDelete(){
 
 document.fm1.action="ShipmentCheckList.asp?sid=<%=SessionID%>";
 document.fm1.whatdo.value="del_file"
 document.fm1.submit();
	
}


function delcheck(){
k=0;
document.fm1.action="ShipmentCheckList.asp?sid=<%=SessionID%>";
	if (document.fm1.mid!=null)
	{
		for(i=0;i<document.fm1.mid.length;i++)
		{
			if(document.fm1.mid[i].checked)
			  {
			  k=1;
			  i=1;
			  break;
			  }
		}
		if(i==0)
		{
			if (!document.fm1.mid.checked)
               k=0;
			else
               k=1;
		}
	}

if (k==0)
  alert("You must select at least one record!");	
else if (k==1)
 {
  var msg = "Are you sure ?";
  if (confirm(msg)==true)
   {
    document.fm1.whatdo.value="del_record";
    document.fm1.submit();
   }
 }

}
//-->
</script>

</head>

<body leftmargin="0" topmargin="0" >

<!-- #include file ="include/Master.inc.asp" -->

<div id="Content">




<form name="fm1" method="post" action="">
  <table width="97%" border="0" class="normal">
 <tr> 
          <td >File Name:</td> 
      <td>
     <input name="FileName" type=text value="<%= Search_FileName %>" size="40">

     <input id="Submit2" type="button" value="  Search " onClick="dosubmit();">


  
 	     </td>
   
	
    
    </tr>

    

    </table>
  

<%



       set frs = server.createobject("adodb.recordset")
                
       fsql = "select  * "

       fsql = fsql & " from Shipment order by shipid asc "

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
<td  width="20%">　</td>
      <td align="center">Check List</td> 
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
      
      <td>Product</td>
      <td>Date</td>
      <td>QTY</td>
      <td>Sale Amount</td>
      <td>Ship To</td>
      <td>File Name</td>
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

<td ><% = i & ". "%><% =frs("Product")%>
</td>
<td>
<% = frs("Date") %>
</td>

<td><% = frs("QTY") %></td>
 
<td><% = frs("SaleAmount") %></td>

<td><% = frs("ShipTo") %></td>



<td><% = frs("FileName") %></td>

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
