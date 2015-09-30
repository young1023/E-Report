<!--#include file="images/conne.inc" -->
<% 

' get the page id
'****************
page_id=request("page_id")

' get the location_id
'*********************
location_id=request("location_id")


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>E-Report</title>
<link rel="stylesheet" type="text/css" href="hse.css" />
<SCRIPT language=JavaScript>
<!--
function newWindow(file,window) {
  msgWindow=open(file,window,'resizable=yes,width=800,height=600');
  if(msgWindow.opener == null) msgWindow.opener = self;
}

function doSearch(){
document.fm1.action="searchhistory.asp";
  if (document.fm1.Location.value == "")
  {
   alert("Please select the location.");
   document.fm1.Location.focus();
   return false;
  }
document.fm1.submit();
}
//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0">
<p>　</p>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="NavyBlue" height="3"></td>
  </tr>
  <tr>
    <td class="NavyBlue">
    <table border="0" width="100%" cellspacing="0" cellpadding="0" class="TitleBar">
      <tr>
        <td width="100%">
        <table border="0" width="100%" cellspacing="0" cellpadding="0" class="TitleBar">
            <tr>
              <td><div align="center"><font class="Head" style="font-size: 13px"> Maintenance Database </font></div></td>                    
            </tr>
            </table>
          </td>
          <td nowrap="true">
          </td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="60%">
    <tr>
      <td width="180" valign="top">
      <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" class="Common">
        <tr valign="top" align="center">
          <td class="HSEBlue" height="21"></td>
        </tr>
        <tr valign="top" align="center">
          <td>
<!-- //#include file ="menu.inc" -->
          </td>
        </tr>
      </table>
      </td>
      <td width="1" class="HSEBlue"></td>
      <td valign="top">
      <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="4" class="HSEBlue"></td>
        </tr>
        <tr valign="top">
          <td height="25"><img src="images/Curve.gif" width="22" height="16" /></td>
        </tr>
        <tr valign="top">
          <td height="100%" align="middle">
<%
'-----------------------------------------------------------------------------
'
'      Main Content of the page is inserted here
'
'-----------------------------------------------------------------------------

%>
  <TABLE border=0 cellPadding=0 cellSpacing=0 height=100% width=99%>
    <TBODY> 
    <TR>
      <TD vAlign=top>
        <table width="100%" border="0" cellpadding="1" cellspacing="1" height="100%">
          <tr> 
            <td bgcolor="#000000">
              <table width="100%" border="0" cellpadding=0 cellspacing="0" bgcolor="#FFFFFF" height="100%">
                <tr>
                  <td valign="top" align="center" bgcolor="#E6EBEF">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E6EBEF" class="normal">
                     <tr> 
                          
                        <td height="48" align="center"><font color="#FF6600"><b>
				Checking Historical Records</b></font></td>
                        </tr>
                        <tr> 
                          <td valign="top" align="center">
                            <form name=fm1 method=post>
<input name="Location_id" type="hidden" value="" >
                            <table width="99%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="normal">
                              <tr>
                                <td height="18">  
      

 <%
						strsql="select * from StoredPro order by ReportID"
						set acres=conn.execute(strsql)
%>
Report Name:         
    
<select name="SPName" class="common" size="1" tabindex="1">
   <option>- Please Select -</option>
<% 
                        acres.MoveFirst
							do while not acres.eof
                              if assignto = trim(acres("ReportID")) then
                                 response.write "<option value="&trim(acres("ReportID"))&" selected>"&trim(acres("ReportID"))&"&nbsp;&nbsp;"&trim(acres("Description"))&"</option>"
                                 else
                                 response.write "<option value="&trim(acres("ReportID"))&">"&trim(acres("ReportID"))&"&nbsp;&nbsp;"&trim(acres("Description"))&"</option>"
                               end if
                               acres.movenext
							loop
%>
        </select></td>
                              </tr>
                              <tr>
                <td class="BlueClr">Start:                     
         <select name="SMonth" class="common">
          <option value="1" selected>1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4" >4</option>
           <option value="5">5</option>
          <option value="6" >6</option>
          <option value="7">7</option>
          <option value="8" >8</option>
          <option value="9">9</option>
          <option value="10" >10</option>
          <option value="11">11</option>
          <option value="12" >12</option>
         
         
         </select>
<select name="SMonth0" class="common">                    
          <option value="1" selected>Jan</option>
          <option value="2">Feb</option>
          <option value="3">Mar</option>
          <option value="4" >Apr</option>
           <option value="5">May</option>
          <option value="6" >Jun</option>
          <option value="7">Jul</option>
          <option value="8" >Aug</option>
          <option value="9">Sep</option>
          <option value="10" >Oct</option>
          <option value="11">Nov</option>
          <option value="12" >Dec</option>
</select>
<select name="SYear" class="common">                    
          <option value="2003" selected>2003</option>
          <option value="2004">2004</option>
          <option value="2005">2005</option>
          <option value="2006" >2006</option>
          <option value="2007">2007</option>
          <option value="2008" >2008</option>
           <option value="2009">2009</option>
          <option value="2010" >2010</option>
         </select>

End:                     
         <select name="NMonth" class="common">
         <option value="1" selected>1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4" >4</option>
           <option value="5">5</option>
          <option value="6" >6</option>
          <option value="7">7</option>
          <option value="8" >8</option>
          <option value="9">9</option>
          <option value="10" >10</option>
          <option value="11">11</option>
          <option value="12" >12</option>
          </select>
<select name="SMonth0" class="common">                    
          <option value="1" selected>Jan</option>
          <option value="2">Feb</option>
          <option value="3">Mar</option>
          <option value="4" >Apr</option>
           <option value="5">May</option>
          <option value="6" >Jun</option>
          <option value="7">Jul</option>
          <option value="8" >Aug</option>
          <option value="9">Sep</option>
          <option value="10" >Oct</option>
          <option value="11">Nov</option>
          <option value="12" >Dec</option>
</select>
<select name="NYear" class="common">                    
          <option value="2003" selected>2003</option>
          <option value="2004">2004</option>
          <option value="2005">2005</option>
          <option value="2006" >2006</option>
          <option value="2007">2007</option>
          <option value="2008" >2008</option>
           <option value="2009">2009</option>
          <option value="2010" >2010</option>
        </select>
 <input type="button" value="Submit" onClick="dosubmit();" class="common">                    
                  &nbsp;                   
</td>
                              </tr>
                              <tr> 
                                <td height="28"> 
                                  <%
		pageid=trim(request.form("pageid"))
		if pageid="" then
		  pageid=1
		end if

' Start the Queries
    ' Start the queries
' *****************
      
       fsql = "select * from SPSchedule s , SystemMessage m where s.JobID = m.JobID "


  ' Search by Date
  ' **************
   if Sdate <> "" then
      fsql = fsql & " and StartDate >= #"& SDate &"# and StartDate < #"& NDate &"# "   
   end if
   
     

  ' Searh by Name
  ' ***********************
   if findnum <> "" then
      fsql = fsql & " and Description like '%"&findnum&"%'"
   end if
  

       fsql = fsql & " order by StartDate desc"

  	    'response.write fsql
  	    'response.end
  	    set frs=createobject("adodb.recordset")
		frs.cursortype=1
		frs.locktype=1
        frs.open fsql,conn
        'response.write fsql

       if frs.RecordCount=0 then

           response.write "<font color=red>No Record</font>"
           'response.end
       else
          findrecord=frs.recordcount
          response.write "Total <font color=red>"&findrecord&"</font> Records ; Total <font color=blue>"

         frs.PageSize = 10
         call countpage(frs.PageCount,pageid)
	     'response.write "&nbsp;&nbsp;<input type='text' name='findnum' size='13' value='"&findnum&"' class='common'>"
		 'response.write "&nbsp;&nbsp;<input type='button' value='Search by PO Number' onClick='findenum();' class='common'>"
	   end if
	  %>
                                </td>
                              </tr>
                              <tr> 
                                <td valign="top" height="28"> 
   <table border="0" align=center cellpadding="5" cellspacing="1" class="normal" width="98%">

<tr>
<td colspan="6" height="100%" class="BlueClr" width="391">
<br>
</td>
</tr>


<tr bgcolor="#efefef">
<td width="37">Job ID</td>         
<td width="38">Name</td>
<td width="80">Started Date & Time</td>
<td width="67">End Date &amp; Time</td>
<td width="51">Status</td> 
<td width="202">System Message</td>
</tr>
<%
 i=0
 if frs.recordcount>0 then
  frs.AbsolutePage = pageid
  do while (frs.PageSize-i)
   if frs.eof then exit do
   i=i+1
   if flage then
     mycolor="#ffffff"
   else
	 mycolor="#efefef"
   end if
%>
<tr>
<td valign="bottom" width="37"><a href="pmo.asp?id=<% = frs("id")%>"><% = frs("JOBID")%></a>
</td>
<td valign="bottom" width="38"><a href="pmo.asp?id=<% = frs("id")%>"><% = frs("StoreProcedure")%></a>
</td>
<td width="80">
<%= frs("StartedDate") %>
</td>
<td width="67"><% = frs("EndDate")  %>
</td>
<td width="51"><% = frs("Status")  %>
</td>
<td width="202">
</td>
</tr>
<%
   flage=not flage
   frs.movenext
  loop
 end if
  %>

                                  </table>
                                </td>
                              </tr>
                              <tr> 
                                <td align="right" height="28"> 
<script language=JavaScript>
<!--
function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="searchhistory.asp"
document.fm1.submit();
}

function findenum()
{
document.fm1.pageid.value=1;
document.fm1.action="searchhistory.asp"
document.fm1.submit();
}
//-->
</script>
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
                                </td>
                              </tr>
                              <tr> 
                                <td height="28" align="center"> 
  &nbsp;<%
if frs.recordcount>0 then
  'response.write "<input type='button' value='Chose the Record To Del' onClick='delcheck();' style='BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 9pt; HEIGHT: 20px; WIDTH:160px'>"
  response.write "<input type=hidden value='' name=whatdo>"
  response.write "<input type=hidden value="&pageid&" name=pageid>"
end if
			  frs.close
			  set frs=nothing
			  conn.close
			  set conn=nothing
%>
                                                         
 </td>
                              </tr>
                              <tr> 
                                <td valign="top">　</td>
                              </tr>
                            </table>
                          </form>


                          </td>
                        </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
</TABLE>
<%

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

'-----------------------------------------------------------------------------
'
'      End of the main Content 
'
'-----------------------------------------------------------------------------
%>
        </td>
            </tr>
                </table>
        </td>
             </tr>
                </table>
   </body>
</html>