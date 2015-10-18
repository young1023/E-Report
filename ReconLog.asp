
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


Title = "Audit Log"

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
function doDelete(){
document.fm1.action="execute.asp?sid=<%=sessionid%>";
document.fm1.whatdo.value = "deleted logs";

    if (document.fm1.DeleteLog.value == "") {
            alert("Please enter day value.");
            document.fm1.DeleteLog.focus();
            return false;
        }
 

document.fm1.submit();
}
 
function doConvert(){
document.fm1.action="execute.asp?sid=<%=sessionid%>";

    if (document.fm1.ConvertLog.value == "") {
            alert("Please enter day value.");
            document.fm1.ConvertLog.focus();
            return false;
        }
ConvertLog = document.fm1.ConvertLog.value;
window.open("ConvertLog.asp?ConvertLog=" + ConvertLog + "&sid=<%=sessionid%>"); 

}
//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0">



<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">
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
                          <td valign="top" align="center">
                            <form name=fm1 method=post>
<input name="Location_id" type="hidden" value="" >
                            <table width="99%" border="0" cellspacing="1" bgcolor="#FFFFFF" class="normal">

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
      
       fsql = "select a.createdate, m.name , a.description from ReconLog a , Member m where a.DoneBy = m.MemberID "


  ' Search by Date
  ' **************
   if Sdate <> "" then
      fsql = fsql & " and m.CreateDate >= #"& SDate &"# and m.CreateDate < #"& NDate &"# "   
   end if
   
     

  ' Searh by Name
  ' ***********************
   if findnum <> "" then
     ' fsql = fsql & " and Description like '%"&findnum&"%'"
   end if
  

       fsql = fsql & " order by a.LogID Desc"

  	    set frs=createobject("adodb.recordset")
		frs.cursortype=1
		frs.locktype=1
        frs.open fsql,conn
        'response.write fsql

       if frs.RecordCount=0 then

           response.write "<font color=red>No Record</font>"
           
       else
          findrecord=frs.recordcount
          response.write "Total <font color=red>"&findrecord&"</font> Records ; Total <font color=blue>"

         frs.PageSize = 10
         call countpage(frs.PageCount,pageid)
	   end if
	  %>
                                </td>
                              </tr>
                              <tr> 
                                <td valign="top" height="28"> 
   <div align="center">
<table border="0" cellpadding="5" cellspacing="1" class="normal" width="99%">
<tr bgcolor="#006699">
<td width="12%"><font color="#FFFFFF">Performed By</font></td>
<td width="61%"><font color="#FFFFFF">Job Performed</font></td>    
<td width="22%"><font color="#FFFFFF">Date &amp; Time</font></td> 
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

<td width="12%">
<% = frs("name")  %>
</td>

<td>
<%= frs("Description") %>
</td>
<td width="22%">on <%= frs("CreateDate") %>
</td>



</tr>
<%
   flage=not flage
   frs.movenext
  loop
 end if
  %>

                                  </table>
                                </div>
                                </td>
                              </tr>
                              <tr> 
                                <td align="right" height="28"> 
<script language=JavaScript>
<!--
function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="ReconLog.asp?sid=<%=SessionID%>"
document.fm1.submit();
}

function findenum()
{
document.fm1.pageid.value=1;
document.fm1.action="ReconLog.asp"
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
 </div>
   </body>
</html>