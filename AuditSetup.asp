<!--#include file="images/conne.inc" -->
<% 
page_id=request("page_id")

if session("shell_power")="" then
  response.redirect "default.asp"
elseif session("shell_power")=0 then
  response.redirect "user.asp"
end if
%>

<html>

<head>
<title>E-Report</title>
<link rel="stylesheet" type="text/css" href="hse.css" />

<SCRIPT language=JavaScript>
<!--
function dosubmit(){
 document.ok.action="create_form.asp";
document.ok.submit();
}
//-->
</SCRIPT>

</head>

<body leftmargin="0" topmargin="0">

<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50"></IMG></td>
    <td align="right"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="3" class="NavyBlue" height="3"></td>
  </tr>
  <tr>
    <td colspan="3" class="NavyBlue">
    <table border="0" width="100%" cellspacing="0" cellpadding="0" class="TitleBar">
      <tr>
        <td width="100%">
        <table border="0" width="100%" cellspacing="0" cellpadding="0" class="TitleBar">
            <tr>
              <td><div align="center">
				<font class="Head" style="font-size: 13px">&nbsp;Maintenance Database </font></div></td>
            </tr>
            </table>
          </td>
          <td nowrap="true">
          </td>
        </tr>
      </table>
      </td>
    </tr>
    <tr>
      <td colspan="3" class="NavyBlue" height="1"></td>
    </tr>
    <tr>
      <td colspan="3" height="1"></td>
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

<DIV align=center>

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
						Setup Audit Item</b></font></td>
                        </tr>
                        <tr> 
                          <td valign="top" align="center">
                            <form name=ok method=post>
<table width="99%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="normal">
                              <tr> 
                                <td height="28"> 
                                  <%
		pageid=trim(request.form("pageid"))
		if pageid="" then
		  pageid=1
		end if
        findnum=replace(trim(request.form("findnum")),"%","¢H")
        findnum=replace(findnum,"'","''")

		  if findnum="" then
            fsql="select * from AuditSetup Order by AuditName"
              else
            fsql="select name from AuditSetup where AuditName like '%"&findnum&"%' order by AuditName desc"
		  end if

        set frs=createobject("adodb.recordset")
		frs.cursortype=1
		frs.locktype=1
        frs.open fsql,conn

       if frs.RecordCount=0 then
           response.write "<font color=red>No Record <a href='javascript:;' onclick='history.go(-1)'>[ Return ]</a></font>"
           'response.end
       else
          findrecord=frs.recordcount
          response.write "Total <font color=red>"&findrecord&"</font> Records ; Total <font color=blue>"

         frs.PageSize = 10
         call countpage(frs.PageCount,pageid)
	     response.write "&nbsp;&nbsp;<input type='text' name='findnum' size='10' value='"&findnum&"' class='common'>"
		 response.write "&nbsp;&nbsp;<input type='button' value='   Search   ' onClick='findenum();' class='common'>"
	   end if
	  %>
                                </td>
                              </tr>
                              <tr> 
                                <td valign="top" height="28"> 
                                  <table width="100%" border="0" align=center cellpadding="1" cellspacing="1" class="normal">
                                    <tr> 
                                      <td bgcolor="#006699" width="356">
										<font color="#FFFFFF">Audit Items</font></td>
                                       <td width="44%" bgcolor="#006699"  align="center">
										<font color="#FFFFFF">Selected</font></td>
                                    </tr>
                                    <%
 i=0
 if frs.recordcount>0 then
  frs.AbsolutePage = pageid
  do while (frs.PageSize-i)
   if frs.eof then exit do
   i=i+1

   response.write "<tr bgcolor="&mycolor&">"
   response.write "<td onmouseover=javascript:style.background='#cccccc' onmouseout=javascript:style.background='"&mycolor&"'>"
%>
<% = frs("AuditName") %>
           </td>
               

   <td align=center> 
       <input type="checkbox" name="mid" <% If Trim(frs("AuditType"))=1 Then%>checked<% End If%> value="<% =frs("AuditID") %>">
       <input type="hidden" value="<% =frs("AuditID") %>" name="ResetID">
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
function doSelect(){
k=0;
document.ok.action="hsemis.asp?page_id=execute"
	if (document.ok.mid!=null)
	{
		for(i=0;i<document.ok.mid.length;i++)
		{
			if(document.ok.mid[i].checked)
			  {
			  k=1;
			  i=1;
			  break;
			  }
		}
		if(i==0)
		{
			if (!document.ok.mid.checked)
               k=0;
			else
               k=1;
		}
	}

if (k==0)
  alert("You must  select one record at least !");	
else if (k==1)
 {
  var msg = "Are you sure ?";
  if (confirm(msg)==true)
   {
    document.ok.whatdo.value="AuditSelect";
    document.ok.submit();
   }
 }

}

function gtpage(what)
{
document.ok.pageid.value=what;
document.ok.action="AuditSetup.asp"
document.ok.submit();
}

function findenum()
{
document.ok.pageid.value=1;
document.ok.action="AuditSetup.asp"
document.ok.submit();
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
                                  <%
  if frs.recordcount>0 then
  response.write "<input type='button' value='  Select   ' onClick='doSelect();' class='common'>"
  response.write "<input type=hidden value='' name=whatdo>"
  response.write "<input type=hidden value="&pageid&" name=pageid>"
  end if
  response.write "<input type=hidden value='' name='doc_type'>"

			  frs.close
			  set frs=nothing
			  conn.close
			  set conn=nothing
%>
                                </td>
                              </tr>
                              <tr> 
                                <td valign="top"></td>
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
</DIV>
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
%>



























<%
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
              <center>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td colspan="2" height="1"></td>
                </tr>
                <tr class="HSEBlue">
                  <td colspan="2" height="1"></td>
                </tr>
              </table>
              </center>
            
              </body>

              </html>