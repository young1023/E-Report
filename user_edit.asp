<%
if session("shell_power")="" then
  response.redirect "index.asp"
end if
%>
<!--#include file="img/conne.inc" -->
<HTML><HEAD><TITLE>Job Management</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="img/home.css" type="text/css">
<STYLE type=text/css>.input {
}
TD {
	FONT-SIZE: 11pt
}
</STYLE>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=4>
<DIV align=center>
<!--#include file="img/head.inc" -->
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 height=20 width=99%>
 <TBODY> 
    <TR>
    <TD align=middle vAlign=top>
    <TABLE border=0 cellPadding=0 cellSpacing=0 height=1 width=100%><TBODY><TR><TD bgColor=#000000 height=1></TD></TR></TBODY></TABLE>
	</TD></TR>
  <TR>
    <TD>
        <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
          <TBODY> 
          <TR bgcolor="#006699"> 
            <TD height=18 width="180" align=center><script language=JavaScript src="img/date.js"></script>
            </TD>
            <TD align=right height=18 vAlign=center bgcolor="#006699"></TD>
          </TR>
          </TBODY> 
        </TABLE>
    </TD>
   </TR>
  </TBODY>
</TABLE>
  <TABLE border=0 cellPadding=0 cellSpacing=0 height=100% width=99%>
    <TBODY> 
    <TR>
      <TD vAlign=top>
        <table width="100%" border="0" cellpadding="1" cellspacing="0">
          <tr>
            <td><img src="img/back.gif" width="306" height="37"></td>
          </tr>
        </table>
        <table width="100%" border="0" cellpadding="1" cellspacing="1" height="100%">
          <tr> 
            <td bgcolor="#000000">
              <table width="100%" border="0" cellpadding=0 cellspacing="0" bgcolor="#FFFFFF" height="100%">
                <tr>
                  <td valign="top" align="center" bgcolor="#FFCC33" width="180">
<!--#include file="img/menu.inc" -->
                  </td>
                  <td valign="top" align="center" bgcolor="#E6EBEF">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E6EBEF">
                      <form name=ok method=post>
                        <tr> 
                          <td height="48" align="center"><b><font color="#FF6600">Job 
                            Management</font></b></td>
                        </tr>
                        <tr> 
                          <td valign="top" align="center"> 
                            <table width="93%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
                              <tr> 
                                <td height="28"> 
                                  <%
		pageid=trim(request.form("pageid"))
		if pageid="" then
		  pageid=1
		end if
        findnum=replace(trim(request.form("findnum")),"%","％")
        findnum=replace(findnum,"'","''")
		if findnum="" then
          fsql="select id,name,flage,createtime from modul where flage=1 order by createtime desc"
		else
          fsql="select id,name,flage,createtime from modul where flage=1 and name like '%"&findnum&"%' order by createtime desc"
		end if
        set frs=createobject("adodb.recordset")
		frs.cursortype=1
		frs.locktype=1
        frs.open fsql,conn

       if frs.RecordCount=0 then
           response.write "<font color=red>No Record</font>"
           'response.end
       else
           findrecord=frs.recordcount
          response.write "Total <font color=red>"&findrecord&"</font> Records ; Total <font color=blue>"

          frs.PageSize = 10
         call countpage(frs.PageCount,pageid)
	     response.write "&nbsp;&nbsp;<input type='text' name='findnum' size='13' value='"&findnum&"' style='BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 10pt' value='"&findnum&"' maxlength='18'>"
		 response.write "<input type='button' value='Find Number' onClick='findenum();' style='BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 9pt; HEIGHT: 20px; WIDTH:80px'>"
	   end if
	  %>
                                </td>
                              </tr>
                              <tr> 
                                <td valign="top" height="28"> 
                                  <table width="100%" border="0" align=center cellpadding="1" cellspacing="1">
                                    <tr> 
                                      <td bgcolor="#006699" align="center" height="21"><font color="#FFFFFF">Bill 
                                        Name</font></td>
                                      <td width="20%" bgcolor="#006699"><font color="#FFFFFF">Property</font></td>
                                      <td bgcolor="#006699" align="center" width="13%"><font color="#FFFF66">Date</font></td>
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
   response.write "<tr bgcolor="&mycolor&">"
   response.write "<td onmouseover=javascript:style.background='#cccccc' onmouseout=javascript:style.background='"&mycolor&"'>"
   if frs("id")=3 then
   %>
                                    <span onClick="javascript:openwin('userpost2.asp?id=<%=frs("id")%>')" style="cursor:hand"><u><font color=blue><%=replace(trim(server.htmlencode(frs("name"))),vbcrlf,"<br>")%></font></u></span>
   <%
   else
   %>
                                    <span onClick="javascript:window.open('userpost.asp?id=<%=frs("id")%>','user','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=30,left=60,width=680,height=480');" style="cursor:hand"><u><%=replace(trim(server.htmlencode(frs("name"))),vbcrlf,"<br>")%></u></span>
                                    <%
   end if
   response.write "</td>"
   response.write "<td>"
   if frs("flage")=1 then
       response.write "Active"
   else
       response.write "Close"
   end if
   response.write "</td>"
   response.write "<td align=center>"&datevalue(frs("createtime"))&"</td>"
   response.write "</tr>"
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
document.ok.pageid.value=what;
document.ok.action="user_edit.asp"
document.ok.submit();
}

function findenum()
{
document.ok.pageid.value=1;
document.ok.action="user_edit.asp"
document.ok.submit();
}

function openwin(what)
{
 window.open(what,"","toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes,resizable=yes,top=0,left=0,width=790,height=520");
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
                                <td valign="top">&nbsp;</td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </form>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <TR>
      <TD colSpan=5 height=11 align=center>
        <script language=JavaScript src="img/copyright.js"></script>
      </TD>
    </TR>
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
		       else
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
			next
		  end if
	   end if
  end sub
%>
</BODY></HTML>