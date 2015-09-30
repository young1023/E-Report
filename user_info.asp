<%
if session("shell_power")="" then
  response.redirect "index.asp"
end if
%>
<!--#include file="img/conne.inc" -->
<%
sql="select name,indicate,phone from member where employeenum='"&session("u_enum")&"' "
set rs=conn.execute(sql)
if not rs.eof then
  name=trim(rs("name"))
  indicate=trim(rs("indicate"))
  phone=trim(rs("phone"))
end if

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<HTML><HEAD><TITLE>Shell</TITLE>
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
                    <table width="98%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E6EBEF">
                      <form name=fm1 method=post>
                        <tr> 
                          <td height="48" align="center" colspan="2"><b><font color="#FF6600">Update 
                            Information</font></b> 
                            <SCRIPT language=JavaScript>
<!--
function check()
{
document.fm1.action="execute.asp";
if(document.fm1.name.value=="")
 {
  alert("Please Input Name");
  document.fm1.name.focus();
  return false;
 }
if(document.fm1.indicator.value=="")
 {
  alert("Please Input Region");
  document.fm1.indicator.focus();
  return false;
 }
if(document.fm1.phone.value=="")
 {
  alert("Please Input Department");
  document.fm1.phone.focus();
  return false;
 }

document.fm1.submit();
}
//-->
</SCRIPT>
                          </td>
                        </tr>
                        <tr> 
                          <td align="right" height="28" width="40%">E-mail: </td>
                          <td align="left">&nbsp;<font color=blue><%=session("u_enum")%></font> 
                            <%
						if session("shell_power")=8 then
						   response.write "&nbsp;&nbsp;( <font color=red>Super Administrator</font> )"
						elseif session("shell_power")=3 then
						   response.write "&nbsp;&nbsp;( <font color=red>Administrator</font> )"
						else
						   response.write "&nbsp;&nbsp;( <font color=red>User</font> )"
						end if
						%>
                          </td>
                        </tr>
                        <tr> 
                          <td align="right" height="28">Name : </td>
                          <td align="left"> 
                            <input type="text" name="name" maxlength="18" size="23" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #FFCC33 1px solid; BORDER-LEFT: #FFCC33 1px solid; BORDER-RIGHT: #FFCC33 1px solid; BORDER-TOP: #FFCC33 1px solid; FONT-SIZE: 10pt" value="<%=name%>">
                            &nbsp;</td>
                        </tr>
                        <tr> 
                          <td height="28" align="right">Region : </td>
                          <td> 
                            <input type="text" name="indicator" maxlength="50" size="23" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #FFCC33 1px solid; BORDER-LEFT: #FFCC33 1px solid; BORDER-RIGHT: #FFCC33 1px solid; BORDER-TOP: #FFCC33 1px solid; FONT-SIZE: 10pt" value="<%=indicate%>">
                            &nbsp;</td>
                        </tr>
                        <tr> 
                          <td height="28" align="right">Department : </td>
                          <td> 
                            <input type="text" name="phone" maxlength="50" size="23" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #FFCC33 1px solid; BORDER-LEFT: #FFCC33 1px solid; BORDER-RIGHT: #FFCC33 1px solid; BORDER-TOP: #FFCC33 1px solid; FONT-SIZE: 10pt" value="<%=phone%>">
                          </td>
                        </tr>
                        <tr align="center"> 
                          <td height="48" colspan="2"> 
                            <input type="button" value="Update" onclick="check();" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 10pt; HEIGHT: 20px; WIDTH: 80px">
                            <input type="reset" value="Reset" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 10pt; HEIGHT: 20px; WIDTH: 80px">
                            <input type=hidden value="muinfo" name=whatdo>
                            <input type="button" value="Change Password" onClick="window.location='user_pwd.asp';" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 10pt; HEIGHT: 20px; WIDTH: 100px">
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
</BODY></HTML>