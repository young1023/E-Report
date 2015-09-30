<!--#include file="images/conne.inc" -->
<% 

if session("shell_power")="" then
  response.redirect "default.asp"
end if
%>
<HTML><HEAD><TITLE>E-Report</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="hse.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=4>
<DIV align=center> 
  <TABLE border=0 cellPadding=0 cellSpacing=0 height=100% width=99%>
    <TBODY> 
    <TR>
      <TD vAlign=top>
        
        <table width="100%" border="0" cellpadding="1" cellspacing="1" height="100%">
          <tr> 
            <td bgcolor="#000000">
              <table width="100%" border="0" cellpadding=4 bgcolor="#FFFFFF" height="100%" class="normal">
                  <tr>
                  
                  <td width="99%" align="center" valign="top" bgcolor="#E6EBEF">
                    <table width="80%"  border="0" cellpadding="4" class="normal">
                      <tr> 
                        <td height="48" align="center" colspan="4"><images src="images/back.gif" width="306" height="37"><b><font color="#FF6600">Login 
                          Information</font></b></td>
                      </tr>
                      <tr> 
                        <td height="28" width="20%">Indicator : </td>
                        <td height="28" width="20%">　</td>
                        <td align="left" colspan="2">&nbsp;<font color=blue><%=session("u_enum")%></font>
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
                        <td height="28">Last 
                          time login at:
                        </td>
                        <td height="28">　</td>
                        <td width="42%"></td>
                        <td width="36%"></td>
                      </tr>

                    </table>
                  </td>
                </tr>
              </table></td>
          </tr>
        </table>
      </td>
    </tr>
</TABLE>
</DIV>
</BODY></HTML>