<%
if session("shell_power")="" then
  response.redirect "index.asp"
end if
id=trim(request("id"))
if id="" then
 response.redirect "index.asp"
end if
%>
<!--#include file="images/conne.inc" -->
<%
sql="SELECT * FROM member "
sql=sql&" where id="&id

'response.write sql
'response.end
set rs=conn.execute(sql)
if not rs.eof then
  name=rs("name")
  flage=rs("flage")
  indicator=rs("indicate")
  employeenum=rs("employeenum")
  phone=rs("phone")
end if
rs.close
set rs=nothing
%>
<HTML><HEAD><TITLE>Member Detail</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="img/home.css" type="text/css">
</HEAD>
<BODY bgColor=#ffffff align=center>
<table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="#99CCCC">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
        <tr>
          <td>
            <table width="100%" border="0" cellspacing="3" cellpadding="3" align=center>
              <tr bgcolor="#006699"> 
                <td align="center" colspan="2" bgcolor="#006699"> <font color="#FFFFFF"><b>Member 
                  Information</b></font></td>
              </tr>
              <tr align="center"> 
                <td colspan="2">
                  <table width="100%" border="0" cellspacing="2" cellpadding="0">
                    <tr> 
                      <td height="23"> 
                        <table width="100%" border="0" cellspacing="2" cellpadding="0">
                          <tr> 
                            <td width="38%">Name</td>
                            <td bgcolor="#E6EBEF"> 
                              <%response.write server.htmlencode(name)%>
                            </td>
                          </tr>
                          <tr> 
                            <td width="38%">E-mail</td>
                            <td bgcolor="#E6EBEF"> 
                              <%response.write server.htmlencode(employeenum)
							if flage=0 then
							  response.write " ( <font color=blue>User</font> )"
							elseif flage=3 then
							  response.write " ( <font color=red>Administrator</font> )"
							elseif flage=8 then
							  response.write " ( <font color=red>Super Administrator</font> )"
							end if
							%>
                            </td>
                          </tr>
<tr> 
                            <td width="38%">Contact Person</td>
                            <td bgcolor="#E6EBEF"> 
                              <%
							  if Not isnull(indicate) then
							  response.write server.htmlencode(indicate)
							  end if
							  %>
                            </td>
                          </tr>
<tr> 
                            <td width="38%">Telephone</td>
                            <td bgcolor="#E6EBEF"> 
                              <%
							  if Not isnull(indicate) then
							  response.write server.htmlencode(indicate)
							  end if
							  %>
                            </td>
                          </tr>
                          <tr> 
                            <td width="38%">Region</td>
                            <td bgcolor="#E6EBEF"> 
                              <%
							  if Not isnull(indicate) then
							  response.write server.htmlencode(indicate)
							  end if
							  %>
                            </td>
                          </tr>
                          <tr> 
                            <td width="38%">Department</td>
                            <td bgcolor="#E6EBEF"> 
                              <%
							  if Not isnull(phone) then
							  response.write server.htmlencode(phone)
							  end if
							  %>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#006699" align="center">
                <td colspan="2" bgcolor="#006699" align="right">&nbsp;</td>
              </tr>
              <tr> 
                <td align="center" colspan="2"><A href="javascript:window.close();">Close 
                  Window</A></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<div align="center"><script language=JavaScript src="img/copyright.js"></script></div>
</BODY></HTML>