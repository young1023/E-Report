<!--#include file="images/conne.inc" -->
<HTML><HEAD><TITLE></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="images/style.css" type="text/css">
<link href="hse.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=4 OnLoad="document.fm1.num.focus();">
<DIV align=center>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 height=20 width="99%">
 <TBODY>
  <TR>
    <TD align=middle vAlign=top>
      <TABLE border=0 cellPadding=0 cellSpacing=0 height=1 width="100%">
        <TBODY>
        <TR>
          <TD bgColor=#000000 height=1></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD>
        <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%" class="Normal">
          <TBODY> 
          <TR bgcolor="#006699"> 
            <TD height=18 width="180" align=center><script language=JavaScript src="images/date.js"></script>
            </TD>
            <TD align=right height=18 vAlign=center class="ValidTitle"></TD>
          </TR>
          </TBODY> 
        </TABLE>
    </TD>
   </TR>
  </TBODY>
</TABLE>
  <TABLE border=0 cellPadding=0 cellSpacing=0 height=408 width="99%" class="Normal">
    <TBODY> 
    <TR> 
      <TD vAlign=center width=99%> 
        <DIV align=center> 
          <div align="center">
          <table border=0 cellpadding=0 cellspacing=0 width="80%" class="Normal">
            <tr> 
              <td height="28" align=center ><b>

 
<script language="JavaScript">
NS4 = (document.layers) ? true : false;

function checkEnter(event)
{ 	
	var code = 0;


	if (NS4)
		code = event.which;
	else
		code = event.keyCode;
	if (code==13)
       document.fm1.submit();

document.fm1.submit();
}

              </script>

  
                Member Login </b></td>
            </tr>
            <tr> 
              <td height=56 align=center>
                <table border=0 cellpadding="2" width="60%" class="Normal">
				<form name=fm1 method=post action="login.asp">
                  <tr> 
                    <td align=center height="28" width="45%"><b><font color="#003366">User 
                      : </font></b></td>
                    <td> 
                        <input name=num size=20 maxlength="30">
                    </td>
                  </tr>
                  <tr> 
                    <td align=center height="28" width="45%"><b><font color="#003366">Password 
                      : </font></b></td>
                    <td> 
                        <input type=password name=password size=20 maxlength="30">
                    </td>
                  </tr>
                  <tr align="center"> 
                    <td height="28" colspan="2">
                      <input type=submit value="    Login   " name="button" class="common">
                        <input type=reset value="   Reset   " class="common">
                    </td>
                  </tr>
                                    <tr align="center"> 
                    <td height="28" colspan="2">
					<p align="left"><a class="Common" href="forgotpassword.asp" style="text-decoration: none">Forgot Password</a>?</td>
                  </tr>
				</form>
                </table>
                      </tr>
                          <tr align="center"> 
                    <td height="28">
                    </td>
                  </tr>
              </td>
            </tr>
          </table>
          </div>
          <br>
          <table width="80%" border="0" cellspacing="1" cellpadding="1" class="Normal" style="border-style: solid">
            <tr> 
              <td align="center" colspan="2"><b>Declaration</b></td>
            </tr>
            <tr> 
              <td align="center" colspan="2">It is illegal to get access the system without the permission of the authorized user.</td>
            </tr>
          </table>
        </DIV>
      </TD>
     
    </tr>
    <TR>
      <TD colSpan=5 height=11 align=center>
        <script language=JavaScript src="images/copyright.js"></script>
<%
conn.close
set conn=nothing
%>
      </TD>
    </TR>
</TABLE>
</DIV>
</BODY></HTML>