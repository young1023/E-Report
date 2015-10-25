
<%

Title = "Intranet"
session("shell_power")= 0
'Session("SystemMessage") = ""

'for each x in Session.Contents
'  Response.Write(x & "=" & Session.Contents(x) & "<br />")
'next
%>
<html>
<head>
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<script language="JavaScript">
    function dosubmit() {
        document.fm1.action = "login_session.asp";
        document.fm1.whatdo.value = "changes password";

        if (document.fm1.num.value == "") {
            alert("Please enter your login name.");
            document.fm1.num.focus();
            return false;
        }

        if (document.fm1.inkey.value == "") {
            alert("Please enter your password.");
            document.fm1.inkey.focus();
            return false;
        }

      document.fm1.submit();
    }
    
//-->
</script>
<script type="text/javascript" src="include/external.js"></script>
</head>

<body leftmargin="0" topmargin="0" OnLoad="document.fm1.num.focus();">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">



  <TABLE border=0 cellPadding=0 cellSpacing=0 height=408 width="99%" class="Normal">
    <TBODY> 
    <TR> 
      <TD vAlign=center width=99%> 
        <DIV align=center> 
          <div align="center">
          <table border=0 cellpadding=0 cellspacing=0 width="80%" class="Normal">
            <tr> 
              <td height="28" align=center>
              <b>Member Login</b></td>
            </tr>
            <tr> 
              <td height=56 align=center>
                <table border=0 cellpadding="2" width="60%" class="Normal">
				<form name=fm1 method=post action="">
                  <tr> 
                    <td align=center height="28" width="45%"><b><font color="#003366">User 
                      : </font></b></td>
                    <td> 
                        <input name=num  autocomplete=off size=20 maxlength="30" onBlur="validate(this)">
                    </td>
                  </tr>
                  <tr> 
                    <td align=center height="28" width="45%"><b><font color="#003366">Password 
                      : </font></b></td>
                    <td> 
                        <input type=password name=inkey size=20  autocomplete=off maxlength="30" onBlur="validate(this)">
                    </td>
                  </tr>
                  <tr align="center"> 
                    <td height="28" colspan="2">
                      <input type="button" value="    Login   " class="common" onclick="dosubmit();">
                        <input type=reset value="   Reset   " class="common">
                    	<input type="hidden" name="whatdo" value=""></td>
                  </tr>
                                    <tr align="center"> 
                    <td height="28" colspan="2">
					<p align="left"><a class="Common" href="forgotpassword.asp" style="text-decoration: none">Forgot Password?</a></td>
                  </tr>


				</form>
                </table>
                      </tr>
                          <tr align="center"> 
                    <td height="28" class="RedClr"><% = Session("SystemMessage") %>
                    </td>
                  </tr>
              </td>
            </tr>
          </table>
       
          <br>
          <table width="80%" border="0" cellspacing="1" cellpadding="1" class="Normal" style="border-style: solid">
            <tr> 
              <td align="center" colspan="2"><b>Declaration</b></td>
            </tr>
            <tr> 
              <td align="center" colspan="2">
								Access to electronic resources to this system is restricted to individuals 
								authorized by UOB or its affiliates. Use of this system 
								is subject to all policies and procedures set forth by UOB. 
								Unauthorized use is prohibited and may result in administrative or legal action. UOB 
								may monitor the use of this system for purposes related to security management, system 
								operations, and intellectual property compliance.
						</td>
            </tr>
          </table>
<% 
Session("SystemMessage") = "" 
Session.Abandon
%>
 
</div>
              </body>
              </html>