<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if
%>


<% 
Title = "Changing Password" 


Username = cstr(session("id"))
OldPassword = cstr(request.form("OldPassword"))
NewPassword = cstr(request.form("NewPassword1"))
ipaddress  = Request.ServerVariables("REMOTE_ADDR")
strURL = Request.ServerVariables("URL") ' Retreive the URL of this page from Server Variables



%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />

<SCRIPT language=JavaScript>
<!--
    function dosubmit() {
        document.fm1.action = "<%= strURL %>?sid=<%=SessionID%>";
        document.fm1.whatdo.value = "changed password";

        if (document.fm1.OldPassword.value == "") {
            alert("Please enter your old password.");
            document.fm1.OldPassword.focus();
            return false;
        }

        if (document.fm1.NewPassword1.value == "") {
            alert("Please enter your new password.");
            document.fm1.NewPassword1.focus();
            return false;
        }
               

        if (document.fm1.NewPassword2.value == "") {
            alert("Please confirm your new password.");
            document.fm1.NewPassword2.focus();
            return false;
        }
        
           if (document.fm1.OldPassword.value == document.fm1.NewPassword1.value) {
            alert("The new password cannot be the same as old password!");
            document.fm1.NewPassword1.focus();
            return false;
        }

        if (document.fm1.NewPassword1.value != document.fm1.NewPassword2.value) {
            alert("New passwords do not match!");
            document.fm1.NewPassword2.focus();
            return false;
        }

        document.fm1.submit();
    }
//-->
</SCRIPT>
<script type="text/javascript" src="include/external.js"></script>
</head>

<body leftmargin="0" topmargin="0"  OnLoad="document.fm1.OldPassword.focus();">

<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">
  
<%          
	' If user enter id/pw, exec store procedure to change the password
	' --------------------------------------------------------------------------------------------------------------------	

'response.write ("Exec changepwd '"&Username&"' , '"&OldPassword&"' , '"&NewPassword&"' , '"&ipaddress&"' ")

	if (Username <> "" and OldPassword <> "" and NewPassword <> "") then
		
            'Response.Write ("Exec changepwd '"&Username&"' ,'"&Username&"' , '"&OldPassword&"' , '"&NewPassword&"', '"&ipaddress&"' ") 
			set Rs1 = server.createobject("adodb.recordset")
			Rs1.open ("Exec changepwd '"&Username&"' ,'"&Username&"' , '"&OldPassword&"' , '"&NewPassword&"', '"&ipaddress&"' ") ,  StrCnn,3,1
		
		
			If (not Rs1.EoF) Then
					if (cint(rs1("memberid") <=1)) then
		
		
							
							'Password change failed
							'''''''''''''''''''''''
							Session("SystemMessage") = rs1("Message")
							
							'Message = "login failed using invalid user name <b>" & num  &  "</b>  from " &  Request.ServerVariables("REMOTE_ADDR")
							'Conn.Execute ("Exec AddLog '"&Message&"', "&num&"") 	
							
							'Call myclose()
							Response.Redirect strURL & "?sid=" &SessionID
					else
					
							'Password change success
							'''''''''''''''
							Session("SystemMessage") = rs1("Message")
							
							'Call myclose()
		   				Response.Redirect strURL & "?sid=" &SessionID
					End if
				end if
	end if
%>


<form name="fm1" method="post" action="">
		<table border="0" width="90%" class="Normal" id="table1" height="200">

				<tr>
						<td width="26%">Old Password</td>
						<td width="71%"> 
						<input type=password name="OldPassword" size=25 maxlength="30" onBlur="validate(this)"></td>
				</tr>
				<tr>
						<td width="26%">New Password</td>
						<td width="71%"> 
						<input type=password name="NewPassword1" size=25 maxlength="30" onBlur="validate(this)"></td>
				</tr>
				<tr>
						<td width="26%"> Confirm new password</td>
						<td width="71%"> 
						<input type=password name="NewPassword2" size=25 maxlength="30">&nbsp;
						<input type="button" value=" Submit" class="common" onclick="dosubmit();">
						<input type="hidden" name="whatdo" value=""></td>
				</tr>
				

				<tr>
						<td colspan="2" class="RedClr"><%=Session("SystemMessage")%></td>
				</tr>
		</table>
</form>
<% 
	Session("SystemMessage") = ""
	user=""
	password=""
    

%>

              </body>
              </html>