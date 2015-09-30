
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


Title = "Member Detail"

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
function doSubmit(){
  document.fm1.action="execute.asp?sid=<%=sessionid%>";
 
  document.fm1.whatdo.value="updated member";


if (document.fm1.Name.value == "")
  {
   alert("Please enter user name.");
   document.fm1.Name.focus();
   return false;
  }
  

document.fm1.submit();
}

function changePassword(){
if (document.fm1.Password.value == "")
  {
   alert("Please enter password.");
   document.fm1.Name.focus();
   return false;
}

if (document.fm1.ConfirmPassword.value == "")
  {
   alert("Please enter the confirmation password.");
   document.fm1.ConfirmPassword.focus();
   return false;
}

if (document.fm1.Password.value != document.fm1.ConfirmPassword.value) {
            alert("The Password and confirm password do not match!");
            document.fm1.ConfirmPassword.focus();
            return false;
}

  document.fm1.action="<%=strURL%>?sid=<%=SessionID%>";
 document.fm1.submit();
}

//-->
</SCRIPT>
<script type="text/javascript" src="include/external.js"></script>
</head>
<body leftmargin="0" topmargin="0" OnLoad="document.fm1.Name.focus();">

<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">

  
<%


MemberID = Request("MemberID")

If MemberID <> "" Then

       sql="select Name, email, loginname, dept, userlevel, lock from Member where MemberID ="&MemberID
       Set Rs=conn.execute(sql)
       
Name  =  Trim(Rs("Name"))
Email = Trim(Rs("Email"))
LoginName = Trim(Rs("LoginName"))
Department = Trim(Rs("Dept"))
UserLevel = Trim(Rs("UserLevel"))
'Password = Trim(Rs("Password"))
Lock = Trim(Rs("Lock"))
ID = session("id")
End If       

%>


<%          
' Handle changing password

'Username = Request.form("Name")
ConfirmPassword = cstr(request.form("ConfirmPassword"))


	if (ConfirmPassword <> "") then
		
			set Rs1 = server.createobject("adodb.recordset")
			Rs1.open ("Exec changepwd '"&ID&"' ,'"&LoginName&"' , '"&ConfirmPassword&"' , '"&ConfirmPassword&"', '"&ipaddress&"' ") ,  StrCnn,3,1
			'Rs1.open ("Exec changepwd '' ,'' , '' , '', '' ") ,  StrCnn,3,1
			'response.write ("Exec changepwd '"&ID&"' ,'"&LoginName&"' , '"&ConfirmPassword&"' , '"&ConfirmPassword&"', '"&ipaddress&"' ") 
		
		
			If (not Rs1.EoF) Then
					if (cint(rs1("memberid") <=1)) then
		
		
							
							'Password change failed
							'''''''''''''''''''''''
							Session("SystemMessage") = rs1("Message")
							
							'Message = "login failed using invalid user name <b>" & num  &  "</b>  from " &  Request.ServerVariables("REMOTE_ADDR")
							'Conn.Execute ("Exec AddLog '"&Message&"', "&num&"") 	
							
					else
					
							'Password change success
							'''''''''''''''
							Session("SystemMessage") = rs1("Message")
					End if
				end if
	end if
%>

<form name="fm1" method="post" action="">

  <table width="99%" border="0" class="normal">
    <tr> 
      <td colspan="2" class="BlueClr">　</td>
    </tr>
    <tr> 
      <td colspan="2">
<font size="4">Member Detail</font></td>  
    </tr>
    <tr> 
      <td colspan="2"></td>
    </tr>
  <tr> 
      <td colspan="2"  align="right">
<font color="red">*</font>	 Denotes a mandatory field</td>
    </tr>
    <tr> 
      <td colspan="2" bordercolor="#999966" class="ValidHead">　</td>
    </tr>
    <tr> 
      <td width="23%"></td>
      <td width="76%">
</td>
</tr>



      <tr>

      <td height="28">Login ID&nbsp;</td>  
      <td valign="bottom" height="28">
<% = LoginName %></td>
    </tr>
    


	  <td><font color="red">*</font>
	 User Name:</td> 
      <td valign="bottom">
<input name="Name" type=text value="<% = Name %>" size="60">

      </td>
<input name="Name2" type=hidden value="<% = Name %>" size="60">
	<tr>

      <td>Department</td>
      <td valign="bottom">

<input name="Department" type=text value="<% = Department %>" size="60"></td>
<input name="Department2" type=hidden value="<% = Department %>" size="60">
    	</tr>



	<tr>

      <td>Email</td>
      <td valign="bottom">

<input name="Email" type=text value="<% = Email %>" size="60"></td>
<input name="Email2" type=hidden value="<% = Email %>" size="60">
    	</tr>
    	
<% 
           
           Lsql = " Select * from UserLevel where LevelNumber < > 0 order by LevelName"
           Set LRs = conn.execute(Lsql)
  
%>


	<tr>

      <td height="18">User Level</td>
      <td valign="bottom" height="18">				    	
<select name="UserLevel" class="common"  size="1">
          <% 
                             If Not LRs.EoF Then
                        LRs.MoveFirst
							do while not LRs.eof
                              if trim(UserLevel) = trim(LRs("LevelNumber")) then
                                 response.write "<option value="&LRs("LevelNumber")&" selected>"&trim(LRs("LevelName"))&"</option>"
                                 else
                                 response.write "<option value="&LRs("LevelNumber")&">"&trim(LRs("LevelName"))&"</option>"
                               end if
                               LRs.movenext
							loop
						
						End if
					%>
        </select>
				
				
				</td>
<input name="LevelName2" type=hidden value="<% = UserLevel %>" size="6">
    </tr>


<% If UserLevel > 7 Then %>
	<tr>

      <td height="18">Received System Email</td>
      <td valign="bottom" height="18">
                                        Yes
										<input type="radio" value="Yes" name="SystemEmail"> 
										No 
										<input type="radio" value="No" checked name="SystemEmail">&nbsp;</td>
    </tr>
<% End If %>


	<tr>

      <td height="18">Unlock Account</td>
      <td valign="bottom" height="18">
                                        Lock
										<input type="radio" value="3" <% If Lock = 3 Then %>checked<% End If %> name="Lock"> 
										unlock 
										<input type="radio" value="0" <% If Lock = 0 Then%>checked<% End If %> name="Lock"> 
</td>
<input name="Lock2" type=Hidden value="<% = Lock %>" size="6">
    </tr>

	<tr>

      <td height="18"></td>
      <td valign="bottom" height="18">
    <input type="button" value="   Save  " onClick="doSubmit();" class="common"></td>
    	<input type="hidden" name="whatdo" value="">
      		<input type="hidden" name="MemberID" value="<% = MemberID %>">
    </tr>
    
	<tr>

      <td height="18"></td>
      <td valign="bottom" height="18">
    	</td>
    </tr>
    
	<tr>

      <td height="18"></td>
      <td valign="bottom" height="18">
    	</td>
    </tr>
    
 	<tr>

      <td>New Password</td>
      <td valign="bottom">

<input name="Password" type=password value="" size="60"">&nbsp;
										</td>
    	</tr>

	<tr>

      <td> Confirm password</td>
      <td valign="bottom">

<input name="ConfirmPassword" type=password value="" size="60"></td>
    	</tr>
	<tr>

      <td height="18"></td>
      <td valign="bottom" height="18">
    <input type="button" value="   Change Password  " onClick="changePassword();" class="common"></td>
    </tr>
				<tr>
						<td colspan="2" class="RedClr"><%=Session("SystemMessage")%></td>
				</tr>

                </table>
                </form>
                
               
      </div>
              </body>

              </html>
<%              								Session("SystemMessage") = "" %>