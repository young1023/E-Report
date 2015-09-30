

<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


Title = "Add New Member"

%>



<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>UOB Intranet</title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<SCRIPT language=JavaScript>
<!--

function getRandomNum(lbound, ubound) {
return (Math.floor(Math.random() * (ubound - lbound)) + lbound);
}
function getRandomChar(number, lower, upper, other, extra) {
var numberChars = "0123456789";
var lowerChars = "abcdefghijklmnopqrstuvwxyz";
var upperChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
var otherChars = "`~!@#$%^&*()-_=+[{]}\\|;:'\",<.>/? ";
var charSet = extra;
if (number == true)
charSet += numberChars;
if (lower == true)
charSet += lowerChars;
if (upper == true)
charSet += upperChars;
if (other == true)
charSet += otherChars;
return charSet.charAt(getRandomNum(0, charSet.length));
}
function getPassword(length, extraChars, firstNumber, firstLower, firstUpper, firstOther,
latterNumber, latterLower, latterUpper, latterOther) {
var rc = "";
if (length > 0)
rc = rc + getRandomChar(firstNumber, firstLower, firstUpper, firstOther, extraChars);
for (var idx = 1; idx < length; ++idx) {
rc = rc + getRandomChar(latterNumber, latterLower, latterUpper, latterOther, extraChars);
}
return rc;
}


function dosubmit(){
 document.addm.action="execute.asp?sid=<%=SessionID%>";
 if (document.addm.LoginName.value == "")
  {
   alert("Please enter the login name.");
   document.addm.LoginName.focus();
   return false;
  }
   if (document.addm.UserName.value == "")
  {
   alert("Please enter the user name.");
   document.addm.UserName.focus();
   return false;
  }

 if (document.addm.Password.value == "")
  {
   alert("Please enter the password.");
   document.addm.Password.focus();
   return false;
  }
  
document.addm.submit();
}
//-->
</SCRIPT>


</head>
<body leftmargin="0" topmargin="0" OnLoad="document.addm.LoginName.focus();">

<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">
<%
'-----------------------------------------------------------------------------
'
'      Main Content of the page is inserted here
'
'-----------------------------------------------------------------------------

%>
<form name=addm method=post>
   <table width="100%" border="0" cellpadding="1" cellspacing="1" class="normal">
                                    
                                      <tr> 
                                        <td width="38%" height="23" align="right">
                                      Login Name</td>
                                        <td> 
                                          <input type="text" name="LoginName" size="35"></td>
                                      </tr>
                                    
                                      <tr> 
                                        <td width="38%" height="23" align="right">
                                      User Name:                                    <td> 
                                          <input type="text" name="UserName" size="35">
                                        </td>
                                      </tr>
                                      <tr> 
 <td width="43%" height="23" align="right">Department</td>
                                        <td> 
                                          <input type="text" name="Dept" size="35"></td>
                                      </tr>
                                      <tr> 
 <td width="43%" height="23" align="right">Password:</td>
                                        <td> 
                                          <input name="Password" size="35" type="text" autocomplete="off" maxlength="20" readonly>
                                          <input type="button" value="Generate password " onClick="document.addm.Password.value=getPassword(10, '',true,false,false,false,false,true,true,false);" class="common">
                                          </td>
                                      </tr>

                                                             
                                      <tr> 
 <td width="43%" height="23" align="right">Email Address : </td>
                                        <td> 
  <input type="text" name="Email" size="35">
                                        </td>
                                      </tr>
<% 
           
           Lsql = " Select * from UserLevel where LevelNumber < > 0 order by LevelName"
           Set LRs = conn.execute(Lsql)
  
%>    
                                                             
                                      <tr>
<td width="38%" height="23" align="right">User Level : </td>
  <td>
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
				
                                      </tr>
    
                                                             

                                      <tr> 
                                        <td colspan="2" align="center"> 
    <input type="button" value="   Submit  " onClick="dosubmit();" class="common">
  <input type="reset" value="   Reset   " class="common">
 <input type="hidden" name="whatdo" value="added member">
<input type="hidden" name="PasswordLength" value="<% = PasswordLength %>">
                                        </td>
                                      </tr>
                                    </table>
              </div>
						  </td>
                        </tr>
                      </table>
			</form>



<%
'-----------------------------------------------------------------------------
'
'      End of the main Content 
'
'-----------------------------------------------------------------------------
%>
</div>
            
              </body>

              </html>