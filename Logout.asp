<%



' Title = "Logout"
' session("shell_power")= 0
' session("u_enum")=""
' session("u_vnum")=""
' session("u_vtime")=""
%>

<% 
Dim LogoutReason 
LogoutReason   = Request.QueryString("r").item

Select case LogoutReason

case 1 :
		 	Session("SystemMessage") = "Session Expired. User logged out automatically"

case else:
			Session("SystemMessage") = ""

end Select
 
		
%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Refresh" content="2; url='default.asp'">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
</head>

<body leftmargin="0" topmargin="0">
<!--#include file="include/SQLConn.inc.asp" -->

<!-- #include file ="include/Master.inc.asp" -->



<div id="Content">
<div align="center">
<table border=0 cellpadding=3 cellspacing=0 class="Normal" width="40%" height="50%">
   <tr align="center"> 
    <td height="28" class="RedClr"><% = Session("SystemMessage") %>
    </td>
  </tr>  
  <tr> 
    <td align=center><b>Logout 
      Successfully</td>
    </tr>
    
  </table>

</div>

</div>
</body>
</html>
<%
	num = 	Session("MemberID") 
	ipaddress  = Request.ServerVariables("REMOTE_ADDR")
	' temrinate session in DB
	set Rs1 = server.createobject("adodb.recordset")
	Rs1.open ("Exec Terminate_UserSession '"&num&"' , '"&ipaddress&"' ") ,  StrCnn,3,1
	
	'response.write "Exec Terminate_UserSession '"&num&"' , '"&ipaddress&"' "

	'Terminate all session
	Session.Abandon 



	Session("SystemMessage") = ""

 if session("shell_power")="" then
  ' response.redirect "Default.asp"
 end if
%>
