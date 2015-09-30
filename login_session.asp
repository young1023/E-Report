
<%
' Get Login and Password
num = replace(trim(request.form("num")),"'","''")
inkey = replace(trim(request.form("inkey")),"'","''")

' Reset System Message
Session("SystemMessage") = ""
%>
<!--#include file="include/SQLConn.inc.asp" -->


<%


	set Rs1 = server.createobject("adodb.recordset")
	Rs1.open ("Exec Check_Session '"&num&"' ") ,  StrCnn,3,1


	If (not Rs1.EoF) Then
			if (cint(rs1("memberid") <0)) then


					
					'Other session exist
					'''''''''''''''''''''''
					
					SessionExistFlag = 1

			else
			
					'No other session exist
					'''''''''''''''
					SessionExistFlag = 0
					'Response.Redirect "login.asp?sid=" & Session("SessionID")
 
			End if
	else
	
			'No other session exist
			'''''''''''''''
			SessionExistFlag = 0
			'Response.Redirect "login.asp?sid=" & Session("SessionID")


	end if




%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />

<SCRIPT language=JavaScript>
	
function SessionChecking(){

 {
  var msg = "Another concurrent session may be existed. This may be caused by improper log out process. Continue?";
  if (<%=SessionExistFlag%> == 0 || confirm(msg)==true)
   {
   	document.fm1.action =  "login.asp";

    document.fm1.submit();
   }
   else
   {
   	document.fm1.action =  "default.asp";
    document.fm1.submit();
   	
  	}
 }

}

//-->
</SCRIPT>

</head>


<body onload="SessionChecking()">
	
	<form name=fm1 method=post action="">

  <table width="99%" border="0" class="normal">
  	<tr>
             	<td><input type="hidden" name="num" value="<%=num %>"></td>
             		<td><input type="hidden" name="inkey" value="<%=inkey %>"></td>

  	</tr>
  </table>

  </form>
  


  
</body>
</html>