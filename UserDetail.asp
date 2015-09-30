<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if
%>


<%
if session("shell_power")="" then
  response.redirect "default.asp"
end if

Title = "User Information"

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
</head>
<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">
	
               <table width="98%" border="0" cellpadding="4" class="normal">
             
                      <tr> 
                        <td height="28" width="50%">Welcome <%=Session("Name")%>,
						</td>
                      </tr>
                     <tr>
                        <td>System Message: <% = Session("SystemMessage") %><td>
                      </tr>
                    <tr>
                        <td height="128"></td>
                      </tr>

                    </table>
            
</div>            
  
            
              </body>

              </html>
              
<%

'reset System message
Session("SystemMessage") = ""
%>