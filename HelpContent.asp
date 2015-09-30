<% 
'*********************************************************************************
'NAME       : HelpContent.asp          
'DESCRIPTION: Search Help Content
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 091101 Gary Yeung
'MODIFIED   : 
'********************************************************************************

Search_Title = Request("Title")
Search_Keyword = Request("Search_Keyword")
%>
<!--#include file="include/SessionHandler.inc.asp" -->



<HTML>
<HEAD>
	<link rel="stylesheet" type="text/css" href="include/uob.css" />
<TITLE>Help Content</TITLE>
</HEAD>
<BODY >
<form name="fm1" method="post" action="HelpContent.asp?sid=<%=SessionID%>">
<table width="99%" border="0" class="normal">
   <tr>
<td class="common" align="Left"><a href="javascript:window.close();">Close This Help Windows</a>
</td>
<td class="common" align="right"> 
	Search: <INPUT name="Search_Keyword" value="<%= Search_Keyword %>">
	<INPUT TYPE=submit value="Search">
				</td></tr>
</table>
<br>
<table width="99%" border="0" class="normal" cellspacing="1" bgcolor="#808080">
		 <tr>
<td bgcolor="#FFFF99" width="153"> 
	Module</td>
<td bgcolor="#FFFF99" width="161"> 
	Section</td>
<td bgcolor="#FFFF99"> 
	Explanation</td></tr>
<%							
     Set Rs = server.createobject("adodb.recordset")  
     Rs.open ("exec Retrieve_HelpContent '"&Search_Title&"', '"&Search_Keyword&"'") ,  StrCnn,3,1
     'Response.write  ("exec Retrieve_HelpContent '"&Search_Title&"', '"&Search_Keyword&"'")
                If Not Rs.EoF Then
                      Rs.MoveFirst
                   Do While Not Rs.EoF
%>		 
<tr bgcolor="white">
<td width="153">
<% = Rs("MenuName") %></td>
<td width="161">
<% = Rs("SectionName") %></td>
<td>
<% = Rs("HelpContent") %></td>
</tr>
						
			<%
					'record found end if
					     Rs.MoveNext
                   Loop
                End If
						
			%>
			
			
								
								
								</table>
								</FORM>
							

					</BODY>
					</HTML>

	<%
	'*******
	' END  for user other than AE and branch manager
	'*******
	%>