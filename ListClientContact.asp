<% 
'*********************************************************************************
'NAME       : ListClientContact.asp          
'DESCRIPTION: A pop up to list client contact information
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 090820 Roger Wong   Prototype
'MODIFIED   : 
'********************************************************************************

%>
<!--#include file="include/SessionHandler.inc.asp" -->



<HTML>
<HEAD>
	<link rel="stylesheet" type="text/css" href="include/uob.css" />

<TITLE>Client List</TITLE>

<!-- Load the javascript code -->
<SCRIPT TYPE="text/javascript" SRC="include/filterlist.js"></SCRIPT>

</HEAD>




								
<%							
		'response.write "lsjflkdfj" & Request.QueryString("clientnumber")
			If (Request.QueryString("clientnumber").Count = 0) Then
					response.write "Unexpected error"
			'**********
			' If no argument
			'**********
			
			'do nothing
			  
			  
			else 
				Dim Search_clientnumber
				Search_clientnumber	    = Request.QueryString("clientnumber")
				set Rs1 = server.createobject("adodb.recordset")
			 	Rs1.open ("exec ListClientContact '"&Search_clientnumber&"'") ,  StrCnn,3,1
			
			
				'iRecordCount = rs1(0) 'total number of records
			
			'response.write ("exec Retrieve_ClientNumber '"&Search_AECode&"', '"&Search_AEGroup&"'")
			'response.write "<BR>"
			
			'response.write  Session("ID") 
			'response.write "<BR>"
			'response.write  Session("GroupID") 

		


					'record found
					
					'response.write iRecordCount 
					
					'cal total no of pages
					
					
					'move to next recordset
			  	'Set rs1 = rs1.NextRecordset() 
	
	%>		 
	

								  <table width="99%" border="0" class="normal">
									
										<% if not rs1.eof then%>
											<tr>
												<td class="common"><b> Client contact</b></td>
												<td class="common"> </td>
											</tr>
											<tr>
												<td class="common"> Client number:</td>
												<td class="common"><%=Search_clientnumber%> </td>
											</tr>
	

												<% do while (  Not rs1.EOF) %>
																<tr>
																	<td class="common"> <%=rs1("FieldName")%>: </td>
																	<td class="common"> <%=rs1("FieldValue")%> </td>
																</tr>
										<% 
													rs1.movenext
													loop 
											else
										%>
											<tr>
												<td class="common">Unexpected error</td>
												<td class="common"> </td>
											</tr>
										<%
											end if		
										%>
									</table>
																				

								
			<%

						
			'argument end if						
			end if 
			%>
			


								</BODY>
					</HTML>
				


