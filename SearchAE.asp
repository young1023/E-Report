<% 
'*********************************************************************************
'NAME       : SearchAE.asp          
'DESCRIPTION: Search and filter client number and used for all reports
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 090710 Roger Wong   Prototype
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

<SCRIPT language=JavaScript>
<!--

function AssignValue(){
	
  	myString = this.myform.myselect.value;
 	  self.opener.document.fm1.AETo.value=myString;
 	  self.opener.document.fm1.AEFrom.value=myString;
 		
 		self.close();
 	
}

//-->
</SCRIPT>


</HEAD>

<%
'define query
Const MaxNumberRealTimeFilter = 1000

Dim Search_AECode 
Dim Search_AEGroup
Dim iRecordCount

Search_AECode = ""
Search_AEGroup = ""
Select Case Session("shell_power")
	case "1"
		'AE shall access their own clients only
		Search_AECode = Session("id")
		'Search_AEGroup = Session("GroupID")
		
  case "5"
		' Branch Manager shall access all AE's clients belongs to 

		Search_AEGroup = Session("GroupID")
		
		'Others having full access

end select




%>




<%
	'*******
	' Start for user other than AE and branch manager
	'*******
	'if Session("shell_power") <> "1" and Session("shell_power") <> "5" then 
	if 1=1 then
				Dim Search_keyword

				if (Request.Form("keyword").Count = 0)  then
					 Search_keyword	    = ""
				else
					
					Search_keyword	    = Request.form("keyword")
				end if


				strURL = Request.ServerVariables("URL") ' Retreive the URL of this page from Server Variables
%>	


					<BODY OnLoad="document.myform.keyword.focus();document.myform.keyword.select();">
								<FORM NAME="myform"  method="post"  action="<%= strURL %>?sid=<%=SessionID%>">
								  <table width="99%" border="0" class="normal">
			
							
										<tr><td class="common"> 
											
											Enter AE Code or AE's Name
										</td></tr>
										<tr><td class="common"> 
										<INPUT name="keyword" value="<%= Search_keyword %>">
											<INPUT TYPE=submit value="Filter">
											<INPUT TYPE=button value="Clear">

										</td></tr>

			

								
<%							
			'if 0 = 1 then 
			If (Request.Form("keyword").Count = 0) Then
			'**********
			' If no argument
			'**********
			
			'do nothing
			  
			  
			else 
				set Rs1 = server.createobject("adodb.recordset")
			 Rs1.open ("exec Retrieve_AECode '"&Search_AECode&"', '"&Search_AEGroup&"', N'"&Search_keyword&"' ") ,  StrCnn,3,1
			' Rs1.open ("select top 10 * from client") ,  StrCnn,3,1
			
			
				iRecordCount = rs1(0) 'total number of records
			
			'response.write ("exec Retrieve_ClientNumber '"&Search_AECode&"', '"&Search_AEGroup&"'")
			'response.write "<BR>"
			
			'response.write  Session("ID") 
			'response.write "<BR>"
			'response.write  Session("GroupID") 



			 if iRecordCount <= 0 then
				
					'no record found
					response.write ("<tr><td>No record found</td></tr>")
				
				else
					'record found
					
					'response.write iRecordCount 
					
					'cal total no of pages
					
					
					'move to next recordset
			  	Set rs1 = rs1.NextRecordset() 
	
	%>		 
								<tr><td class="common">
								Select client (Total <%=iRecordCount%> clients)
								</td></tr>
								
								
								<tr><td class="common">
								
								<SELECT NAME="myselect" SIZE=10 class="common">
									<% do while (  Not rs1.EOF) %>
												<Option value="<%=rs1("loginname")%>"> <% response.write rs1("loginname") + " : " +  rs1("name")%> </OPTION>
			
															
									<%
										rs1.movenext 
										loop 
									%>
								</SELECT>
								
								
								
								<SCRIPT TYPE="text/javascript">
								<!--
								myClientNumber = ""
								var myfilter = new filterlist(document.myform.myselect);
								//-->
								</SCRIPT>
							</td></tr>
							
							<tr><td class="common">
													<INPUT TYPE=button onClick="AssignValue();" value="Select Client">
							</td></tr>
			<%
					'record found end if
					end if
						
			'argument end if						
			end if 
			%>
			
			
								
								
								</table>
								</FORM>
							

					</BODY>
					</HTML>

			<% 

		end if
			%>
	<%
	'*******
	' END  for user other than AE and branch manager
	'*******
	%>