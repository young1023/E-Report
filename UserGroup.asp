

<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


Title = "User Group"

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
<%
'-----------------------------------------------------------------------------
'
'      Main Content of the page is inserted here
'
'-----------------------------------------------------------------------------
id = Request("ID")

FunctionID = Request("FunctionID")

If FunctionID = "" Then

FunctionID = "edit"

End IF

if id <> "" then

         sql = " Select * from UserGroup where GroupID="&id
         set rs = conn.execute(sql)

Name = rs("Name")
Description = rs("Description")
SharingGroup = rs("Sharing")

'response.write SharingGroup
			
			'display shared group (from table "sharedgroup") or non-shared group (from table "member")
			if SharingGroup = 0 then
					
					set RsMemberInclude = server.createobject("adodb.recordset")
					RsMemberInclude.open ("Exec ListAEMemberInclude "&id&" ") ,  StrCnn,3,1
					
					set RsMemberExclude = server.createobject("adodb.recordset")
					RsMemberExclude.open ("Exec ListAEMemberExclude "&id&" ") ,  StrCnn,3,1
			
			else
			
					set RsMemberInclude = server.createobject("adodb.recordset")
					RsMemberInclude.open ("Exec List_SharedGroupAEMemberInclude "&id&" ") ,  StrCnn,3,1
					
					set RsMemberExclude = server.createobject("adodb.recordset")
					RsMemberExclude.open ("Exec List_SharedGroupAEMemberExclude "&id&" ") ,  StrCnn,3,1
			
			end if

end if
%>



<SCRIPT language=JavaScript>
<!--
function dosubmit(){
 document.fm1.action="execute.asp?sid=<%=SessionID%>";
 document.fm1.whatdo.value="added group";
 if (document.fm1.Name.value == ""){
  	alert("Please enter the group name!");
            document.fm1.Name.focus();
            return false;
			}

document.fm1.submit();
}

function doReturn(){
document.fm1.action="sa_group.asp?sid=<%=SessionID%>";
document.fm1.submit();
}

<%
	' For non-shared group  
	if SharingGroup = 0 then 
	 
	%>

														function AddMember(){
														document.fm1.action="execute.asp?sid=<%=SessionID%>";
														document.fm1.whatdo.value="add member";
														if (document.fm1.Name.value == ""){
														  	alert("Please enter the group name!");
														            document.fm1.Name.focus();
														            return false;
																	}
														
														document.fm1.submit();
														}
														
														
														function RmMember(){
														document.fm1.action="execute.asp?sid=<%=SessionID%>";
														document.fm1.whatdo.value="remove member";
														if (document.fm1.Name.value == ""){
														  	alert("Please enter the group name!");
														            document.fm1.Name.focus();
														            return false;
																	}
														
														
														document.fm1.submit();
														}
														
														
														

<%
	' For shared group   
	' action value changed
	else 
	
	%>
														function AddMember(){
														document.fm1.action="execute.asp?sid=<%=SessionID%>";
														document.fm1.whatdo.value="add shared member";
														if (document.fm1.Name.value == ""){
														  	alert("Please enter the group name!");
														            document.fm1.Name.focus();
														            return false;
																	}
														
														document.fm1.submit();
														}
														
														
														function RmMember(){
														document.fm1.action="execute.asp?sid=<%=SessionID%>";
														document.fm1.whatdo.value="remove shared member";
														if (document.fm1.Name.value == ""){
														  	alert("Please enter the group name!");
														            document.fm1.Name.focus();
														            return false;
																	}
														
														
														document.fm1.submit();
														}
														
	
<% end if%>


														function dosave(){
														document.fm1.action="execute.asp?sid=<%=SessionID%>";
														document.fm1.whatdo.value="updated group";
														if (document.fm1.Name.value == ""){
														  	alert("Please enter the group name!");
														            document.fm1.Name.focus();
														            return false;
																	}
														
														
														document.fm1.submit();
														}
//-->
</SCRIPT>

  <table width="80%" border="0" class="normal">

    <tr> 
      <td  class="BlueClr" width="30%">
      <a href="UserGroup.asp?sid=<% =SessionID %>&functionid=edit&id=<% =ID %>">Edit
      </a>
      </td> <td  class="BlueClr" width="30%">
      <a href="UserGroup.asp?sid=<% =SessionID %>&functionid=add&id=<% =ID %>">Add/Remove User
      </a></td> 
     
    </tr>

   </table>
    
 <br>


<form name="fm1" method="post" action="">




  <% If FunctionID <> "edit" Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>
  	
  <table width="90%" border="0" class="normal">

    <tr> 
      <td colspan="2" class="BlueClr"></td> 
    </tr>
    
    
    <tr> 
      <td colspan="2"></td>
    </tr>
    
    
		<tr> 
				<td colspan="2"  align="right">
				<font color="red">*</font>	 Denotes a mandatory field</td>
		</tr>
    
    
		<tr> 
				<td colspan="2"  align="right">
				</td>
		</tr>
    
		<tr> 
				<td>　</td> 
				<td>	</td>
		</tr>
    
		<tr> 
				<td>
				<font color="red">*</font>Group Name</td> 
				<td> 	     
				<Input name="Name" type=text value="<% = Name %>" size="50">
				<Input type="hidden" name="OldName" value="<% = Name %>" size="30">
				</td>
		</tr>
    
		<tr> 
				<td></td> 
				<td></td>
		</tr>

		<tr> 
				<td>
				Description </td> 
				<td>
				<input name="Description" type=text value="<% = Description %>" size="50"></td>
				<Input type="hidden" name="OldDesc" value="<% = Description %>" size="30">
		</tr>
				<tr> 
				<td>
				Shared Group</td> 
				<td>
						
								<select size="1" name="SharingGroup" class="common">
								<option selected value="0" <% If SharingGroup = 0 Then %>selected<% End If %>>No</option>
								<option value="1" <% If SharingGroup = 1 Then %>selected<% End If %>>Yes</option>
								</select></td>
			<Input type="hidden" name="OldSharing" value="<% = SharingGroup %>" size="30">	
		</tr>
		<tr> 
				<td></td>
				<td>
				<% if id <> "" then %>
				<input type="button" value="   Modify and Save  " onClick="dosave();">&nbsp;
				<input type="button" value="   Return  " onClick="doReturn();">&nbsp;
				
				<input type=hidden name=id value='<% = id %>'>
				
				</td>
		</tr>
	</table>
	
	</div>
	
	
	
 <% If FunctionID <> "add" Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>
	
 	<table width="90%" border="0"  height="300">

 <tr> 
	<td>
		<table width="100%" border="0" class="normal" height="300">
 			<tr> 
      <td colspan="2" class="BlueClr"></td> 
    </tr>


		<tr> 
				<td colspan="2"  ><b>Member section: <p></b> </td> 

		</tr>


		

    <tr>
		    <td> Current members </td>
				
				<td >
						<% if (  RsMemberInclude.RecordCount > 0) then %>
						    	
						    	<select size="10" name="removemember" class="common">
									<%
											do while (  Not RsMemberInclude.EOF)
									%>
											<option value="<%=RsMemberInclude("memberid")%>" ><% response.write "(" & RsMemberInclude("loginname") & ") " & RsMemberInclude("name")%></option>
									
									<%
											RsMemberInclude.movenext
											Loop
									%>
									</select>
									
						<% else %>
						
						No member available 
						
						<% end if %>
				</td>
		</tr>    

				<% if (  RsMemberInclude.RecordCount > 0) then %>
				<tr>
						<td>&nbsp;</td>
						<td > <input type="button" value="   Remove user  " onClick="RmMember();">&nbsp;
					</td>
				</tr>
		<% end if %>

		</table>
      		</td>
      <td width="50%">
<table width="100%" border="0" class="normal" height="300">
<tr> 
				<td colspan="2"  >&nbsp;<p></td> 

		</tr>
 			<tr> 
		    <td> Add new member  </td>
		    <td>
						<% if (  RsMemberexclude.RecordCount > 0) then %>
						
								<select size="10" name="addmember" class="common">
								<%
								do while (  Not RsMemberexclude.EOF)
								%>
										<option value="<%=RsMemberexclude("memberid")%>" ><% response.write "(" & RsMemberexclude("loginname") & ") " & RsMemberexclude("name")%></option>
								
								<%
								RsMemberexclude.movenext
								Loop
								%>
								</select>
						<% else %>
							No member available
						<% end if %>	
							
				</td>
		</tr>



				<% if (  RsMemberexclude.RecordCount > 0) then %>
					<tr><td></td><td>
						<input type="button" value="   Add User  " onClick="AddMember();">&nbsp;
					</td></tr>
				<% end if %>
<% else %>
<input type="button" value="    Submit  " onClick="dosubmit();">
<% End If %>
<input type="hidden" name="whatdo" value="">
	</td>
</tr>
</table>
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