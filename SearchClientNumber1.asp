<% 
'on error resume next
'*********************************************************************************
'NAME       : SearchClientNumber.asp          
'DESCRIPTION: Search and filter client number and used for all reports
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 090621 Roger Wong   Prototype
'MODIFIED   : 
'********************************************************************************
pageid=trim(request.form("pageid"))


if pageid="" then
  pageid=1
end if
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
 	  self.opener.document.fm1.ClientTo.value=myString;
 	  self.opener.document.fm1.ClientFrom.value=myString;
 		
 		self.close();
 	
}


function gtpage(what)
{
document.myform.pageid.value=what;
document.myform.action="SearchClientNumber.asp?sid=<%=SessionID%>"
document.myform.submit();
}


//-->
</script>


</HEAD>

<%
'define query
Const MaxNumberRealTimeFilter = 1000

Dim Search_AECode 
Dim Search_AEGroup
Dim iRecordCount
	
Search_keyword	= Request.form("keyword")
			
strURL = Request.ServerVariables("URL") 
%>	

<BODY OnLoad="document.myform.keyword.focus();document.myform.keyword.select();">
	<FORM NAME="myform"  method="post"  action="<%= strURL %>?sid=<%=SessionID%>">
		<table width="99%" border="0" class="normal">
			<tr><td class="common"> 
					Enter client No., English Name or Chinese Name
							</td></tr>
			<tr><td class="common"> 
					<INPUT name="keyword" value="<%= Search_keyword %>">
						<INPUT TYPE=submit value="Filter">
							<INPUT TYPE=button value="Clear">
</td></tr>
								
<%							
			
	If Request.Form("keyword") <> ""  Then

			    If Session("shell_power") > 5 Then

                sql = "Select * From Client Where 1 = 1"

                Elseif Session("shell_power") > 1 Then
	       
                Search_AEGroup = Session("GroupID")

                ' Branch Manager 

                sql = "Select * from Client c , Member m Where c.AECode = m.LoginName "
 
                sql = sql & " and m.GroupID = "&Search_AEGroup

                Else


		        Search_AECode = Session("id")
		
               
                ' Check if the AE belong to any Shared Group

                Set Rs = server.createobject("adodb.recordset")  
                Rs.open ("exec Retrieve_SharedGroup '"&Search_AECode&"'") ,  StrCnn,3,1

                    If Not Rs.EoF Then
             
                Search_SharedGroup = Rs("SharedGroupID") 

                'sql = "Select * from Client c, (Member m Left Join SharedGroup s on "

		        'sql = sql & "m.MemberID = s.MemberID) Where  c.AECode = m.LoginName"


                'sql = sql & "and s.SharedGroupID in (Select SharedGroupID From SharedGroup s , Member n Where s.memberID = n.MemberID and n.LoginName = '"&Search_SharedGroup&"')"

                    Else
                
                ' AE only, not belong to any shared group
                

                sql = "Select * from Client Where AECode = '889' "


                    End If

                End if

                End If

                
                If Search_keyword <> "" Then
                
                sql = sql & "and  (ClntCode like '%"&Search_keyword&"%' or c.name like  '%"&Search_keyword&"%' or cname like '%"&Search_keyword&"%' or Ename like '%"&Search_keyword&"%' )"

                End If

                
                Response.write sql

                
            
       		     Set Rs1 = server.createobject("adodb.recordset")
		         Rs1.cursortype=1
		         Rs1.locktype=1
                 Rs1.open sql,conn

			if Rs1.RecordCount=0 then

					'no record found
					response.write ("<tr><td>No record found</td></tr>")
				
				else

			  	
	
         Rs1.PageSize = 10
         
             end if
	 
	
	%>		 
								<tr><td class="common">
								Select client (Total <%=Rs1.RecordCount%> clients), 
 <%
	 if Rs1.recordcount>0 then
             
			 response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			 if Clng(pageid) <>1 then
                 response.write " <a href=javascript:gtpage('1') style='cursor:hand' >First</a> "
                 response.write " <a href=javascript:gtpage('"&(pageid-1)&"') style='cursor:hand' >Previous</a> "
			 end if
             call countpage(Rs1.PageCount,pageid)
	         if Clng(pageid)<>Clng(Rs1.PageCount) then
                 response.write " <a href=javascript:gtpage('"&(pageid+1)&"') style='cursor:hand' >Next</a> "
                 response.write " <a href=javascript:gtpage('"&Rs1.PageCount&"') style='cursor:hand' >Last</a> "
			 end if
	         response.write "&nbsp;&nbsp;"
	 end if
%>
								</td></tr>
							
								
		<tr>
           <td class="common">
	<SELECT NAME="myselect" SIZE=10 class="common">
									
<%
									
 i=0
 if Rs1.recordcount>0 then
  Rs1.AbsolutePage = pageid
  do while (Rs1.PageSize-i)
   if Rs1.eof then exit do
   i=i+1

  
%>

<Option value="<%=rs1("clntcode")%>"> 
<% response.write Rs1("clntcode") + " : " +  rs1("ename") + " " +  rs1("cname")%> </OPTION>
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
			
			     response.write "<input type=hidden value="&pageid&" name=pageid>"

					
						
			'argument end if						
			end if 
			%>
			
			
								
								
								</table>
								</FORM>
							

	
<%

  ' function
  Sub countpage(PageCount,pageid)
	   if PageCount>=1 and PageCount<=10 then
		 for i=1 to PageCount
		   if (pageid-i =0) then
              response.write "<font color=green> "&i&"</font> "
		   else
             response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		   end if
		 next
	   elseif PageCount>11 then
	      if pageid<=5 then
		     for i=1 to 10
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       else
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
		     next
		  else
		    for i=(pageid-4) to (pageid+5)
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       elseif i=<pagecount then
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
			next
		  end if
	   end if
  end sub
%>
</BODY>
</HTML>