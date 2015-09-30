<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



Title = "Help Content Setup"
%>

<%
' Add Menu
'*********
response.expires = 0

pageno = trim(request("pageno"))


' Which Main Menu to show
'************************
FunctionID = Trim(Request("FunctionID"))

If FunctionID = "" Then

FunctionID = 1

End IF

Message = ""


' Add Section
'************
if trim(request("action_button")) = "add section" then


		SectionName = trim(request(replace("SectionName1","'","''")))
		
		HelpContent = trim(request(replace("HelpContent1","'","''")))

        MenuID = trim(request(replace("MenuID","'","''")))
  
		sql1 = "insert into HelpSection (MenuID, SectionName, HelpContent) "

        sql1 = sql1 & "values ("& MenuID &", '"& SectionName & "','"& HelpContent &"')"

		Conn.Execute sql1

        Message =  "The section was added."
	
		set acres=nothing
end if

' Modify the Section
'******************
if trim(request("action_button")) = "modify help content" then

	HelpContent2 = split(trim(request.form("HelpContent2")),",")
	
	mid4 = split(trim(request.form("mid4")),",")


	for i=0 to ubound(mid4)
	
		strsql="Update HelpSection set HelpContent = '"& trim(replace(HelpContent2(i),"'","''")) &"' where SectionID="& trim(mid4(i))
		
	    conn.execute strsql 
	next
	

	
end if


' Delete Section
'***************
if trim(request("action_button")) = "delete section" then

	delid = split(trim(request("id4")),",")

	for j=0 to ubound(delid)
	
		sql4 = "Delete HelpSection where SectionID="& trim(delid(j))
	
	    conn.execute sql4 
	next
	
end if


%>

<!--#include file="include/SQLconn.inc.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<SCRIPT language=JavaScript>
<!--


function addSection()
{
	if(document.fm1.SectionName1.value =="")
       {
		alert("Please enter the Section Name!");
        document.fm1.SectionName1.focus();
        return false;
       }
	else
		{
		document.fm1.action_button.value="add section";
		document.fm1.submit();
		}
}


function EditSection()
{
	
		{
		document.fm1.action_button.value="modify help content";
		document.fm1.submit();
		}
}


function DeleteSection(){
k=0;
document.fm1.action="HelpSetup.asp?sid=<%=SessionID%>&FunctionID=<%=FunctionID%>&UserLevel=<%=UserLevel%>";
	if (document.fm1.id4!=null)
	{
		for(i=0;i<document.fm1.id4.length;i++)
		{
			if(document.fm1.id4[i].checked)
			  {
			  k=1;
			  i=1;
			  break;
			  }
		}
		if(i==0)
		{
			if (!document.fm1.id4.checked)
               k=0;
			else
               k=1;
		}
	}

if (k==0)
  alert("You must select at least one record!");	
else if (k==1)
 {
  var msg = "Are you sure ?";
  if (confirm(msg)==true)
   {
    document.fm1.action_button.value="delete section";
    document.fm1.submit();
   }
 }

}


//-->
</SCRIPT>
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

%>
 <table width="90%" border="0" class="normal">

    <tr> 
      <td  class="BlueClr" <% If FunctionID <> 1 Then %>bgcolor="#C0C0C0"<% End If %> width="25%">
      <a href="HelpSetup.asp?sid=<% =SessionID %>&functionid=1">Add Help Section
      </a>
      </td> 
      <td  class="BlueClr" width="30%" <% If FunctionID <> 2 Then %>bgcolor="#C0C0C0"<% End If %> width="25%">
      <a href="HelpSetup.asp?sid=<% =SessionID %>&functionid=2">Edit/Delete Help Section
      </a></td> 
    </tr>

   </table>

<br>

<form name="fm1" method="post" action="">

<!---- Start of Add Section ---->

  <% If FunctionID <> 1 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>

<% 
           
           Lsql = " Select ID, MenuName,OrderID from Menu Order by OrderID Asc"
           Set LRs = conn.execute(Lsql)
  
%>

<table width="80%" border="0" cellpadding="4" class="normal">



 <tr> 
      <td width="27%">Module Name</td> 
      <td width="69%">
      	  <select name="MenuID" class="common"  size="1">
          <% 
                             If Not LRs.EoF Then
                        LRs.MoveFirst
							do while not LRs.eof
                                 response.write "<option value="&LRs("ID")&">"&trim(LRs("MenuName"))&"</option>"
                               LRs.movenext
							loop
						
						End if
					%>
        </select>
</td>
    </tr>
    
 <tr> 
      <td width="27%">
Section Name</td> 
      <td width="69%">
      <Input name="SectionName1" type=text value="" size="50"></td>
    </tr>
 <tr> 
      <td width="27%">
Help Content</td> 
      <td width="69%">
      <textarea name="HelpContent1" cols="100" rows="10"></textarea></td>
    </tr>
 <tr> 
      <td width="27%">
¡@</td> 
      <td width="69%">
      	<input type="button" value="    add    " onClick="javascript:addSection();">
         <input type="hidden" name="action_button" value="">   
¡@</td>
    </tr>

<tr> 
      <td colspan="2" align =center><font color="red"><% = Message %></font></td> 
      <td >
¡@</td>
    </tr>

    </table>
</div>

<!---- End of Add Section ------>
<!---- Start of Modify/Delete Section ---->

  <% If FunctionID <> 2 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>
  
<table width="90%" border="0" class="normal">

	<tr> 
			<td width="20%" bgcolor="#FFFFCC">Module</td> 
			<td width="20%" bgcolor="#FFFFCC">Section</td>
            <td width="55%" bgcolor="#FFFFCC">Content</td>
            <td width="5%" bgcolor="#FFFFCC">Delete</td>
	</tr>



<%
        
		sql4 = " Select * From Menu m Join HelpSection s "

        sql4 = sql4 & " on m.ID = s.MenuID Order By m.OrderID, s.SectionName Desc"

	    set Rs4 = nothing
				set Rs4 = createobject("adodb.recordset")
				Rs4.cursortype = 3
				Rs4.locktype = 1
				Rs4.open sql4,conn
				
				recount = Rs4.recordcount
			
				if pageno="" then
				   pageno=1
			    end if 
			    
				 if pageno > 1 then
				   i = (pageno - 1) * 10
				   Rs4.move i
				 end if
				 
				   
	
%>
<tr bgcolor="#ffffff"> 
      <td colspan="4">
        
          <%call showpageno(pageno)%>
          
      </td>
    </tr>
<%		
		If Not Rs4.EoF Then

             		
			Do While Not Rs4.EoF

       if j> 9 then exit do
		
%>
<tr> 
	<td width="20%">
<% = Rs4("MenuName") %>
</td> 
			<td width="20%"><% = Rs4("SectionName") %>
			¡@</td>
<td width="55%">
<textarea name="HelpContent2" cols="80" rows="2"><% = Rs4("HelpContent") %></textarea>
			¡@</td>
<td><input type="checkbox" name="id4" value="<% = Rs4("SectionID") %>"></td> 
	</tr>
<input type="hidden" name="mid4" value="<% = Rs4("SectionID") %>">	
<%
	Rs4.movenext 

      j=j+1 

	   loop 
 
	End If
%>
<tr bgcolor="#ffffff"> 
      <td colspan="4">
        <div align="right"> 
          <%call showpageno(pageno)%>
          </div>
      </td>
    </tr>
	
      <tr> 
      <td colspan="2"></td> 
      <td width="30%">
      <input type="button" value="Edit" onClick="EditSection();"></td>
      <td width="30%">
      <input type="button" value="Delete" onClick="DeleteSection();"></td>
    </tr>

</table> 
</div>
</div>
               
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
<% 
function showpageno(pageno)
	if recount>10 then
		lastpage=recount\10
		if recount mod 10 >0 then
			lastpage=lastpage+1
		end if
		if pageno>10 then
		     response.write "<a href='HelpSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='HelpSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;"
		end if
		strtemp=pageno
		if (pageno Mod 10 )=0 then
		   strtemp=strtemp-10
		end if
	 for i=(strtemp-(strtemp mod 10)+1) to (strtemp+10-(strtemp mod 10))
	         if lastpage<i then  exit for			 
            if i- pageno=0 then
				response.write cstr(i)&"&nbsp;&nbsp;"
			else
				response.write "<a href='HelpSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&i&"'>"&cstr(i)&"</a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\10)<(lastpage\10) then
		        response.write "<a href='HelpSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&(pageno+1-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;"
			   response.write "<a href='HelpSetp.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="& (lastpage) &"'>Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>