
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



Title = "System Setup"
%>

<%

response.expires = 0

pageno = trim(request("pageno"))

MenuName = trim(request(replace("Menuname1","'","''")))

' Which Main Menu to show
'************************
Userlevel = trim(request(replace("Userlevel","'","''")))

Message = ""

' sorting the menu

if trim(request("action_button")) = "match" then

	vid = split(trim(request("nid")),",")
	
		sql1 = "Delete From MenuUserLevelTable Where UserLevel = "&UserLevel
		
		Conn.Execute sql1
	
	
	for i=0 to ubound(vid)
	
		sql2 = "Insert into MenuUserLevelTable (MenuID, Userlevel) Values ("&  vid(i)  &","& UserLevel &")"
		
	    conn.execute sql2 
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


function doSelect(){
k=0;
document.fm1.action="MenuAccess.asp?sid=<%=SessionID%>&Userlevel=<%=Userlevel%>";
	if (document.fm1.nid!=null)
	{
		for(i=0;i<document.fm1.nid.length;i++)
		{
			if(document.fm1.nid[i].checked)
			  {
			  k=1;
			  i=1;
			  break;
			  }
		}
		if(i==0)
		{
			if (!document.fm1.nid.checked)
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
    document.fm1.action_button.value="match";
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
 <table width="80%" border="0" class="normal">

    <tr> 
      <td  class="BlueClr" <% If FunctionID <> 1 Then %>bgcolor="#C0C0C0"<% End If %> width="30%" >
      <a href="MenuSetup.asp?sid=<% =SessionID %>&functionid=1">Add Menu
      </a>
      </td> 
      <td  class="BlueClr" width="30%" <% If FunctionID <> 2 Then %>bgcolor="#C0C0C0"<% End If %>>
      <a href="MenuSetup.asp?sid=<% =SessionID %>&functionid=2">Edit/Delete Menu
      </a></td> 
      <td  class="BlueClr" <% If FunctionID <> 3 Then %>bgcolor="#C0C0C0"<% End If %>>
      <a href="MenuSetup.asp?sid=<% =SessionID %>&functionid=3">Menu Order</a></td>
       <td  class="BlueClr" <% If FunctionID <> 4 Then %>bgcolor="#C0C0C0"<% End If %>>
      <a href="MenuAccess.asp?sid=<% =SessionID %>&functionid=4&Userlevel=1">Access Right</a></td>  
    </tr>

   </table>

<br>

<form name="fm1" method="post" action="">
<table width="80%" border="0" class="normal">
<tr> 

<%
        
		sql2 = " Select * From UserLevel Order By LevelNumber Desc"
		
		Set Rs2 = Conn.Execute(sql2)
		
		If Not Rs2.EoF Then
		
			Do While Not Rs2.EoF
			
			
		
%>
	
			<td <% If Userlevel <>  Trim(Rs2("LevelNumber"))  Then %>bgcolor="#FFFFCC"<% End If %>>
			<a href="MenuAccess.asp?Userlevel=<% = Rs2("LevelNumber") %>&sid=<% =SessionID %>"><% = Rs2("LevelName") %></a>
			</td>
<%
		Rs2.MoveNext
			Loop
			
End If

%>

	</tr>
	
</table>
<br>
<table width="80%" border="0" class="normal">

	<tr> 
			<td width="21%" bgcolor="#FFFFCC">User Level</td> 
			<td width="77%" bgcolor="#FFFFCC">
			
			　</td>
	</tr>
</table> 	


<table width="80%" border="0" class="normal">
 
 
	<tr> 
			<td width="39%"></td> 
			<td width="60%">
			　</td>
	</tr>
	
<%
		sql = " SELECT * from Menu order by OrderID"
		
		set acres = Conn.Execute(sql)
	
    
	if not acres.eof then
	  	do while not acres.eof
		  
%>
	  <tr> 
      <td width="39%"></td> 
      <td width="60%">
      
<% 
	  Sql2 = "Select * From MenuUserLevelTable where MenuID = "& acres("ID") &" and UserLevel ="&UserLevel
 
	  Set Rs2 = Conn.Execute(Sql2)
	  
	  If Not Rs2.EoF Then
	  	
	  		SelectFlag = 1
	  		
	  End If
	  
%>
      
      <input type="checkbox" name="nid" value="<% = acres("ID") %>" <% If SelectFlag = 1 Then%>Checked<% End If %>>&nbsp;
      <% = acres("MenuName") %>
           　</td>
    </tr>
	
<%
	acres.movenext 
	
		SelectFlag = 0 
	
	loop 

 
	End If
%>
	
	
      <tr> 
      <td width="39%"> <input type="hidden" name="action_button" value=""></td> 
      <td width="60%">
      <input type="button" value="Submit" onClick="doSelect();"></td>
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
<% 
function showpageno(pageno)
	if recount>10 then
		lastpage=recount\10
		if recount mod 10 >0 then
			lastpage=lastpage+1
		end if
		if pageno>10 then
		     response.write "<a href='MenuSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='MenuSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;"
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
				response.write "<a href='MenuSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&i&"'>"&cstr(i)&"</a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\10)<(lastpage\10) then
		        response.write "<a href='MenuSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&(pageno+1-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;"
			   response.write "<a href='MenuSetp.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="& (lastpage) &"'>Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>