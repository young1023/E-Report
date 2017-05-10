
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



Title = "System Setup"
%>

<%
' Add Menu
'*********
response.expires = 0

pageno = trim(request("pageno"))

MenuName = trim(request(replace("Menuname1","'","''")))

' Which Main Menu to show
'************************
FunctionID = Trim(Request("FunctionID"))

If FunctionID = "" Then

FunctionID = 1

End IF

Message = ""

if trim(request("action_button")) = "add" then


		MenuName = trim(request(replace("Menuname1","'","''")))
		
		PageLink = trim(request(replace("Link1","'","''")))

	    strsql="Select * from Menu where MenuName='"& MenuName & "'"
	    
		set acres=conn.execute(strsql)
	    if acres.eof then
	    
			strsql="insert into Menu (MenuName, PageLink, OrderID) values ('"& MenuName & "', '"& PageLink & "', 1)"
		
			conn.execute strsql
		else
		   
		   Message =  "This Menu Name Exit,Please Input Other Menu Name"
		   
		end if
		set acres=nothing
end if

' modify the menu

if trim(request("action_button")) = "Modify" then

	vMenuName = split(trim(request("MenuName2")),",")
	
	vPageLink = split(trim(request("Link2")),",")
	
	vid = split(trim(request("xid")),",")
	
	for i=0 to ubound(vid)
	
		strsql="Update Menu set MenuName = '"& trim(replace(vMenuName(i),"'","''")) &"' where id="& trim(vid(i))
		response.write 	strsql
	    conn.execute strsql 
	next
	
	for i=0 to ubound(vid)
	
		strsql="Update Menu set PageLink = '"& trim(replace(vPageLink(i),"'","''")) &"' where id="& trim(vid(i))
		
	    conn.execute strsql 
	next
	
	
end if


if trim(request("action_button")) = "Delete" then

	vid = split(trim(request("xid")),",")

	for i=0 to ubound(vid)
	
		strsql="Delete Menu where id="& trim(vid(i))
		
	    conn.execute strsql 
	next
	
end if


' sorting the menu

if trim(request("action_button")) = "sort" then

	vOrder = split(trim(request("OrderID")),",")
	
	vid = split(trim(request("mid")),",")
	
	for i=0 to ubound(vid)
	
		strsql="Update Menu set OrderID = "& trim(replace(vOrder(i),"'","''")) &" where id="& trim(vid(i))
		
	    conn.execute strsql 
	next
	
end if


' sorting the menu

if trim(request("action_button")) = "MenuUp" then

	MenuID = trim(request("MenuOrder"))
	
        sql1 = "Select OrderID - 1 as OrderID From Menu Where ID="&MenuID

        Set Rs1 = Conn.Execute(sql1)

        'response.write rs1("orderid")

        sql2 = "Select ID as id From menu where orderid="&rs1("orderid")

        Set Rs2 = Conn.Execute(sql2)

        'response.write rs2("id")

	
	    sql3="Update Menu set OrderID = OrderID - 1 where ID="&MenuID
		
	    conn.execute(sql3) 

        sql4="Update Menu set OrderID = OrderID + 1 where ID="&Rs2("ID")

        conn.execute(sql4)

   
	
end if

' sorting the menu

if trim(request("action_button")) = "MenuDown" then

	MenuID = trim(request("MenuOrder"))
	
        sql1 = "Select OrderID + 1 as OrderID From Menu Where ID="&MenuID

        Set Rs1 = Conn.Execute(sql1)

        'response.write rs1("orderid")

        sql2 = "Select ID as id From menu where orderid="&rs1("orderid")

        Set Rs2 = Conn.Execute(sql2)

        'response.write rs2("id")

	
	    sql3="Update Menu set OrderID = OrderID + 1 where ID="&MenuID
		
	    conn.execute(sql3) 

        sql4="Update Menu set OrderID = OrderID - 1 where ID="&Rs2("ID")

        conn.execute(sql4)

     

	
end if

' sorting the menu

if trim(request("action_button")) = "sort" then

	vOrder = split(trim(request("OrderID")),",")
	
	vid = split(trim(request("mid")),",")
	
	for i=0 to ubound(vid)
	
		strsql="Update Menu set OrderID = "& trim(replace(vOrder(i),"'","''")) &" where id="& trim(vid(i))
		
	    conn.execute strsql 
	next
	
end if


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

function doadd()
{
	if((document.fm1.MenuName1.value)=="")
		alert("Please enter Menu Name!");
	else
		{
		document.fm1.action_button.value="add";
		document.fm1.submit();
		}
}


function doModify()
{
	
		{
		document.fm1.action_button.value="Modify";
		document.fm1.submit();
		}
}

function doOrder()
{
	var i,strmsg,strval,j;
	strmsg="";
	j=0;
	if(document.fm1.OrderID != null)
	{
		for(i=0;i<document.fm1.OrderID.length;i++)
			{
				strval=document.fm1.OrderID[i].value;
				if(strval=="")
				{
					strmsg="The sorting field cannot be null.";
					j=1;
					break;
				}
			}
		if (j==0 && i==0)
	    {
			if(document.fm1.OrderID.value=="")
				strmsg="The sorting field cannot be null.";
		}
	}
	if (strmsg!="")
		alert(strmsg);
	else
	{
		document.fm1.action_button.value="sort";
		document.fm1.submit();
	}
}


function doSelect(){
k=0;
document.fm1.action="MenuSetup.asp?sid=<%=SessionID%>&FunctionID=<%=FunctionID%>&UserLevel=<%=UserLevel%>";
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

function delcheck(){
k=0;
document.fm1.action="MenuSetup.asp?sid=<%=SessionID%>&FunctionID=<%=FunctionID%>";
	if (document.fm1.xid!=null)
	{
		for(i=0;i<document.fm1.xid.length;i++)
		{
			if(document.fm1.xid[i].checked)
			  {
			  k=1;
			  i=1;
			  break;
			  }
		}
		if(i==0)
		{
			if (!document.fm1.xid.checked)
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
    document.fm1.action_button.value="Delete";
    document.fm1.submit();
   }
 }

}


function doUp()
{
        if(document.getElementById('MenuOrder').selectedIndex == -1){
        alert("Please select a option in the menu!");
        return false;
        }
        
        if(document.getElementById('MenuOrder').selectedIndex == 0){
        alert("It is already at the top position in the menu!");
        return false;
        }

        {
		document.fm1.action_button.value="MenuUp";
		document.fm1.submit();
		}
	
}

function doDown()
{
        if(document.getElementById('MenuOrder').selectedIndex == -1){
        alert("Please select a option in the menu!");
        return false;
        }

       var x = document.getElementById('MenuOrder').length;
       x = x-1;
       if(document.getElementById('MenuOrder').selectedIndex == x){
        alert("It is already at the bottom position in the menu!");
        return false;
        }

        {
		document.fm1.action_button.value="MenuDown";
		document.fm1.submit();
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
      <a href="MenuSetup.asp?sid=<% =SessionID %>&functionid=4&UserLevel=8">Access Right</a></td>  
    </tr>

   </table>

<br>

<form name="fm1" method="post" action="">
<% If FunctionID <> 1 Then %>		
  		
  <div style="display:none" align=center>
  		
<% End If %>
  

<table width="80%" border="0" class="normal">
  <tr> 
      <td colspan="2"  align="right">
<font color="red">*</font>	 Denotes a mandatory field</td>
    </tr>
  <tr> 
      <td colspan="2"  align="right">
¡@</td>
    </tr>
    
  
    
 <tr> 
      <td width="27%">Menu Name</td> 
      <td width="69%">
      	     
				<Input name="MenuName1" type=text value="" size="50"></td>
    </tr>
    
 <tr> 
      <td width="27%">
Page Link</td> 
      <td width="69%">
      <Input name="Link1" type=text value="" size="50"></td>
    </tr>
    
 <tr> 
      <td width="27%">
¡@</td> 
      <td width="69%">
      	<input type="button" value="    Add   " onClick="javascript:doadd();">
        <input type="hidden" name="action_button" value="">     
¡@</td>
    </tr>
    
 <tr> 
      <td colspan="2" align =center><font color="red"><% = Message %></font></td> 
      <td >
¡@</td>
    </tr>
    
<%
	sql = " SELECT * from Menu order by ID Asc"
				set acres = nothing
				set acres = createobject("adodb.recordset")
				acres.cursortype = 1
				acres.locktype = 1
				acres.open sql,conn
				
				recount = acres.recordcount
			
				if pageno="" then
				   pageno=1
			    end if 
			    
				 if pageno>1 then
				   i = (pageno-1) * 10
				   acres.move i
				 end if
				 
				   
	
%>
<tr> 
      <td width="27%"><% call showpageno(pageno) %></td> 
      <td width="69%">
      ¡@</td>
    </tr>

     <tr> 
      <td width="27%" bgcolor="#FFFFCC">ID</td> 
      <td width="69%" bgcolor="#FFFFCC">Menu Name
      ¡@</td>
    </tr>
<% 
	if not acres.eof then
	  	do while not acres.eof
		   if j> 9 then exit do
%>

     <tr> 
      <td width="27%"><% = acres("ID") %>¡@</td> 
      <td width="69%"><% = acres("MenuName") %></td>
    </tr>
<%
	acres.movenext  
	j=j+1 
	loop 
%>
<% end if  %>
    <tr bgcolor="#ffffff"> 
      <td colspan="2">
        <div align="right"> 
          <%call showpageno(pageno)%>
          </div>
      </td>
    </tr>
</table>
</div>

<!---- End of Add Menu ------>

<!---- Start of Edit/Detete Menu ---->

  <% If FunctionID <> 2 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>


<table width="80%" border="0" class="normal">

     <tr> 
      <td width="85%" colspan="2">¡@</td> 
      <td width="13%">

      ¡@</td>
    </tr>
    
    <%
	sql = " SELECT * from Menu order by OrderID Asc"
				
	Set acres = Conn.Execute(sql)
				 
	
%>
     <tr> 
      
      <td width="43%" bgcolor="#FFFFCC">Menu Name</td>
      
      <td width="42%" bgcolor="#FFFFCC">Page Link</td>
       <td width="13%" bgcolor="#FFFFCC">Select</td>
    </tr>
<% 
	if not acres.eof then
	  	do while not acres.eof
		 
%>

     <tr> 
      <td width="43%">     	     
				<Input name="MenuName2" type=text value="<% = acres("MenuName") %>" size="50"></td>
      <td width="42%">     	     
				<Input name="Link2" type=text value="<% = acres("PageLink") %>" size="50"></td>
				      <td width="13%"><input type="checkbox" name="xid" value="<% = acres("ID") %>"></td> 

    </tr>
<%
	acres.movenext  
	
	loop 
%>
<% end if  %>

     <tr> 
      <td width="85%" colspan="2">
      <input type="button" value="Modify" onClick="doModify();">
        ¡@</td>
              <td width="13%">
        <input type="button" value="Delete" onClick="delcheck();"></td> 

    </tr>
</table>
</div>

<!---- End of Edit/Delete Menu ------>
<!---- Start of Sorting Menu ---->

  <% If FunctionID <> 3 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>


<table width="80%" border="0" class="normal">



	    
    <%
	sql = " SELECT * from Menu order by OrderID Asc"
	
	Set acres = Conn.Execute(sql)

				
%>
     <tr> 
      <td bgcolor="#FFFFCC" align="center" colspan="2">Menu Name
      ¡@</td>
      
    </tr>


     <tr> 
       
      <td align="center" width="60%">     	 

	<select size="30" name="MenuOrder" id="MenuOrder" class="common">

    <% 
	if not acres.eof then
	  	do while not acres.eof
		   
    %>

<option value="<%=acres("ID")%>" <%If Trim(MenuID)=Trim(acres("ID")) Then%>selected<%End If%> >&nbsp;&nbsp;&nbsp;&nbsp;<% = acres("MenuName") %>&nbsp;&nbsp;&nbsp;&nbsp;</option>

<%

                   acres.movenext

							loop
						
						End if
					%>


    </select>
				</td>

    <td>
<input type="button" value="&#916;" onClick="doUp();">

<br/><br/><br/>

<input type="button" value="&#8711;" onClick="doDown();">

     </td>

    </tr>


    </table>
</div>

<!---- End of Sorting Menu ------>
<!---- Start of Acess Right Menu ---->

  <% If FunctionID <> 4 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>
  

<table width="80%" border="0" class="normal">
<tr> 

<%
        
		sql2 = " Select * From UserLevel Order By LevelNumber Desc"
		
		Set Rs2 = Conn.Execute(sql2)
		
		If Not Rs2.EoF Then
		
			Do While Not Rs2.EoF
			
			
		
%>
	
			<td <% If Userlevel <>  Trim(Rs2("LevelNumber"))  Then %>bgcolor="#FFFFCC"<% End If %>>
			<a href="MenuSetup.asp?FunctionID=4&Userlevel=<% = Rs2("LevelNumber") %>&sid=<% =SessionID %>"><% = Rs2("LevelName") %></a>
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
			
			¡@</td>
	</tr>
</table> 	


<table width="80%" border="0" class="normal">
 
 
	<tr> 
			<td width="39%"></td> 
			<td width="60%">
			¡@</td>
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
	  Sql2 = "Select * From MenuUserLevelTable where MenuID = " & acres("ID") & " and UserLevel = " & Userlevel 
	  'Sql2 = "Select * From MenuUserLevelTable where  UserLevel = " & UserLevel & " and MenuID = " & acres("ID") 
	  
	  'response.write sql2
          'response.end
	  
	  Set Rs2 = Conn.Execute(Sql2)
	  
	  If Not Rs2.EoF Then
	  	
	  		SelectFlag = 1
	  		
	  End If
	  
%>
      
      <input type="checkbox" name="nid" value="<% = acres("ID") %>" <% If SelectFlag = 1 Then%>Checked<% End If %>>&nbsp;
      <% = acres("MenuName") %>
           ¡@</td>
    </tr>
	
<%
	acres.movenext 
	
		SelectFlag = 0 
	
	loop 

 
	End If
%>
	
	
      <tr> 
      <td width="39%"></td> 
      <td width="60%">
      <input type="button" value="Submit" onClick="doSelect();"></td>
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