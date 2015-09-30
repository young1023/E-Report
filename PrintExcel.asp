
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



Title = "Printing and Excel Exporting Access Right Setup"
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

Userlevel = trim(request(replace("Userlevel","'","''")))

If Userlevel = "" Then

UserLevel = 1

End IF

Message = ""

' Setup Print for Member
'************************
if trim(request("action_button")) = "member print" then

	xid = split(trim(request("xid")),",")
	
		sql1 = "Delete From PrintExcelTable Where GroupID is null "

        sql1 = sql1 & "and PrintAllowed = 1 and UserLevel ="& UserLevel 

		Conn.Execute sql1
	
	
	for i=0 to ubound(xid)
	
		sql2 = "Insert into PrintExcelTable (MemberID, Userlevel, PrintAllowed) "

        sql2 = sql2 & "Values ("&  xid(i)  &","& UserLevel &", 1)"
		
	    conn.execute sql2 
	next

Message = "System was changed"
	
end if

' Setup Group to Print
'************************
if trim(request("action_button")) = "group print" then

	gid = split(trim(request("gid")),",")
	
		sql1 = "Delete From PrintExcelTable Where MemberID is null "

        sql1 = sql1 & "and PrintAllowed = 1"

		Conn.Execute sql1
	
	
	for i=0 to ubound(gid)
	
		sql2 = "Insert into PrintExcelTable (GroupID,  PrintAllowed) "

        sql2 = sql2 & "Values ("&  gid(i) & " , 1)"
		
	    conn.execute sql2 
	next

Message = "System was changed"
	
end if


' Setup Excel
'************************
if trim(request("action_button")) = "allow Excel" then

	vid = split(trim(request("nid")),",")
	
		sql1 = "Delete From PrintExcelTable Where GroupID is null "

        sql1 = sql1 & "and ExcelAllowed = 1 and UserLevel ="& UserLevel 

		Conn.Execute sql1
	
	
	for i=0 to ubound(vid)
	
		sql2 = "Insert into PrintExcelTable (MemberID, Userlevel, ExcelAllowed) "

        sql2 = sql2 & "Values ("&  vid(i)  &","& UserLevel &", 1)"
		
	    conn.execute sql2 
	next

Message = "System was changed"
	
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

function doExcel(){
k=0;
document.fm1.action="PrintExcel.asp?sid=<%=SessionID%>&FunctionID=<%=FunctionID%>&UserLevel=<%=UserLevel%>";
   
    document.fm1.action_button.value="allow Excel";
    document.fm1.submit();
}

function doPrintMember(){
k=0;
document.fm1.action="PrintExcel.asp?sid=<%=SessionID%>&FunctionID=<%=FunctionID%>&UserLevel=<%=UserLevel%>";
   
    document.fm1.action_button.value="member print";
    document.fm1.submit();
}

function doPrintGroup(){
k=0;
document.fm1.action="PrintExcel.asp?sid=<%=SessionID%>&FunctionID=<%=FunctionID%>&UserLevel=<%=UserLevel%>";
   
    document.fm1.action_button.value="group print";
    document.fm1.submit();
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
      <a href="PrintExcel.asp?sid=<% =SessionID %>&functionid=1&UserLevel=1">Printing Access Right for Member
      </a>
      </td> 
      <td  class="BlueClr" width="30%" <% If FunctionID <> 2 Then %>bgcolor="#C0C0C0"<% End If %>>
      <a href="PrintExcel.asp?sid=<% =SessionID %>&functionid=2&UserLevel=1">Printing Access Right for Group
      </a></td> 
      <td  class="BlueClr" <% If FunctionID <> 3 Then %>bgcolor="#C0C0C0"<% End If %>>
      <a href="PrintExcel.asp?sid=<% =SessionID %>&functionid=3&UserLevel=1">Access Right for Exporting Excel</a></td>
    </tr>

   </table>

<br>

<form name="fm1" method="post" action="">

<!---- Start of Printing Access Right for Member Menu ---->

  <% If FunctionID <> 1 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>


<table width="80%" border="0" class="normal" border="1">
<tr> 

<%
        
		sql1 = " Select * From UserLevel Order By LevelNumber Desc"
		
		Set Rs2 = Conn.Execute(sql1)
		
		If Not Rs2.EoF Then
		
			Do While Not Rs2.EoF
			
			
		
%>
	
			<td <% If Userlevel <>  Trim(Rs2("LevelNumber"))  Then %>bgcolor="#FFFFCC"<% End If %>>
			<a href="PrintExcel.asp?FunctionID=1&Userlevel=<% = Rs2("LevelNumber") %>&sid=<% =SessionID %>"><% = Rs2("LevelName") %></a>
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
		sql6 = " SELECT * from Member Where UserLevel = " & UserLevel & " Order by LoginName"
		
        set acres = Conn.Execute(sql6)
	
    
	if not acres.eof then
	  	do while not acres.eof
		  
%>
	  <tr> 
      <td width="39%"></td> 
      <td width="60%">
      
<% 
	  Sql7 = "Select * From PrintExcelTable where PrintAllowed = 1 and MemberID = " & acres("MemberID") & " and Userlevel = " & Userlevel  
	  
	  'response.write sql7
  	  
	  Set Rs2 = Conn.Execute(Sql7)
	  
	  If Not Rs2.EoF Then
	  	
	  		SelectFlag = 1
	  		
	  End If
	  
%>
      
      <input type="checkbox" name="xid" value="<% = acres("MemberID") %>" <% If SelectFlag = 1 Then%>Checked<% End If %>>&nbsp;
      <% = acres("Name") %> ( <% = acres("LoginName") %> )
           ¡@</td>
    </tr>
	
<%
	acres.movenext 
	
		SelectFlag = 0 
	
	loop 

 
	End If
%>
 <tr> 
      <td colspan="2" align =center><font color="red"><% = Message %></font></td> 
      <td >
¡@</td>
    </tr>

  <tr> 
      <td width="39%"></td> 
      <td width="60%">
      <input type="button" value="Submit" onClick="doPrintMember();"></td>
    </tr>
</table>
</div>

<!---- End of Printing Access Right for Member Menu ------>
<!---- Start of Printing Access Right for Group Menu ---->

  <% If FunctionID <> 2 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>


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
		sql_group = " SELECT * from UserGroup"
		
        set acres = Conn.Execute(sql_group)
	
    
	if not acres.eof then
	  	do while not acres.eof
		  
%>
	  <tr> 
      <td width="39%"></td> 
      <td width="60%">
      
<% 
	  Sql7 = "Select * From PrintExcelTable where PrintAllowed = 1 and GroupID = " & acres("GroupID")   
	  
	  'response.write sql7
  	  
	  Set Rs2 = Conn.Execute(Sql7)
	  
	  If Not Rs2.EoF Then
	  	
	  		SelectFlag = 1
	  		
	  End If
	  
%>
      
      <input type="checkbox" name="gid" value="<% = acres("GroupID") %>" <% If SelectFlag = 1 Then%>Checked<% End If %>>&nbsp;
      <% = acres("Name") %>
           ¡@</td>
    </tr>
	
<%
	acres.movenext 
	
		SelectFlag = 0 
	
	loop 

 
	End If
%>
 <tr> 
      <td colspan="2" align =center><font color="red"><% = Message %></font></td> 
      <td >
¡@</td>
    </tr>

  <tr> 
      <td width="39%"></td> 
      <td width="60%">
      <input type="button" value="Submit" onClick="doPrintGroup();"></td>
    </tr>
</table>
</div>

<!---- End of Start of Printing Access Right for Group Menu ------>
<!---- Start of Acess Right for Exporting Excel Menu ---->

  <% If FunctionID <> 3 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>
  

<table width="80%" border="0" class="normal" border="1">
<tr> 

<%
        
		sql5 = " Select * From UserLevel Order By LevelNumber Desc"
		
		Set Rs2 = Conn.Execute(sql5)
		
		If Not Rs2.EoF Then
		
			Do While Not Rs2.EoF
			
			
		
%>
	
			<td <% If Userlevel <>  Trim(Rs2("LevelNumber"))  Then %>bgcolor="#FFFFCC"<% End If %>>
			<a href="PrintExcel.asp?FunctionID=3&Userlevel=<% = Rs2("LevelNumber") %>&sid=<% =SessionID %>"><% = Rs2("LevelName") %></a>
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
		sql6 = " SELECT * from Member Where UserLevel = " & UserLevel & " order by LoginName"
		
        set acres = Conn.Execute(sql6)
	
    
	if not acres.eof then
	  	do while not acres.eof
		  
%>
	  <tr> 
      <td width="39%"></td> 
      <td width="60%">
      
<% 
	  Sql7 = "Select * From PrintExcelTable where ExcelAllowed = 1 and MemberID = " & acres("MemberID") & " and Userlevel = " & Userlevel  
	  
	  'response.write sql7
  	  
	  Set Rs2 = Conn.Execute(Sql7)
	  
	  If Not Rs2.EoF Then
	  	
	  		SelectFlag = 1
	  		
	  End If
	  
%>
      
      <input type="checkbox" name="nid" value="<% = acres("MemberID") %>" <% If SelectFlag = 1 Then%>Checked<% End If %>>&nbsp;
      <% = acres("Name") %> ( <% = acres("LoginName") %> )
           ¡@</td>
    </tr>
	
<%
	acres.movenext 
	
		SelectFlag = 0 
	
	loop 

 
	End If
%>
 <tr> 
      <td colspan="2" align =center><font color="red"><% = Message %></font></td> 
      <td >
¡@</td>
    </tr>	
	
      <tr> 
      <td width="39%"></td> 
      <td width="60%">
      <input type="button" value="Submit" onClick="doExcel();"></td>
    </tr>
<input type="hidden" name="action_button" value="">     
</table> 
</div>
<!---- End of Start of Printing Access Right for Group Menu ------>

	
	
                 
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
		     response.write "<a href='PrintExcel.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='PrintExcel.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;"
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
				response.write "<a href='PrintExcel.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&i&"'>"&cstr(i)&"</a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\10)<(lastpage\10) then
		        response.write "<a href='PrintExcel.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&(pageno+1-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;"
			   response.write "<a href='MenuSetp.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="& (lastpage) &"'>Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>