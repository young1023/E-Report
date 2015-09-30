
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



Title = "Printing Setup"
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

MenuID = trim(request(replace("MenuID","'","''")))

If MenuID = "" Then

MenuID = 1

End IF

Message = ""

' Setup Print for Member
'************************
if trim(request("action_button")) = "allow printing" then

	xid = split(trim(request("xid")),",")
	
		sql1 = "Delete From AllowPrint Where "

        sql1 = sql1 & "PrintAllowed = 1 and MenuID ="& MenuID 

		Conn.Execute sql1
	
	
	for i=0 to ubound(xid)
	
		sql2 = "Insert into AllowPrint (MemberID, MenuID, PrintAllowed) "

        sql2 = sql2 & "Values ("&  xid(i)  &","& MenuID &", 1)"
		
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

function doPrint(){
k=0;
document.fm1.action="PrintSetup.asp?sid=<%=SessionID%>&FunctionID=<%=FunctionID%>&MenuID=<%=MenuID%>";
   
    document.fm1.action_button.value="allow printing";
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

<form name="fm1" method="post" action="">

<!---- Start of Printing Access Right for Member Menu ---->

  <% If FunctionID <> 1 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>


<table width="80%" border="0" class="normal" border="1">
<tr bgcolor="#FFFF00"> 

<%
        
		sql1 = " Select * From Menu where OrderID < 7 Order By OrderID Desc"
		
		Set Rs2 = Conn.Execute(sql1)
		
		If Not Rs2.EoF Then
		
			Do While Not Rs2.EoF
			
			
		
%>
	
			<td <% If Trim(MenuID) <>  Trim(Rs2("ID"))  Then %>bgcolor="#FFFFFF"<% End If %>>
			<a href="PrintSetup.asp?FunctionID=1&MenuID=<% = Rs2("ID") %>&sid=<% =SessionID %>"><% = Rs2("MenuName") %></a>
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
			<td width="39%"></td> 
			<td width="60%">
			¡@</td>
	</tr>
	
<%
		sql6 = " SELECT * from Member Order by Name"
		
        set acres = Conn.Execute(sql6)
	
    
	if not acres.eof then
	  	do while not acres.eof
		  
%>
	  <tr> 
      <td width="39%"></td> 
      <td width="60%">
      
<% 
	  Sql7 = "Select * From AllowPrint where PrintAllowed = 1 and MemberID = " & acres("MemberID") & " and MenuID = "&MenuID 
	  
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
      <input type="button" value="Submit" onClick="doPrint();"></td>
      <input type="hidden" name="action_button" value="">  
    </tr>
</table>
</div>

<!---- End of Printing Access Right for Member Menu ------>


	
	
                 
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
		     response.write "<a href='PrintSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='PrintSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;"
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
				response.write "<a href='PrintSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&i&"'>"&cstr(i)&"</a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\10)<(lastpage\10) then
		        response.write "<a href='PrintSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&(pageno+1-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;"
			   response.write "<a href='MenuSetp.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="& (lastpage) &"'>Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>