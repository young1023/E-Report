<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



Title = "Stock Reconciliation Setup"
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


' Add Depot
'************
if trim(request("action_button")) = "add section" then


		DepotName1 = trim(request(replace("DepotName1","'","''")))
		
		DepotFolder1 = trim(request(replace("DepotFolder1","'","''")))

        Market = trim(request(replace("Market","'","''")))

        FileType1 = trim(request(replace("FileType1","'","''")))

        FirstRow1 = trim(request(replace("FirstRow1","'","''")))


        Delimiter1 = trim(request(replace("Delimiter1",",","\,")))

		sql1 = "insert into ReconDepotFolder (Market, DepotName, DepotFolder, FileType, FirstRow,delimiter , ReadyToConvert) "

        sql1 = sql1 & "values ('"& Market &"', '"& DepotName1 & "', '"& DepotFolder1 &"', '"& FileType1 &"' , "& FirstRow1 &" ,"& Delimiter1 &" , 1)"

		Conn.Execute sql1

        Message =  "The depot ("&DepotName1&") was added."
	
		set acres=nothing
end if

' Add Field
'************
if trim(request("action_button")) = "add field" then


		FieldName1 = trim(request(replace("FieldName1","'","''")))

        FieldType1 = trim(request(replace("FieldType1","'","''")))
		
		FieldLength1 = trim(request(replace("FieldLength1","'","''")))

      
		sql2 = "insert into ReconFile (FieldName, FieldType, FieldLength) "

        sql2 = sql2 & "values ( '"& FieldName1 & "', '"& FieldType1 & "', "& FieldLength1 &")"

        Conn.Execute sql2


        sql21 = "Alter Table StockReconciliation Add " & FieldName1 & " nvarchar(" & FieldLength1 & ")"
  
      	Conn.Execute sql21


        Message =  "The field ("&FieldName1&") was added."
	
		set acres=nothing
end if

' Modify the Depot
'*******************
if trim(request("action_button")) = "modify depot" then

	DepotFolder2 = split(trim(request.form("DepotFolder2")),",")

    FileType2 = split(trim(request.form("FileType2")),",")

    FirstRow2 = split(trim(request.form("FirstRow2")),",")

    Delimiter2 = split(trim(request.form("Delimiter2")),",")

	mid4 = split(trim(request.form("mid4")),",")


	for i=0 to ubound(mid4)
	
		strsql="Update ReconDepotFolder set DepotFolder = '"

        strsql= strsql & trim(replace(DepotFolder2(i),"'","''")) 

        strsql= strsql & "' , FileType ='"

        strsql= strsql & trim(replace(FileType2(i),"'","''")) 

        strsql= strsql & "' , FirstRow ="& trim(FirstRow2(i)) 

        strsql= strsql & " , Delimiter ="& trim(Delimiter2(i)) 

        strsql= strsql & " where DepotID= "& trim(mid4(i))

        'response.write strsql
		
	    conn.execute strsql 

        Message =  "Done."

	next
	

	
end if

' Modify field
'*******************
if trim(request("action_button")) = "modify field" then

	FieldLength2 = split(trim(request.form("FieldLength2")),",")

	FieldName2 = split(trim(request.form("FieldName2")),",")

    FieldType2 = trim(request(replace("FieldType2","'","''")))

	
	mid7 = split(trim(request.form("mid7")),",")


	for i=0 to ubound(mid7)
	
		sql6="Update ReconFile set FieldType = '"&  trim(replace(FieldType2(i),"'","''")) &"', FieldLength = '"& trim(replace(FieldLength2(i),"'","''")) &"' where FieldID="& trim(mid7(i))
		
	    conn.execute sql6 

        Message =  "Done."

	next



	
end if


' Delete Depot
'***************
if trim(request("action_button")) = "delete depot" then

	delid = split(trim(request("id4")),",")

	for j=0 to ubound(delid)
	
		sql4 = "Delete ReconDepotFolder where DepotID="& trim(delid(j))
	
	    conn.execute sql4 
	next
	
end if

' Delete Field
'***************
if trim(request("action_button")) = "delete field" then

	delid = split(trim(request("id6")),",")


	for j=0 to ubound(delid)

        sql6 = "Select FieldName from ReconFile Where FieldID ="&  delid(j) 

        Set Rs6 = Conn.Execute(sql6)

		sql61 = "Alter Table StockReconciliation drop column " & Rs6("FieldName") & "; Delete ReconFile where FieldID="& trim(delid(j))
	
	    conn.execute sql61 
	next

  	
        Message =  "Done."


	
end if

' add field to  Depot 
'********************
if trim(request("action_button")) = "append field" then

        DepotID = trim(Request("DepotFileID"))

        FieldID = Trim(Request("AppendField"))

        sqlc = "Select count(*) as tcount from ReconFileOrder where DepotID = "&DepotId

        Set Rsc1 = Conn.Execute(sqlc)

        tcount = Rsc1("tcount") + 1

		sql8 = "Insert into ReconFileOrder (DepotID, FieldID, Priority) Values ("&depotid&","&FieldID&","&tcount&")"
	    
        conn.execute sql8
	
end if


' remove field from Depot 
'************************
if trim(request("action_button")) = "remove field" then

        DepotID = trim(Request("DepotFileID"))

        FieldID = Trim(Request("RemoveField"))

		sql12 = "delete from ReconFileOrder where depotid= "&depotid&" and FieldID="&FieldID
	    
        conn.execute sql12
	
end if




' sorting the menu
'*******************

if trim(request("action_button")) = "MenuUp" then

	FieldID = trim(request("RemoveField"))

    DepotID = trim(Request("DepotID"))
	
        sql1 = "Select Priority - 1 as Priority From ReconFileOrder Where FieldID="&FieldID&" and DepotID="&DepotID

        Set Rs1 = Conn.Execute(sql1)

        'response.write rs1("priority")
        'response.write depotid

        sql2 = "Select FieldID From ReconFileOrder Where Priority="&Rs1("Priority")&" and DepotID="&DepotID

        Set Rs2 = Conn.Execute(sql2)

        'response.write rs2("Fieldid")

	
	    sql3="Update ReconFileOrder set Priority = Priority - 1 where DepotID="&DepotID&" and FieldId="&FieldID
		
	    conn.execute(sql3) 

        sql4="Update ReconFileOrder set Priority = Priority + 1 where DepotID="&DepotID&" and FieldId="&rs2("Fieldid")

        conn.execute(sql4)

	
end if

' sorting the menu
'*********************

if trim(request("action_button")) = "MenuDown" then

	FieldID = trim(request("RemoveField"))

    DepotID = trim(Request("DepotID"))
	
        sql1 = "Select Priority + 1 as Priority From ReconFileOrder Where FieldID="&FieldID&" and DepotID="&DepotID

        Set Rs1 = Conn.Execute(sql1)

        sql2 = "Select FieldID From ReconFileOrder Where Priority="&Rs1("Priority")&" and DepotID="&DepotID

        Set Rs2 = Conn.Execute(sql2)

        sql3="Update ReconFileOrder set Priority = Priority + 1 where DepotID="&DepotID&" and FieldId="&FieldID
	
        conn.execute(sql3) 

         sql4="Update ReconFileOrder set Priority = Priority - 1 where DepotID="&DepotID&" and FieldId="&rs2("Fieldid")


        conn.execute(sql4)

	
end if

' sorting the menu
'*********************

if trim(request("action_button")) = "CreateProfile" then


     DepotID = trim(Request("DepotID"))

     ' Check if view exist
     sql = "Select count(*) as count1 FROM sys.views WHERE name = 'vw_"&DepotID&"'"

    'response.write sql

     Set Rs = Conn.Execute(sql)

     If Rs("count1") = 1 then

         sqv_d = "drop view vw_"&DepotID&""

          Conn.Execute(sqv_d)

     End if

	
     Sql2 = "Select f.depotid, fieldname from  (ReconDepotFolder f join reconfileorder o "

     Sql2 = Sql2 & "on f.depotid = o.depotid) join ReconFile r on o.fieldid = r.fieldid "

     Sql2 = Sql2 & " and f.depotid=" &DepotID

     Sql2 = Sql2 & "order by f.depotid, o.priority desc"

     Set Rs2 = Conn.Execute(Sql2)

     Do While Not Rs2.EoF

     FieldName =  Rs2("fieldname") & "," & FieldName

     Rs2.MoveNext

     Loop 

     FieldName = Left(FieldName,Len(FieldName)-1) 

     'Response.write FieldName

     sqv = "create view vw_"&DepotID&" as select "&FieldName&" from StockReconciliation"

     'response.write sqv & "<br/>"

     Conn.Execute(sqv)
  
     Message = "Profile Created."
	
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

function addDepot()
{
       

	if(document.fm1.DepotName1.value =="")
       {
		alert("Please enter the Depot Name!");
        document.fm1.DepotName1.focus();
        return false;
       }
    if(document.fm1.DepotFolder1.value =="")
       {
		alert("Please enter the Depot Folder!");
        document.fm1.DepotFolder1.focus();
        return false;
       }
 
      if(document.fm1.FileType1.value =="")
       {
		alert("Please enter the file extension!");
        document.fm1.FileType1.focus();
        return false;
       }
 if(document.fm1.FirstRow1.value =="")
       {
		alert("Please enter first row value!");
        document.fm1.FirstRow1.focus();
        return false;
       }
  if(document.fm1.Delimiter1.value =="")
       {
		alert("Please enter the delimiter!");
        document.fm1.Delimiter1.focus();
        return false;
       }
	else
		{
		document.fm1.action_button.value="add section";
		document.fm1.submit();
		}
}

function addField()
{
	if(document.fm1.FieldName1.value =="")
       {
		alert("Please enter the Field Name!");
        document.fm1.FieldName1.focus();
        return false;
       }
     if(document.fm1.FieldType1.value =="")
       {
		alert("Please enter the Field Type!");
        document.fm1.FieldType1.focus();
        return false;
       }
     if(document.fm1.FieldLength1.value =="")
       {
		alert("Please enter the Field Length!");
        document.fm1.FieldLength1.focus();
        return false;
       }
    
	else
		{
		document.fm1.action_button.value="add field";
		document.fm1.submit();
		}
}

function EditDepot()
{
	  
		{
		document.fm1.action_button.value="modify depot";
		document.fm1.submit();
		}
}

function EditField()
{
	
		{
		document.fm1.action_button.value="modify field";
		document.fm1.submit();
		}
}

function appendField()
{
        if(document.getElementById('appendfield').selectedIndex == -1){
        alert("You must select a record!");
        return false;
        }
        {
		document.fm1.action_button.value="append field";
		document.fm1.submit();
		}
	
}

function removeField()
{
	  if(document.getElementById('removefield').selectedIndex == -1){
        alert("You must select a record!");
        return false;
        }
		{
		document.fm1.action_button.value="remove field";
		document.fm1.submit();
		}
}

function DeleteDepot(){
k=0;
document.fm1.action="ReconciliationSetup.asp?sid=<%=SessionID%>&FunctionID=<%=FunctionID%>&UserLevel=<%=UserLevel%>";
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
    document.fm1.action_button.value="delete depot";
    document.fm1.submit();
   }
 }

}


function DeleteField(){
k=0;
document.fm1.action="ReconciliationSetup.asp?sid=<%=SessionID%>&FunctionID=<%=FunctionID%>&UserLevel=<%=UserLevel%>";
	if (document.fm1.id6!=null)
	{
		for(i=0;i<document.fm1.id6.length;i++)
		{
			if(document.fm1.id6[i].checked)
			  {
			  k=1;
			  i=1;
			  break;
			  }
		}
		if(i==0)
		{
			if (!document.fm1.id6.checked)
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
    document.fm1.action_button.value="delete field";
    document.fm1.submit();
   }
 }

}

function doChange(flag)
{
       {
		document.fm1.submit();
		}
}


function doUp()
{
        if(document.getElementById('removefield').selectedIndex == -1){
        alert("Please select a option in the menu!");
        return false;
        }
        
        if(document.getElementById('removefield').selectedIndex == 0){
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
        if(document.getElementById('removefield').selectedIndex == -1){
        alert("Please select a option in the menu!");
        return false;
        }

       var x = document.getElementById('removefield').length;
       x = x-1;
       if(document.getElementById('removefield').selectedIndex == x){
        alert("It is already at the bottom position in the menu!");
        return false;
        }

        {
		document.fm1.action_button.value="MenuDown";
		document.fm1.submit();
		}
	
}

function createProfile()
{
       

        {
		document.fm1.action_button.value="CreateProfile";
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
 <table width="90%" border="0" class="normal">

    <tr> 
      <td  class="BlueClr" <% If FunctionID <> 1 Then %>bgcolor="#C0C0C0"<% End If %> width="20%">
      <a href="ReconciliationSetup.asp?sid=<% =SessionID %>&functionid=1">Depot Folder and Import File Information
      </a>
      </td> 
      <td  class="BlueClr" width="20%" <% If FunctionID <> 2 Then %>bgcolor="#C0C0C0"<% End If %> width="25%">
      <a href="ReconciliationSetup.asp?sid=<% =SessionID %>&functionid=2">Edit/Delete Depot Folder and File Info
      </a></td>
      <td  class="BlueClr" width="20%" <% If FunctionID <> 3 Then %>bgcolor="#C0C0C0"<% End If %> width="25%">
      <a href="ReconciliationSetup.asp?sid=<% =SessionID %>&functionid=3">Add Field in Master Table
      </a></td>
      <td  class="BlueClr" width="20%" <% If FunctionID <> 4 Then %>bgcolor="#C0C0C0"<% End If %> width="25%">
      <a href="ReconciliationSetup.asp?sid=<% =SessionID %>&functionid=4">Delete field in Master Table
      </a></td>
      <td  class="BlueClr" width="20%" <% If FunctionID <> 5 Then %>bgcolor="#C0C0C0"<% End If %> width="25%">
      <a href="ReconciliationSetup.asp?sid=<% =SessionID %>&functionid=5">Import File Fileds matching
      </a></td>
    </tr>

   </table>

<br>

<form name="fm1" method="post" action="">

<%  
    '---- Start of Add Depot ----
%>

  <% If FunctionID <> 1 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>


<table width="80%" border="0" cellpadding="4" class="normal">


<% 'Market

   set RsMarket = server.createobject("adodb.recordset")
   RsMarket.open ("Exec Retrieve_AvailableMarket") ,  StrCnn,3,1

%>

 <tr>
	<td >Market:</td> 
	<td >
	 
	<select size="1" name="Market" class="common">
			<%
					do while (Not rsMarket.EOF)
			%>
					<option value="<%=rsMarket("Market")%>"><%=rsMarket("Market")%></option>
			
			<%
					rsMarket.movenext
					Loop
			%>
	</select></td>
 </tr>
 
    
 <tr> 
      <td width="27%">
Depot Name</td> 
      <td width="69%">
      <Input name="DepotName1" type=text value="" size="50"></td>
    </tr>


 <tr> 
      <td width="27%">
Depot Folder</td> 
      <td width="69%">
      <Input name="DepotFolder1" type=text value="" size="80">&nbsp;
       </td>
    </tr>

<tr> 
      <td width="27%">
File Extension</td> 
      <td width="69%">
      <Input name="FileType1" type=text value="" size="80">&nbsp;


       </td>
    </tr>

<tr> 
      <td width="27%">
First row of data</td> 
      <td width="69%">
      <Input name="FirstRow1" type=text value="" size="80">&nbsp;


       </td>
    </tr>

<tr> 
      <td width="27%">
Delimiter</td> 
      <td width="69%">
      	<select size="1" name="Delimiter1" class="common">
		
		  <option value="0">Comma</option>
			
		  <option value="1">|</option>
	
	      <option value="2">Tab</option>

          <option value="3">Fixed Width</option>
	
	    </select>

       </td>
    </tr>
 
 <tr> 
      <td width="27%">
¡@</td> 
      <td width="69%">
      	<input type="button" value="    Add    " onClick="javascript:addDepot();">
         <input type="hidden" name="action_button" value="">   
¡@</td>
    </tr>

<tr> 
      <td  align =center><font color="red"><% = Message %></font></td> 
      <td >
¡@</td>
    </tr>

    </table>
</div>

<%

  '---- End of Add Section ------>

  '---- Start of Modify/Delete Depot ----
%>

  <% If FunctionID <> 2 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>
  
<table width="90%" border="0" class="normal">

	<tr> 
            
			<td width="30%" bgcolor="#FFFFCC">Depot Name</td> 
			<td width="30%" bgcolor="#FFFFCC">Folder</td>
         	<td width="10%" bgcolor="#FFFFCC">File Extension</td>
	        <td width="10%" bgcolor="#FFFFCC">First Row</td>
     	    <td width="10%" bgcolor="#FFFFCC">Delimiter</td>
            <td width="10%" bgcolor="#FFFFCC">Delete</td>
	</tr>



<%
        
		sql4 = " Select * From ReconDepotFolder order by DepotName Asc"

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
      <td colspan="6">
        
          <%call showpageno(pageno)%>
          
      </td>
    </tr>
<%		
		If Not Rs4.EoF Then

             		
			Do While Not Rs4.EoF

       
		
%>
<tr> 
	<td width="20%"><% = Rs4("Market") %> - <% = Rs4("DepotName") %>
			¡@</td>
<td width="20%">
<Input type="text" name="DepotFolder2" value="<% = Rs4("DepotFolder") %>" size="40">
			¡@</td>
<td>
<Input type="text" name="FileType2" value="<% = Rs4("FileType") %>" size="3">
</td>

<td>
<Input type="text" name="FirstRow2" value="<% = Rs4("FirstRow") %>" size="3">
</td>

<td>
<select size="1" name="Delimiter2" class="common">
		
		  <option value="0" <%if trim(Rs4("Delimiter"))=0 then%>Selected<%end if%>>Comma</option>
			
		  <option value="1" <%if trim(Rs4("Delimiter"))=1 then%>Selected<%end if%>>|</option>
	
	      <option value="2" <%if trim(Rs4("Delimiter"))=2 then%>Selected<%end if%>>Tab</option>

          <option value="3" <%if trim(Rs4("Delimiter"))=3 then%>Selected<%end if%>>Fixed Width</option>
	
	    </select>
</td>

<td><input type="checkbox" name="id4" value="<% = Rs4("DepotID") %>"></td> 
	</tr>
<input type="hidden" name="mid4" value="<% = Rs4("DepotID") %>">	
<%
	Rs4.movenext 

      

	   loop 
 
	End If
%>
<tr bgcolor="#ffffff"> 
      <td colspan="6">
        <div align="right"> 
          <%call showpageno(pageno)%>
          </div>
      </td>
    </tr>

<tr> 
      <td colspan="6" align =center><font color="red"><% = Message %></font></td> 
      
    </tr>
	
      <tr> 
      <td colspan=4></td> 
      <td >
      <input type="button" value="Edit" onClick="EditDepot();"></td>
      <td >
      <input type="button" value="Delete" onClick="DeleteDepot();"></td>
    </tr>

</table> 

</div>
               

<%
' End of Edit/Delete Depot 
' Start of Add Field Section 

   If FunctionID <> 3 Then 
%>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>


<table width="80%" border="0" cellpadding="4" class="normal">
    
 <tr> 
      <td width="27%">
Field Name</td> 
      <td width="69%">
      <Input name="FieldName1" type=text value="" size="50"></td>
    </tr>

<tr> 
      <td width="27%">
Field Type</td> 
      <td width="69%">
      <Input name="FieldType1" type=text value="" size="50">
       </td>
    </tr>
 <tr> 
      <td width="27%">
Field Length</td> 
      <td width="69%">
      <Input name="FieldLength1" type=text value="" size="50">
       </td>
    </tr>
 <tr> 
      <td width="27%">
¡@</td> 
      <td width="69%">
      	<input type="button" value="    Add    " onClick="javascript:addField();">
¡@</td>
    </tr>

<tr> 
      <td colspan="2" align =center><font color="red"><% = Message %></font></td> 
      <td >
¡@</td>
    </tr>

    </table>
</div>

<% 

   '---- End of Add Section ------>

   '---- Start of Modify/Delete Field ---- 
%>

  <% If FunctionID <> 4 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>
  
<table width="90%" border="0" class="normal">

	<tr> 
			<td width="30%" bgcolor="#FFFFCC">Field Name</td>
            <td width="30%" bgcolor="#FFFFCC">Field Name</td>  
			<td width="20%" bgcolor="#FFFFCC">Length</td>
            <td width="20%" bgcolor="#FFFFCC">Delete</td>
	</tr>



<%
        
		sql5 = " Select * From ReconFile order by FieldID Asc"

	  
				set Rs5 = conn.execute(sql5)
				 
				   
	
%>
<tr bgcolor="#ffffff"> 
      <td colspan="4">
        
          <%call showpageno(pageno)%>
          
      </td>
    </tr>
<%		
		If Not Rs5.EoF Then

             		
			Do While Not Rs5.EoF

      
		
%>
<tr> 
	<td width="30%"><% = Rs5("FieldID") %>. <% = Rs5("FieldName") %>
			¡@</td>

<td width="30%">
<input type=text name="FieldType2" value ="<%= Rs5("FieldType") %>"  readonly>
			¡@</td>

<td width="20%">
<input type=text name="FieldLength2" value ="<%= Rs5("FieldLength") %>" readonly>
			¡@</td>
<td><input type="radio" name="id6" value="<% = Rs5("FieldID") %>"></td> 
	</tr>
<input type="hidden" name="mid7" value="<% = Rs5("FieldID") %>">	
<%
	Rs5.movenext 

      

	   loop 
 
	End If
%>


<tr> 
      <td colspan="3" align =center><font color="red"><% = Message %></font></td> 
      <td >
¡@</td>
    </tr>
	
      <tr> 
      <td colspan=2></td> 
      <td width="30%"></td>
      <td width="30%">
      <input type="button" value="Delete" onClick="DeleteField();"></td>
    </tr>

</table> 

</div>
              

<% '---- End of Edit/Delete Field ----- 


   '---- Start of Depot Vs File Setup ---- 
%>

  <% If FunctionID <> 5 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>



<%
        DepotID = Request("DepotID")

        'Response.write DepotID

        DepotID = Request("DepotID")
        
		sql8 = " Select * From ReconDepotFolder order by DepotName Asc"
	   
		set Rs8 = conn.execute(sql8)
		 
        
%>

<select name="DepotID" class="common"  size="1" onchange="doChange(1)">
          <% 
                             If Not Rs8.EoF Then

                        Rs8.MoveFirst

                             If DepotID = "" Then

                             DepotID = Rs8("DepotID")

                             End If

							do while not Rs8.eof

           %>               

<option value="<% =Rs8("DepotID") %>" <%If DepotID=trim(Rs8("DepotID")) Then%>Selected<%End If%>><% =trim(Rs8("DepotID")) %>. <% =trim(Rs8("Market")) %> - <% =trim(Rs8("DepotName")) %></option>

                  <%

                               Rs8.movenext

							loop
						
						End if
					%>
</select>

<br/>
  
<table width="90%" border="0" class="normal">

	<tr> 
			
			<td width="35%" align="center" bgcolor="#FFFFCC">Field available</td>
            <td width="35%" align="center" bgcolor="#FFFFCC">Current Field</td>
            <td width="30%" align="center" bgcolor="#FFFFCC"></td>
	</tr>



<%
       
           sql9 = " Select * From ReconFile Order by FieldName Asc"

           Set Rs9 = Conn.Execute(sql9)

           If Not Rs9.EoF Then
                      
%>


<tr> 

<td align="center">

	<select size="20" name="appendfield" id="appendfield" class="common">

<%
                        Rs9.MoveFirst

							do while not Rs9.eof
%>

     <option value="<%=Rs9("FieldID")%>" >&nbsp;&nbsp;<% = Rs9("FieldName") %>&nbsp;&nbsp;&nbsp;&nbsp;</option>

<%

                   Rs9.movenext

							loop
						
						End if
					%>


    </select>


		
        
			¡@</td>

<%

         sql10 = " Select * from ReconFileOrder o , ReconFile r "

         sql10 = sql10 & " where o.FieldID = r.FieldID and o.DepotID = " & DepotID

         sql10 = sql10 & " order by o.priority asc"

         'response.write sql10

         Set Rs10 = Conn.Execute(sql10)

         


%>
         
<td align="center">

        <select size="20" name="RemoveField" id="removefield" class="common">

<%
                        If Not Rs10.EoF Then
                          
                         Rs10.MoveFirst

							do while not Rs10.eof
%>

     <option value="<%=Rs10("FieldID")%>" >&nbsp;&nbsp;<% = Rs10("FieldName") %>&nbsp;&nbsp;&nbsp;&nbsp;</option>

<%

                   Rs10.movenext

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
<tr> 
      <td colspan="3">
      </td> 
      
    </tr>
<tr> 
      <td align="center">
			<input type="button" value="Append Field to Depot" onClick="appendField();">
      </td> 
      <td align="center">
      
			<input type="button" value="Remove Field from Depot" onClick="removeField();">
¡@</td>
  <td></td>
    </tr>
	<tr> 
      <td colspan="3">
      </td> 
      
    </tr>
<tr> 
<tr> 
      <td>
			
      </td> 
      <td colspan="2">

			<input type="button" value="  Create Profile  " onClick="createProfile();">
¡@</td>
    </tr>
	

<tr> 
      <td colspan="2" align =center><font color="red"><% = Message %></font></td> 
      <td >
¡@</td>
    </tr>
	
     
   <input type="hidden" name="DepotFileID" value="<% = DepotID %>">	

</table> 

</div>
              

<% '---- End of Depot Vs File Setup ----- %>

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
		     response.write "<a href='ReconciliationSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='ReconciliationSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;"
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
				response.write "<a href='ReconciliationSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&i&"'>"&cstr(i)&"</a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\10)<(lastpage\10) then
		        response.write "<a href='ReconciliationSetup.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&(pageno+1-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;"
			   response.write "<a href='ReconciliationSetp.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="& (lastpage) &"'>Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>