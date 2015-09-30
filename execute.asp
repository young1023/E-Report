
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

%>

<%
'------------------------------------------------------------------------------------
'
'
' The function of this page is to add and modify data into the records.
'
'
'------------------------------------------------------------------------------------

response.expires=0
flag=trim(request.form("whatdo"))
'Response.write flag
'Response. End

'-----------------------------------------------------------------------------------------------------
'
'
' Adding the member information
'
'
'----------------------------------------------------------------------------------------------------- 


If flag="added member" then

  ' Get System Setting
  '------------------------------------------------------------------------
   'SQL = "Select Setting Value From Sys
   'Set Rs = Conn.Execute(SQL)
	 
 '  Passwordlength = Trim(Rs("Passwordlength"))
	 
'   PasswordExpireDays = Trim(Rs("PasswordExpireDays"))
	 
   'ExpireDay = DateADD(day, "&PasswordExpireDays&", getdate())
   
   ulevel=trim(request.form("UserLevel"))
   LoginName = replace(trim(request.form("LoginName")),"'","''")
   UserName = replace(trim(request.form("UserName")),"'","''")
   Email = replace(trim(request.form("email")),"'","''")
   Password = replace(trim(request.form("Password")),"'","''")
   udepartment=replace(trim(request.form("dept")),"'","''")
   MemberID = Session("MemberID")
	 CreationDate=FormatDateTime(now(),0)

 
   if email="" then
     email="No email"
   end if

    ' Check if the Login Name is existed
  sql="Select LoginName From Member Where LoginName = '"&LoginName&"' "
  set rs=conn.execute(sql)
  
  If Not Rs.EoF Then
  
     Message = "The login name <u><b>"&LoginName&"</b></u> is existed!"
     
  Else
  
     
     SQL1 = "Insert Into Member (name, loginname, password, email, dept, lock, userlevel, groupid, lastpasswordchangedate) "
     
     SQL1 = SQL1 & " Values ('"&UserName&"', '"&LoginName&"', '"&Password&"','"&Email&"','"&udepartment&"',0,"&ulevel&", 0, '1-1-1900' )"
  
  'response.write sql1
  'response.end
     Conn.Execute(SQL1)

  'refresh group table
  'SQL = "exec Process_DistinctTable"
  'Conn.Execute(SQL)
  
  
  
'response.write sql1
' Write system log
  Message = flag & " <b>" & UserName & "</b> from " &  Request.ServerVariables("REMOTE_ADDR")
 
  Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
  
  End if
  rs.close
  set rs=nothing
  whatgo="sa_member.asp?sid=" & SessionID


'-------------------------------------------------------------------------------------
'
'
' Updating the member information
'
'
'-------------------------------------------------------------------------------------

ElseIf flag="updated member" then   

  UserName = replace(trim(request.form("Name")),"'","''")
  Department = replace(trim(request.form("Department")),"'","''")  
  UserName2 = replace(trim(request.form("Name2")),"'","''")
  Department2 = replace(trim(request.form("Department2")),"'","''")  

  Password = replace(trim(request.form("Password")),"'","''") 
    
  Email = replace(trim(request.form("Email")),"'","''")
  Email2 = replace(trim(request.form("Email2")),"'","''")
  
  UserLevel = replace(trim(request.form("UserLevel")),"'","''")
  
  sql1 = " Select LevelName From UserLevel where LevelNumber ="&UserLevel
  Set Rs1 = Conn.Execute(sql1)
  LevelName = Trim(Rs1("LevelName"))
  
  LevelName2 = replace(trim(request.form("LevelName2")),"'","''")
  sql2 = " Select LevelName From UserLevel where LevelNumber ="&LevelName2
  Set Rs2 = Conn.Execute(sql2)
  LevelName2 = Trim(Rs2("LevelName"))
  
  
  SystemEmail = replace(trim(request.form("SystemEmail")),"'","''")
  
  Lock = replace(trim(request.form("Lock")),"'","''")
  Lock2 = replace(trim(request.form("Lock2")),"'","''")
  
  MemberID = Request("MemberID")
  UserID = Session("MemberID")
          
  SQL = "Update Member Set Name='"&UserName&"', "
  
  SQL = SQL & " Dept = '"&Department&"', "
  
  SQL = SQL & "Email='"&Email&"', "
  
  SQL = SQL & "Lock = "&Lock&", "
  
  SQL = SQL & "UserLevel = "&Userlevel&" where MemberID ="&MemberID
  
  conn.execute(sql)
  
  
  
  ' Write system log
  Message = flag & " <b>" & UserName2 & "</b> "

  If UserName2 <> UserName Then

  Message = Message & "to " & UserName 

  End If

  If Department2 <> Department Then
  
  Message = Message & " <b>Dept:</b> " & Department2 & " to " & Department 

  End If

  If Email2 <> Email Then
  
  Message = Message & " <b>Email:</b> " & Email2 & " to " & Email 

  End If

  If LevelName2 <> LevelName Then

  Message = Message & " <b>UserLevel:</b> " & LevelName2 & " to " & LevelName 

  End If

  Message = Message & " at " &  Request.ServerVariables("REMOTE_ADDR")
  
  Conn.Execute ("Exec AddLog '"&Message&"', "&UserID&"") 

  
  whatgo="MemberDetail.asp?sid=" & SessionID & "&MemberID="&MemberID
  

'----------------------------------------------------------------------------------------------------
'
'  Add group
'
'----------------------------------------------------------------------------------------------------

ElseIf flag="added group" Then

  	Description = replace(trim(request.form("Description")),"'","''")
  	Name = replace(trim(request.form("Name")),"'","''")
  	SharingGroup = replace(trim(request.form("SharingGroup")),"'","''")


 	 MemberID = Session("MemberID")
    Sql4="Insert into UserGroup (Name, Description, Sharing) "
    Sql4 = Sql4 & "Values ('"&Name&"', '"&Description&"', "&SharingGroup&")"
    
    Conn.Execute(Sql4)
    
  
	  'refresh group table
	  SQL = "exec Process_DistinctTable"
	  Conn.Execute(SQL)

      
  	Message = flag & " <b> " & Name & " </b> at " &  Request.ServerVariables("REMOTE_ADDR")

  	Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
 
  	whatgo="sa_group.asp?sid=" & SessionID
  
'----------------------------------------------------------------------------------
'
'
' Modifying User Group
'
'
'-----------------------------------------------------------------------------------
  
ElseIf flag="updated group" Then
  
  id=replace(trim(request.form("id")),"'","''")
  Description = replace(trim(request.form("Description")),"'","''")
  Name = replace(trim(request.form("Name")),"'","''")
  OldName = replace(trim(request.form("OldName")),"'","''")
  OldDesc = replace(trim(request.form("OldDesc")),"'","''")
  SharingGroup = replace(trim(request.form("SharingGroup")),"'","''")
  OldSharing = replace(trim(request.form("OldSharing")),"'","''")



  MemberID = Session("MemberID")


  SQL = "Update UserGroup set Description='"&Description&"'"
  SQL = SQL & ", Name='"&Name&"'"
  SQL = SQL &  ", Sharing = "&SharingGroup
  SQL = SQL & " where GroupID="&id
  
  Conn.Execute(SQL)
  
  'refresh group table
  SQL = "exec Process_DistinctTable"
  Conn.Execute(SQL)
  
  
    ' Write system log

  Message = flag & " <b>" & Name & "</b> " 

  If OldName <> Name Then
  
  Message = Message & " <b>Group Name:</b> " & OldName & " to " & Name 

  End If

  If Description <> OldDesc Then
  
  Message = Message & " <b>Description:</b> " & OldDesc & " to " & Description 

  End If

  If SharingGroup = 0 Then

  SharingGroup = "No"

  Else

  SharingGroup = "Yes"

  End If

  If OldSharing <> SharingGroup Then
  
  Message = Message & " <b>Sharing Group:</b> " & SharingGroup 

  End If

  Message = Message & " from " &  Request.ServerVariables("REMOTE_ADDR")
 
  Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
  
  whatgo="UserGroup.asp?functionid=add&sid=" & sessionID & "&id=" &ID

'----------------------------------------------------------------------------------
'
'
' Add member to group
'
'
'-----------------------------------------------------------------------------------
  
ElseIf flag="add member" Then
  
  id=replace(trim(request.form("id")),"'","''")
  ' get the group name
  Set Rs1  = Conn.Execute("Select Name From UserGroup Where GroupID ="&id)

  addmember=replace(trim(request.form("addmember")),"'","''")
  'get the user name
   Set Rs2  = Conn.Execute("Select Name From Member Where MemberID ="&addmember)
   
  MemberID = Session("MemberID")


  SQL = "Update member set groupid='"&id&"'"
  SQL = SQL & " where memberid=" & addmember
  
  Conn.Execute(SQL)
  'response.write addmember
  ' Write system log
  
  'refresh group table
  SQL = "exec Process_DistinctTable"
  Conn.Execute(SQL)
    

  Message = flag & " ( " & Rs2("Name") & " ) into group ( " & Rs1("Name") & " ) from " &  Request.ServerVariables("REMOTE_ADDR")
 
  Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
  
  whatgo="UserGroup.asp?functionid=add&sid=" & SessionID & "&id="&ID

'----------------------------------------------------------------------------------
'
'
' Add member to Shared group
'
'
'-----------------------------------------------------------------------------------
  
ElseIf flag="add shared member" Then
  
  id=replace(trim(request.form("id")),"'","''")
  ' get the group name
  Set Rs1  = Conn.Execute("Select Name From UserGroup Where GroupID ="&id)
    
  addmember=replace(trim(request.form("addmember")),"'","''")
  'get the user name
   Set Rs2  = Conn.Execute("Select Name From Member Where MemberID ="&addmember)


  MemberID = Session("MemberID")


  SQL = "insert into sharedgroup (sharedgroupid, memberid) values ("&id&", "&addmember&")  " 
  
  Conn.Execute(SQL)
  'response.write addmember
  ' Write system log
  


  Message = "Add member ( " & Rs2("Name") & " ) into share group ( " & Rs1("Name") & " ) from " &  Request.ServerVariables("REMOTE_ADDR")
 
  Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
  
  whatgo="UserGroup.asp?functionid=add&sid=" & SessionID & "&id="&ID
  
'----------------------------------------------------------------------------------
'
'
' Delete member from group
'
'
'-----------------------------------------------------------------------------------
  
ElseIf flag="remove member" Then
  
  id=replace(trim(request.form("id")),"'","''")
  ' get the group name
  Set Rs1  = Conn.Execute("Select Name From UserGroup Where GroupID ="&id)
  
  removemember=replace(trim(request.form("removemember")),"'","''")
  'get the user name
   Set Rs2  = Conn.Execute("Select Name From Member Where MemberID ="&removemember)

  MemberID = Session("MemberID")


  SQL = "Update member set groupid='0'"
  SQL = SQL & " where memberid=" & removemember
  
  Conn.Execute(SQL)
  'response.write addmember

  'refresh group table
  SQL = "exec Process_DistinctTable"
  Conn.Execute(SQL)
  

  ' Write system log
  Message = flag & " (" & Rs2("Name") & " ) from group ( " & Rs1("Name") & " ) at " & Request.ServerVariables("REMOTE_ADDR")
 
  Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
  
  whatgo="UserGroup.asp?sid=" & SessionID & "&id="&ID


'----------------------------------------------------------------------------------
'
'
' Delete member from Shared group
'
'
'-----------------------------------------------------------------------------------
  
ElseIf flag="remove shared member" Then
  
 id=replace(trim(request.form("id")),"'","''")
  ' get the group name
  Set Rs1  = Conn.Execute("Select Name From UserGroup Where GroupID ="&id)
  
  removemember=replace(trim(request.form("removemember")),"'","''")
  'get the user name
   Set Rs2  = Conn.Execute("Select Name From Member Where MemberID ="&removemember)

  MemberID = Session("MemberID")

  
  SQL = "delete from sharedgroup where sharedgroupid = " & ID
  SQL = SQL & " and memberid=" & removemember
  
  Conn.Execute(SQL)
  'response.write addmember



  ' Write system log
  Message = "removed member (" & Rs2("Name") & " ) from shared group ( " & Rs1("Name") & " ) at " & Request.ServerVariables("REMOTE_ADDR")
 
  Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
  
  whatgo="UserGroup.asp?sid=" & SessionID & "&id="&ID  
     
'---------------------------------------
'
' Modify Setup
'
'---------------------------------------  

ElseIf flag="modified setup" Then

  SystemTimeOut = replace(trim(request.form("SystemTimeOut")),"'","''")
  PasswordExpireDays = replace(trim(request.form("PasswordExpireDays")),"'","''")
  PasswordLength = replace(trim(request.form("PasswordLength")),"'","''")
  SMTPServer = replace(trim(request.form("SMTPServer")),"'","''")
  SenderName = replace(trim(request.form("SenderName")),"'","''")
  SenderEmail = replace(trim(request.form("SenderEmail")),"'","''")
  SMTPServer = replace(trim(request.form("SMTPServer")),"'","''")
  MemberID = Session("ID")
	CreationDate=now()
  SQL = "Update SystemSetup Set SystemTimeOut='"&SystemTimeOut&"'"
  SQL = SQL & ",PasswordExpireDays = '"&PasswordExpireDays&"'"
  SQL = SQL & ",PasswordLength = '"&PasswordLength&"' "
  SQL = SQL & ",SMTPServer = '"&SMTPServer&"' "
  SQL = SQL & ",SenderEmail = '"&SenderEmail&"' "

  Conn.Execute(sql)
  
 ' Write system log
  Message = flag & " from " &  Request.ServerVariables("REMOTE_ADDR")
 
  Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 

  whatgo="Setup.asp"

'==========================================================================
'
'
'  Unlock User Account
'
'
'==========================================================================


ElseIf  flag = "unlocked account" Then
 
  delid=split(trim(request.form("mid")),",")
  username = trim(request.form("username"))
  MemberID = Session("Memberid")
  
   for i=0 to Ubound(delid)
     sql="Update Member Set Lock = 0 where MemberID="&trim(delid(i))
     conn.execute(sql)
	 'response.write sql&"<br>"
   next
   

  Message = flag & " <b>" & username & "</b> at " & Request.ServerVariables("REMOTE_ADDR") 
 
  Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 

   whatgo="unlock.asp?sid=" & SessionID
  
'---------------------------------------
'
'  Delete the Audit logs
'
'---------------------------------------  
ElseIf flag ="deleted logs" Then

	DeleteLog = replace(trim(request.form("DeleteLog")),"'","''")
	MemberID = Session("MemberID")
	
	SQL = "Delete From AuditLog where CreateDate < DateADD(day, -"&DeleteLog&", getdate())"
	
	Conn.Execute(SQL)
	
    ' Write Log
    Message = flag & " longer then " & DeleteLog & " days from " & Request.ServerVariables("REMOTE_ADDR") 
 
    Response.write ("Exec AddLog '"&Message&"', "&MemberID&"") 
    Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
	
    whatgo="Audit.asp?sid=" & SessionID

'--------------------------------------------------
'
'  Delete the Transaction Record (Downsize the db)
'
'--------------------------------------------------  
ElseIf flag ="deleted db" Then

	DeleteMonth = replace(trim(request.form("DeleteMonth")),"'","''")
	MemberID = Session("MemberID")
	
	SQL1 = "Delete From ClientStatement where StatementDate < DateADD(month, -"&DeleteMonth&", getdate())"
	
	Conn.Execute(SQL1)
	
    ' Write Log
    Message = flag & " longer then " & DeleteLog & " months from " & Request.ServerVariables("REMOTE_ADDR") 
 
    Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
	
    whatgo="Maintenance.asp?sid=" & SessionID



'---------------------------------------
'
'  Change Password
'
'---------------------------------------

ElseIf flag="changed password" then

  OldPassword = replace(trim(request.form("OldPassword")),"'","''")

  NewPassword1 = replace(trim(request.form("NewPassword1")),"'","''")
  
  MemberID = Session("ID")
  
  
  ' Check if the password is correct
    SQL3 = "Select * From Member Where Password = '"&OldPassword&"' and MemberID ="&MemberID
    
	Set Rs2 = Conn.Execute(SQL3) 

	If Rs2.EoF Then

      	Session("SystemMessage") = "Your password is not corrected, please try again."
      	
      	Response.Redirect("ChangePassword.asp")
      	
	End If
  
    ' Check if the password used before
  	SQL = "Select Password From PasswordControl Where Password = '"&NewPassword1&"' and MemberID = "&MemberID
  	Set RS = conn.execute(SQL)
  
  	If Not Rs.EoF Then
  	
  	Session("SystemMessage") = "The password was used before. Please use another combination."
  	
  	Response.Redirect("ChangePassword.asp")
  	
  	Else
  
    SQL1 = "Update Member set Password='"&NewPassword1&"' Where MemberID="&MemberID
    
    Conn.Execute(SQL1)
    
    SQL2 = "Insert Into PasswordControl (Password, MemberID) Values ('"&OldPassword&"',"&MemberID&")"
    
    Conn.Execute(SQL2)
    
    ' Write Log
    Message = flag  & " from " & Request.ServerVariables("REMOTE_ADDR") 
 
    Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"")
         
    whatgo="changepassword.asp"
    
  	End if
  
  rs.close
  set rs=nothing
  
Else

  ' For any error, redirect to home page
  Response.Redirect "UserDetail.asp"
  
End if


Conn.Close
Set Conn = Nothing


%>


<html>
<head>
<html>
<head>
<title>UOB Intranet</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Refresh" content="2; url='<%=whatgo%>'">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
</head>

<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">


<div align="center">


<table border=0 cellpadding=3 cellspacing=0 width="90%" class=Normal height="100">

  <tr> 
    <td align="center" height="50"><% = Message %></td>
  <tr>
    <td align="center" height="50"><br><a href="<%=whatgo%>">Return</a></td>
  </tr>
  
</table>
            
</div>
            
</div>            
</body>
</html>