

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

  ulevel=trim(request.form("UserLevel"))
  LoginName = replace(trim(request.form("LoginName")),"'","''")
  UserName = replace(trim(request.form("UserName")),"'","''")
  Email = replace(trim(request.form("email")),"'","''")
    Password = replace(trim(request.form("Password")),"'","''")
  udepartment=replace(trim(request.form("dept")),"'","''")
 
  'if email="" then
    'email="No email"
 ' end if

    ' Check if the Login Name is existed
  sql="Select LoginName From Member Where LoginName = '"&LoginName&"' "
  set rs=conn.execute(sql)
  if not rs.eof then
     message="The Login Name <u><b>"&LoginName&"</b></u> is existed!"
  else
     sql="Insert into member (LoginName,Email,UserLevel, Name) values('"&LoginName&"','"&Email&"',"&ulevel&",'"&UserName&"')"
  response.write sql
  'response.end
     conn.execute(sql)
     
     
     message="The member is added."
  end if
  rs.close
  set rs=nothing
  whatgo="sa_member.asp"


elseif flag="modifymember" then
  pid=trim(request.form("pid"))
  id=trim(request.form("id"))
  uname=replace(trim(request.form("name")),"'","''")
  employeenum=replace(trim(request.form("employeenum")),"'","''")    'e-mail
  employeenum2=replace(trim(request.form("employeenum2")),"'","''")
  indicate=replace(trim(request.form("indicate")),"'","''")

  if indicate="" then
    indicate=" "
  end if
  phone=replace(trim(request.form("phone")),"'","''")
  if phone="" then
    phone=" "
  end if

  sql="select id from member where employeenum='"&employeenum&"'"
  set rs=conn.execute(sql)
  if not rs.eof then
     if employeenum=employeenum2 then
       sql="Update member set name='"&uname&"',indicate='"&indicate&"',phone='"&phone&"',flag="&pid&" where id="&id
       conn.execute(sql)
       message="Modify Member Successfully"
     else
       message="The Employee Number <u><b>"&employeenum&"</b></u> is exist !"
     end if
  else
     sql="Update member set name='"&uname&"',pwd='"&employeenum&"',employeenum='"&employeenum&"',indicate='"&indicate&"',phone='"&phone&"',flag="&pid&" where id="&id&" "
     conn.execute(sql)
     message="Modify Member Successfully"
  end if
  rs.close
  set rs=nothing
  whatgo="sa_modify.asp?id="+id

'-------------------------------------------------------------------------------------
'
'
' Updating the member information
'
'
'-------------------------------------------------------------------------------------

ElseIf flag="Modify Member" then   

  name=replace(trim(request.form("name")),"'","''")
  LoginName = replace(trim(request.form("LoginName")),"'","''")   
  Email = replace(trim(request.form("Email")),"'","''")
    MemberID = Request.form("MemberID")
          
  sql="update member set name='"&name&"',Email='"&Email&"' where id="&MemberID
  response.write sql
  conn.execute(sql)
  whatgo="sa_member.asp?id=3"
  message="Update Successfully"

elseif flag="userpost" then
  modulnum=trim(request.form("modulnum"))
  if modulnum="" then
    conn.close
    set conn=nothing
    response.redirect "listbill.asp"
  end if


  terminal=replace(trim(request.form("terminal")),"'","''")
  if request.form("u_select")=1 then
    country=trim(request.form("u_region"))
    sql="select mnemonic from country where location='"&country&"'"
    set rs2=conn.execute(sql)
    if rs2.eof then
      sql="Insert into country(country,location,mnemonic) values('"&country&"','"&country&"','"&country&"')"
      conn.execute(sql)
    else
      country=rs2("mnemonic")
    end if
    rs2.close
    set rs2=nothing
  else
    country=trim(request.form("country"))
  end if

  money=split(trim(request.form("money")),",")
  money2=split(trim(request.form("money2")),",")
  for i=0 to Ubound(money)
    if trim(money(i))<>"" then
      p1=trim(money2(i))+","+p1
      p2=trim(money(i))+","+p2
    end if
  next
 if p1<>"" and p2<>"" and terminal<>"" then
  p1=p1&"userid,terminal,flag"
  p2=p2&session("shell_id")&",'"&terminal&"','"&country&"'"
  sql="Insert into "&modulnum&"_b("&p1&") values("&p2&")"
  'response.write sql
  'response.end
  conn.execute(sql)
  message="Post Value Successfully"
 else
  message="Post Value Error"
 end if
  whatgo="userpost.asp?id="&modulnum
  whatgo="window.close()"

'----------------------------------------------------------------------------------------------------
'
'  Add group
'
'----------------------------------------------------------------------------------------------------

ElseIf flag="added group" Then

  	Description = replace(trim(request.form("Description")),"'","''")
  	Name = replace(trim(request.form("Name")),"'","''")


    Sql4="Insert into UserGroup (Name, Description) "
    Sql4 = Sql4 & "Values ('"&Name&"', '"&Description&"')"
    

    Conn.Execute(Sql4)
    
    

  	whatgo="sa_group.asp"
  
'----------------------------------------------------------------------------------
'
'
' Modifying User Group
'
'
'-----------------------------------------------------------------------------------
  
elseif flag="modified group" Then
  
  id=replace(trim(request.form("id")),"'","''")
  Description = replace(trim(request.form("Description")),"'","''")
  Name = replace(trim(request.form("Name")),"'","''")


  SQL = "Update UserGroup set Description='"&Description&"'"
  SQL = SQL & ", Name='"&Name&"'"
  SQL  =  SQL & " where GroupID="&id

  conn.execute(sql)
  
  whatgo="UserGroup.asp?id="&ID
  message="Thank for your information"
  
     
'---------------------------------------
'
' Modify Setup
'
'---------------------------------------  

ElseIf flag="modified setup" Then

  PasswordExpireDays = replace(trim(request.form("PasswordExpireDays")),"'","''")
  PasswordLength = replace(trim(request.form("PasswordLength")),"'","''")
  SMTPServer = replace(trim(request.form("SMTPServer")),"'","''")
  SenderName = replace(trim(request.form("SenderName")),"'","''")
  SenderEmail = replace(trim(request.form("SenderEmail")),"'","''")
  SMTPServer = replace(trim(request.form("SMTPServer")),"'","''")

  SQL = "Update SystemSetup Set PasswordExpireDays = '"&PasswordExpireDays&"'"
  SQL = SQL & ",PasswordLength = '"&PasswordLength&"' "
  SQL = SQL & ",SMTPServer = '"&SMTPServer&"' "
  SQL = SQL & ",SenderEmail = '"&SenderEmail&"' "

  conn.execute(sql)

	whatgo="Setup.asp"
    message="Thank for your information"
  
'---------------------------------------
'
'  Delete the Audit logs
'
'---------------------------------------  
ElseIf flag ="deleted logs" Then

	DeleteLog = replace(trim(request.form("DeleteLog")),"'","''")
	MemberID = Session("ID")
	
	SQL = "Delete From AuditLog where CreateDate < DateADD(day, -"&DeleteLog&", getdate())"
	
	Conn.Execute(SQL)
	
    Message = flag & " longer then " & DeleteLog & " days from " & Request.ServerVariables("REMOTE_ADDR") 
 
    Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
	
    whatgo="Audit.asp"

   
'---------------------------------------
'
'  Audit Setup
'
'---------------------------------------  

elseif flag="AuditSelect" Then

    pageid=request.form("pageid")
 	delid=split(trim(request.form("mid")),",")
  	ResetID=split(trim(request.form("ResetID")),",")
	
 	   
     for i=0 to Ubound(ResetID)
     sql = "Update AuditSetup Set AuditType = 0 where AuditID="&trim(ResetID(i))
     conn.execute(sql)
	 next
 	
     for i=0 to Ubound(delid)
     sql2 = "Update AuditSetup Set AuditType = 1 where AuditID="&trim(delid(i))

     conn.execute(sql2)
     
       next

	whatgo="AuditSetup.asp?page_id=pm"
    message="Thank for your information"
  


'---------------------------------------
'
'  Change Password
'
'---------------------------------------

ElseIf flag="change password" then

  oldpwd = trim(request.form("oldpassword"))
  newpwd=trim(request.form("newpassword"))
  
  sql = "Select Password From PasswordControl Where Password = '"&newpwd&"' and MemberID = "&session("shell_id")&" "
  set rs=conn.execute(sql)
  
  If rs.eof then
    sql="update member set Password='"&newpwd&"' where Memberid="&Session("id")&" "
    conn.execute(sql)
    message="<font color=white>Change Password Successfully</font>"
     whatgo="changepassword.asp"
  else
    message="The password was used before."
    whatgo="changepassword.asp"
  end if
  rs.close
  set rs=nothing
else
  response.redirect "hsemis.asp?page_id=user"
end if


conn.close
set conn=nothing
wait=10

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Refresh" content="2; url='<%=whatgo%>'">
<link rel="stylesheet" type="text/css" href="hse.css" />
<title></title>
</head>
<body topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" >
<br><br>
<table border=0 cellpadding=3 cellspacing=0 class=hardcolor width="90%" align=center class=normal">
  <tbody>
  <tr> 
    <td align=center bgcolor="#006699" height="28" ><font color=white><%=message%></font></td>
  </tr>
  <tr>
    <td align=center height="38"><br><a href='<%=whatgo%>'>Return</a></td>
  </tr>
  </tbody>
</table>
</body>
</html>