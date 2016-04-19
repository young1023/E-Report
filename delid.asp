

<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


%>

<!--#include file ="include/SQLconn.inc.asp" -->
<%
response.expires=0
whatdo=trim(request.form("whatdo"))
'response.write whatdo
'response.end
pageid=trim(request("pageid"))
Select case whatdo
 case "del"
   delid=split(trim(request.form("mid")),",")
   for i=0 to Ubound(delid)
     sql="delete id from nomination where id="&delid(i)&" and employeenum='"&session("u_enum")&"' "
     conn.execute(sql)
	 'response.write sql&"<br>"
   next
   mymessage="Delete the Record Successfully"
   whatgo="user_edit.asp"

 
'==========================================================================
'
'
'  Delete member
'
'
'==========================================================================


 case "delmember"
 
    MemberID = Session("MemberID")
 
   delid=split(trim(request.form("mid")),",")
   
   for i=0 to Ubound(delid)
   
    ' get the Member name
    Set Rs1  = Conn.Execute("Select Name From Member Where MemberID ="&trim(delid(i)))
         
     Message = "deleted member <b>" & Rs1("Name") & "</b> at " & Request.ServerVariables("REMOTE_ADDR") 
 
     Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&" ") 

   
     sql = "delete from member where Memberid="&trim(delid(i))
     
     
     conn.execute(sql)
   next
  
  


   whatgo="sa_member.asp?sid=" & SessionID

'==========================================================================
'
'
'  Delete Group
'
'
'==========================================================================


 case "delete group"
 
        MemberID = Session("MemberID") 
        
       delid=split(trim(request.form("mid")),",")
  
        for i=0 to Ubound(delid)

     'Reset member to default group that belonged to the deleted group
     sql="update member set groupid = 0 where groupid = "&trim(delid(i))
     conn.execute(sql)
     
     ' get the group name
         Set Rs1  = Conn.Execute("Select Name From UserGroup Where GroupID ="&trim(delid(i)))
         
     Message = whatdo & " <b>" & Rs1("Name") & "</b> at " & Request.ServerVariables("REMOTE_ADDR") 
 
        Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&" ") 

     sql="delete from UserGroup where groupid="&trim(delid(i))
     
     conn.execute(sql)
     
       next

  
	  'refresh group table
	  SQL = "exec Process_DistinctTable"
	  Conn.Execute(SQL)

   whatgo="sa_group.asp?sid=" & SessionID

'==========================================================================
'
'
'  Unlock User Account
'
'
'==========================================================================


 case "unlocked account"
 
  delid=split(trim(request.form("mid")),",")
  username = trim(request.form("username"))
  MemberID = Session("ID")
  
   for i=0 to Ubound(delid)
    ' sql="Update Member Set Lock = 0 where MemberID="&trim(delid(i))
     'conn.execute(sql)
	 'response.write sql&"<br>"
   next
   

  Message = whatdo & " <b>" & username & "</b> from " & Request.ServerVariables("REMOTE_ADDR") 
 
  Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 

   whatgo="unlock.asp"


 
case ""
  conn.close
  set conn=nothing
  response.redirect "user.asp"
End select

conn.close
set conn=nothing
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