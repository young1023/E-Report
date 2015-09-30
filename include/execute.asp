<!--#include file="inc/conn.inc"-->
<%
  flage=trim(request.form("whatdo"))
  'Response.Write flage

'======================
'
' Add Customer
'
'======================  
  
  If flage = "add_cmo" Then
  
  Email=replace(trim(request.form("email")),"'","''")
  cname=replace(trim(request.form("cname")),"'","''")
  Title=replace(trim(request.form("title")),"'","''")
  phone=replace(trim(request.form("phone")),"'","''")
  Lname=replace(trim(request.form("Lname")),"'","''")
    If Lname = "" then
     Lname = " "
  End If
 
  If phone = "" then
     phone = 0
  End If
  company = trim(request.form("company"))
  country = trim(request.form("country"))
  Address = trim(request.form("Address"))
  Industry = trim(request.form("Industry"))
  shen = trim(request.form("shen"))
  Product = trim(request.form("Product"))
  Province =  trim(request.form("shen"))
  Applicant = trim(request.form("Interested_Application"))
  Source =  trim(request.form("heard_from"))
  FullName = cname & lname
  Current_Time = now()


    Sql = " Insert Into Registration (Salutation, FirstName, LastName, BusinessPhone, CompanyName, Address, Country, Source, Email, Industry, Province,"
 
    Sql = Sql & "InterestedProduct, InterestedApplicant, CreateDate) Values ("&Title&", '"&cname&"', '"&Lname&"', '"&phone&"', '"&Company&"', '"&Address&"', "

    Sql = Sql & Country&", "&Source&", '"&Email&"', "&Industry&", "&shen&", "&Product&", "&Applicant&", '"&Current_Time&"')"


    conn.execute(sql)


    Sql1 = "Insert into Customers (Name, Address,  Email, Source, Group_ID) Values "
   
    Sql1 = Sql1 & "('"&FullName&"', '"
    
    Sql1 = Sql1 &Address&"', '"&Email&"', '"

    Sql1 = Sql1 & "Registration', 21)"


    Conn.Execute(Sql1)
    
    message="Thank you for your information."
     whatgo="registration.asp?lan=2"

'======================
'
' Edit Customer
'
'======================  
  
  ElseIf flage = "edit_customer" Then
  
  ID = Request.Form("ID")
  Email=replace(trim(request.form("email")),"'","''")
  cname=replace(trim(request.form("cname")),"'","''")
  Title=replace(trim(request.form("title")),"'","''")
  phone=replace(trim(request.form("phone")),"'","''")
  Lname=replace(trim(request.form("Lname")),"'","''")
    If Lname = "" then
     Lname = " "
    End If
 
  If phone = "" then
     phone = 0
  End If
  company = trim(request.form("company"))
  country = trim(request.form("country"))
  Address = trim(request.form("Address"))
  Industry = trim(request.form("Industry"))
  shen = trim(request.form("shen"))
  Product = trim(request.form("Product"))
  Province =  trim(request.form("shen"))
  Applicant = trim(request.form("Interested_Application"))
  Source =  trim(request.form("heard_from"))
 


    Sql = "Update Registration Set Salutation = "&Title&", FirstName = '"&cname&"', "
    Sql = Sql & "Last_Name = '"&Lname&"', CompanyName = '"&company&"', "
    Sql = Sql & "BusinessPhone = '"&phone&"', Country = "&country&", email = '"&email&"', "
    Sql = Sql & "InterestedProduct = "&Product&", InterestedApplicant = "&Applicant&", "
    Sql = Sql & "Address = '"&address&"', Industry = "&Industry&", Source = "&Source&", "
    Sql = Sql & " Province = "&Province&" Where RegistrationID = "&ID
     conn.execute(sql)
    
    message="The record was updated"
     whatgo="updata.asp?chkid="&ID

'=========================
'
' Add Push Mail List
'
'=========================  

 ElseIf flage = "AddPushMailRecord" Then
  
  Email=replace(trim(request.form("email")),"'","''")
  name=replace(trim(request.form("name")),"'","''")
  Title=replace(trim(request.form("title")),"'","''")
  phone=replace(trim(request.form("phone")),"'","''")
  name=replace(trim(request.form("name")),"'","''")
  If phone = "" then
     phone = 0
  End If
  company = trim(request.form("company"))
  country = trim(request.form("country"))
  Address = trim(request.form("Address"))
  shen = trim(request.form("shen"))

  city =  trim(request.form("city"))
  Applicant = trim(request.form("Interested_Application"))
  Source_ID =  trim(request.form("selgroup"))


    Sql1 = "Insert into Customers (Name, Company, Tel, Address, Title, Email, Group_ID) Values "
   
    Sql1 = Sql1 & "('"&name&"', '"&Company&"','"&phone&"','"
    
    Sql1 = Sql1 &Address&"', '"&Title&"', '"&Email&"', "

    Sql1 = Sql1 & Source_ID&")"

    'response.write sql1
    'response.end
    conn.execute(sql1)
    
    message="The record was Added"
     whatgo="sa_ExistCustomer.asp"

     
'=========================
'
' Edit Push Mail List
'
'=========================  
  
  ElseIf flage = "Edit_ExistCustomer" Then
  
  ID = Request.Form("ID")
  Email=replace(trim(request.form("email")),"'","''")
  name=replace(trim(request.form("name")),"'","''")
  Title=replace(trim(request.form("title")),"'","''")
  phone=replace(trim(request.form("phone")),"'","''")
  name=replace(trim(request.form("name")),"'","''")
 
  If phone = "" then
     phone = 0
  End If
  company = trim(request.form("company"))
  country = trim(request.form("country"))
  Address = trim(request.form("Address"))
  Fax = trim(request.form("Fax"))
  City = trim(request.form("City"))
  system = trim(request.form("system"))
  city =  trim(request.form("city"))
  Applicant = trim(request.form("Interested_Application"))
  Source =  trim(request.form("Source"))
  SubScribe = trim(request.form("SubScribe"))
  Source_ID =  trim(request.form("selgroup"))


    Sql = "Update Customers Set Title = '"&Title&"', Name = '"&name&"', "
    Sql = Sql & "Company = '"&company&"', "
    Sql = Sql & "Tel = '"&phone&"', Fax = '"&Fax&"', Country = '"&country&"', email = '"&email&"', "
    Sql = Sql & "SubScribe="&SubScribe&", "
    Sql = Sql & "Address = '"&address&"', Group_ID = "&Source_ID&", BounceType = '"&BounceType&"',"
    Sql = Sql & " City = '"&City&"' Where CustomerID = "&ID
    'response.write sql
    'response.end
    conn.execute(sql)
    
    message="The record was updated"
     whatgo="ExistCustomer.asp?id="&ID

     
'======================
'
' Add Staff
'
'======================  

 ElseIf flage = "add_member" Then
  
  Email=replace(trim(request.form("email")),"'","''")
  firstname=replace(trim(request.form("firstname")),"'","''")
  Password = replace(trim(request.form("Password")),"'","''")
  lastname=replace(trim(request.form("lastname")),"'","''")
  phone=replace(trim(request.form("phone")),"'","''")
  LoginID = replace(trim(request.form("LoginID")),"'","''")
    If Email = "" then
     Email = " "
  End If
 
  If phone = "" then
     phone = " "
  End If

    sql = "insert into Staff (FirstName, LastName,StaffID, LogonPassword, Tel1, Email) "
    sql = sql &  "values ('"&firstname&"', '"&lastname&"', '"&LOGINID&"','"&PASSWORD&"', "
    sql = sql &  "'"&PHONE&"', '"&EMAIL&"')"
   ' response.write sql
   ' response.end
    conn.execute(sql)
    message="Thank you for your information"
  whatgo="sa_sales.asp"



'=======================================================================
'
'         Modify Staff
'
'=======================================================================  

 ElseIf flage = "ModifyStaff" Then

  ID = replace(trim(request.form("ID")),"'","''")
  Email=replace(trim(request.form("email")),"'","''")
  firstname=replace(trim(request.form("firstname")),"'","''")
  Password = replace(trim(request.form("Password")),"'","''")
  lastname=replace(trim(request.form("lastname")),"'","''")
  phone=replace(trim(request.form("phone")),"'","''")
  LoginID = replace(trim(request.form("LoginID")),"'","''")
    If Email = "" then
     Email = " "
  End If
 
  If phone = "" then
     phone = " "
  End If
  City = replace(trim(request.form("City")),"'","''")


    sql = "Update Staff Set FirstName = '"&firstname&"', LastName = '"&lastname&"', StaffID = '"&LoginID&"', LogonPassword = '"&Password&"', "
    sql = sql &  " Tel1 = '"&PHONE&"', Email = '"&EMAIL&"', RegionCode = '"&City&"' Where ID =  "&ID
    'response.write sql
    'response.end
    conn.execute(sql)
    message="Thank you for your information"
  whatgo="sa_sales.asp"


'=======================================================================
'
'         Modify Member
'
'=======================================================================  

 ElseIf flage = "modify" Then
  
  ID = replace(trim(request.form("ID")),"'","''")
  Email=replace(trim(request.form("email")),"'","''")
  name=replace(trim(request.form("username")),"'","''")
  Password = replace(trim(request.form("Password")),"'","''")
  phone=replace(trim(request.form("phone")),"'","''")
  LoginID = replace(trim(request.form("LoginID")),"'","''")
    If Email = "" then
     Email = " "
  End If
 
  If phone = "" then
     phone = " "
  End If

    sql = "Update MEMBER Set Name = '"&name&"', INDICATE = '"&LoginID&"', PWD = '"&Password&"', "
    sql = sql &  " Phone = '"&PHONE&"',  Email = '"&EMAIL&"' Where ID =  "&ID
    response.write sql
    'response.end
    conn.execute(sql)
    message="Thank you for your information"
  whatgo="sa_sales.asp"

'=======================================================================
'
'         Follow UP
'
'=======================================================================  

 ElseIf flage = "followup" Then
  
  Comment = replace(trim(request.form("comment")),"'","''")
  id = replace(trim(request.form("id")),"'","''")
  Informed_Party = replace(trim(request.form("Informed_Party")),"'","''")
  If Comment = "" then
     Comment = "No Comment"
  End If
 
 '-----------  send email To the action party ---------------

  sql="select name,email from member where indicate='"&Informed_Party&"'"
  set rs=conn.execute(sql)


    strbody="Dear sirs,"& VbCrLf & VbCrLf & _ 
   "There is a prospect registered in our website at http://www.hrstec.com.hk/irf_testing/irfsite/registration.asp?lan=2&id="&id& VbCrLf & _ 
      " Comment: "&  VbCrLf & _ 
   Comment &VbCrLf &  VbCrLf & _ 
   "Remark: This is an email alert generated by the system automatically. Please don't reply to the sendor."& VbCrLf &  VbCrLf & _
   "IRF"
    strsubject="Prospect Followup."
    '\\\\\\\\\\\ send email \\\\\\\\\\\\\\\\
    Set Msg = Server.CreateObject("JMail.Message")
    rem change this RemoteHost to a valid SMTP address before testing
    Msg.Logging = true
    Msg.silent = true
    Msg.From = "gary@hrstec.com.hk"
    Msg.FromName = "IRF"
    Msg.AddRecipient rs("email")
    Msg.Subject = strsubject
    Msg.Body = strbody

    'if not Msg.send("smtp.hrstec.com.hk") then  smtp2.artrend-tec.com port 8825
	
	'if not Msg.send("smtp2.artrend-tec.com port 8825") then
     if not Msg.send("smtp.hrstec.com.hk") then
      error = 5
      Response.write "<pre>" & Msg.log & "</pre>"
    else
      message = "An email was sent to the informed party."
    end if

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


  '---------------------------------------

'======================
'
' Add Prospect
'
'======================  

 ElseIf flage = "add_prospect" Then
  
  Email=replace(trim(request.form("email")),"'","''")
  firstname=replace(trim(request.form("firstname")),"'","''")
  cfirstname=replace(trim(request.form("cfirstname")),"'","''")
  lastname=replace(trim(request.form("lastname")),"'","''")
  clastname=replace(trim(request.form("clastname")),"'","''")

  address = replace(trim(request.form("address")),"'","''")
  caddress = replace(trim(request.form("caddress")),"'","''")
  City = replace(trim(request.form("City")),"'","''")
  phone=replace(trim(request.form("phone")),"'","''")
  Fax = replace(trim(request.form("fax")),"'","''")
  Title = replace(trim(request.form("Title")),"'","''")
  company = replace(trim(request.form("company")),"'","''")
  Ccompany = replace(trim(request.form("Ccompany")),"'","''")
  Country = replace(trim(request.form("Country")),"'","''")
  Source = replace(trim(request.form("Source")),"'","''")
    If Email = "" then
     Email = " "
      End If
 
  If phone = "" then
     phone = " "
  End If
  
  CurrentDate = DateValue(now())
  
  ' Check if the prospect is existing in the present database
  ' *******************************************************
  
  sql1 = " Select Fax,Email From Registration Where Email = '"&Email&"'"
  Set Rs1 = Conn.Execute(sql1)
  
  If Not Rs1.EoF Then
  flag = "Existing"
  Else 
  flag = "New"
  End if
  


    sql = "insert into Prospect (FirstName, CFirstName, LastName, CLastName, "
    sql  =  sql  &    " Company, CCompany, Address, Caddress, Tel, Fax, Email, CreateDate, Flag, Country, Title, City, Source) "
    sql = sql &  "values ('"&firstname&"','"&cfirstname&"','"&lastname&"', '"&clastname&"',"
    sql = sql &  "'"&Company&"', '"&CCompany&"', '"&Address&"', '"&CAddress&"', "
    sql = sql &  "'"&PHONE&"', '"&Fax&"', '"&EMAIL&"', '"&CurrentDate&"', '"&Flag&"', '"&Country&"', '"&Title&"', "&City&","&Source&")"

'Response.write sql
'Response.end

    conn.execute(sql)

    message="The record was added"
      whatgo="Lead.asp"

'=======================================================================
'
'         Modify Prospect
'
'=======================================================================  

 ElseIf flage = "EditLead" Then
 
  ID = Request("ID") 
  Email=replace(trim(request.form("email")),"'","''")
  pageid = trim(Request.form("pageid"))
  k_record = trim(request.form("k_record"))
  If k_record = "" Then
     k_record = 2
  End If
 
  Title = request.form("Title")
  firstname=replace(trim(request.form("firstname")),"'","''")
  cfirstname=replace(trim(request.form("cfirstname")),"'","''")
  lastname=replace(trim(request.form("lastname")),"'","''")
  clastname=replace(trim(request.form("clastname")),"'","''")

  address = replace(trim(request.form("address")),"'","''")
  caddress = replace(trim(request.form("caddress")),"'","''")

  phone=replace(trim(request.form("phone")),"'","''")
  Fax = replace(trim(request.form("fax")),"'","''")

  company = replace(trim(request.form("company")),"'","''")
  ccompany = replace(trim(request.form("ccompany")),"'","''")
  
  City = request.form("city")
  If City = "" Then
  City = 0
  End If
  
  Source = Request.Form("Source")
  If Source = "" Then
  Source = 0
  End If

    sql = "Update Prospect Set FirstName = '"&firstname&"', lastname = '"&Lastname&"', Cfirstname = '"&cFirstName&"', "
    sql = sql &  " CLastName = '"&cLastName&"',Company = '"&Company&"',cCompany = '"&cCompany&"'," 
    sql = sql &  " Address = '"&Address&"',CAddress = '"&CAddress&"',Fax = '"&Fax&"'," 
    sql = sql &  " Tel = '"&PHONE&"',  Email = '"&EMAIL&"', City = "&City&", Title='"&Title&"', Status="&k_record&", Source = "&Source&" where ProspectID =  "&ID

    conn.execute(sql)
    message="The record was edited."
      whatgo="sa_lead.asp?page_id="&pageid
       wait = 1

'======================
'
' Add Footer
'
'======================  

 ElseIf flage = "AddFooter" Then
  
  FooterName = replace(trim(request.form("FooterName")),"'","''")
  FooterContent = replace(trim(request.form("FooterContent")),"'","''")
 

    sql = "insert into EmailFooter (FooterName, FooterContent) "
    sql = sql &  "values ('"&FooterName&"', '"&FooterContent&"')"
   ' response.write sql
   ' response.end
    conn.execute(sql)
    message="Thank you for your information"
  whatgo="Footer.asp"

'======================
'
' Add Footer
'
'======================  

 ElseIf flage = "AddSource" Then
  
  SourceName = replace(trim(request.form("SourceName")),"'","''")
 

    sql = "insert into SourceOfCustomers (Source) "
    sql = sql &  "values ('"&SourceName&"')"
   ' response.write sql
   ' response.end
    conn.execute(sql)
    message="Thank you for your information"
  whatgo="lead.asp"
  
'======================
'
' Edit Footer
'
'======================  

 ElseIf flage = "EditFooter" Then
  
  FooterName = replace(trim(request.form("FooterName")),"'","''")
  FooterContent = replace(trim(request.form("FooterContent")),"'","''")
   FooterID = Request("ID")

    sql = "Update EmailFooter Set FooterName = '"&FooterName&"', FooterContent = '"&FooterContent&"' where FooterID="&FooterID
   ' response.write sql
   ' response.end
    conn.execute(sql)
    message="Thank you for your information"
      whatgo="Footer.asp?id="&FooterID
      
 
  
conn.close
set conn=nothing
wait=2

End If

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Refresh" content="2; url='<%=whatgo%>'">
<title>Shell execute file</title>
</head>
<body topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" >
<!--#include file ="inc/basic_upper.inc"-->
<table border=0 cellpadding=3 cellspacing=0 class=hardcolor width="90%" align=center>
  <tbody>
  <tr> 
    <td align=center height="68"></td>
  </tr>
  <tr> 
    <td align=center height="28"><%=message%></td>
  </tr>
  <tr>
    <td align=center height="38"><a href="<%=whatgo%>">Return</a> </td>
  </tr>
  </tbody>
</table>
<!--#include virtual ="inc/basic_lower.inc"-->

</body>
</html>