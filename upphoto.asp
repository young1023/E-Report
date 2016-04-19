<!--#include file="images/conne.inc" -->
<% 
order_id = session("order_id")
'response.write order_id
%>
<html>
<head>
<title>Shell Gas Maintenance Database</title>
<link rel="stylesheet" type="text/css" href="hse.css" />
<meta http-equiv="Refresh" content="2; url='upload_photo.asp?flag=<% =order_id %>'">
</head>

<body leftmargin="0" topmargin="0" >

     <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="10">
         <tr valign="top">
          <td height="100%" align="middle">
<%
 

'  Variables
'  *********
   Dim mySmartUpload
   Dim file
   Dim oConn
   Dim oRs
   Dim intCount
   Dim item
   intCount=0
        
'  Object creation
'  ***************
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

'  Upload
'  ******
   mySmartUpload.Upload


'  Open a recordset
'  ****************
   strSQL = "SELECT Order_id,name,path FROM Photo"
   'response.write strSQL 
   'response.end

   Set oRs = Server.CreateObject("ADODB.recordset")
   Set oRs.ActiveConnection = Conn
   oRs.Source = strSQL
   oRs.LockType = 3
   oRs.Open

'  Select each file
'  ****************
   For each file In mySmartUpload.Files
   '  Only if the file exist
   '  **********************
      If not file.IsMissing Then

      '  Add the current file in a DB field
      '  **********************************
         oRs.AddNew
		 oRs("Path") = file.FileName
     '  file.SaveAs("document" & file.FileName)
         intCount = mySmartUpload.Save("Photo")
             oRs.Update
    
  '      intCount = intCount + 1
      End If
   Next

'  Display the number of files uploaded
'  ************************************
   Response.Write (intCount & " file(s) uploaded.<BR>")

'  Select each item
'  ****************
   dim postform(2)
   i=0
   For each item In mySmartUpload.Form
   '  Select each value of the current item
   '  *************************************
      For each value In mySmartUpload.Form(item)
         postform(i)=value
'         Response.Write(item & " = " & value & "<BR>")
         i=i+1
      Next
   Next


'  insert form information
'  ***********************

   sql = " select max(id) as id from Photo"
   set rs = conn.execute(sql)

   id = rs("id")
   
   usql = " Update Photo set Order_ID = "&postform(0)&" where id ="&id
 '  response.write usql
  ' response.end
   conn.execute(usql)


   
   
'  Destruction
'  ***********
   oRs.Close
   Conn.Close
   Set oRs = Nothing 
   Set Conn = Nothing 

'-----------------------------------------------------------------------------
'
'      End of the main Content 
'
'-----------------------------------------------------------------------------
%>
</td>
              </tr>
                </table>
              </body>
              </html>