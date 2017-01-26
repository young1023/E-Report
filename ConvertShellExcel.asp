<% Response.Buffer = False %>
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

response.expires=0


dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")

Title = "Asia Mile File Conversion"

DepotID = trim(Request("DepotID"))


' Retrieve Folder
'****************

   SQL1 = "select * from ReconDepotFolder where depotid="&DepotID
   Set Rs1 = Conn.Execute(SQL1)

%>   

<html>
<head>
<title>UOB Intranet</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
</head>

<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">




<table border=0 cellpadding=3 cellspacing=0 width="90%" class=Normal height="100">

  <tr> 
    <td align="center">


<%
      

       ' Retrieve folder information from database
       sFolder = Trim(Server.MapPath(Rs1("DepotFolder")))


       set fo=fs.GetFolder(sFolder)


     ' Retrieve file in folder
      for each x in fo.files 

     Set objFile = fs.OpenTextFile(sFolder&"\"&x.Name, 1)

   ExcelFile = sFolder&"\"&x.Name

response.write ExcelFile


SQL = "SELECT [ISO 3166-1], [Country Name] FROM [Sheet1$]"
Set ExcelConnection = Server.createobject("ADODB.Connection")

ExcelConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ExcelFile & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"

SET RS = Server.CreateObject("ADODB.Recordset")
RS.Open SQL, ExcelConnection
Response.Write "<table border=""1""><thead><tr>"
FOR EACH Column IN RS.Fields
	Response.Write "<th>" & Column.Name & "</th>"
NEXT
Response.Write "</tr></thead><tbody>"
IF NOT RS.EOF THEN
	WHILE NOT RS.eof
		Response.Write "<tr>"
		FOR EACH Field IN RS.Fields
			Response.Write "<td>" & Field.value & "</td>"
		NEXT
		Response.Write "</tr>"
		RS.movenext
	WEND
END IF

   next

Response.Write "</tbody></table>"
RS.close
ExcelConnection.Close


 



      


' **************************************************
'
' Copy csv file into required file name and format
'
' **************************************************

       fs.CopyFile sFolder&"\"&x.Name, sFolder&"\001.txt"
  


       ' Get current url
        curPageURL = "http://" & Request.ServerVariables("SERVER_NAME") & "/intranet/recon/001.txt" 

    
       
%>


<a href="<% = curPageURL %>">Download File Here</a> 

<%



 set fs=nothing
 Rs1.Close
 set Rs1 = Nothing
 Conn.Close
 Set Conn = Nothing
    
       
%>



</td>
</tr>
</table>
            
    
</div>
       
</body>
</html>
