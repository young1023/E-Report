<% Response.Buffer = False %>
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

response.expires=0


dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")

Title = "Depot File Conversion"

DepotID = trim(Request("DepotID"))


' Retrieve Folder
'****************

   SQL1 = "select * from ReconDepotFolder where depotid="&DepotID
   Set Rs1 = Conn.Execute(SQL1)

FileType = Rs1("FileType")



%>   

<html>
<head>
<title>UOB Intranet</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />

<SCRIPT language=JavaScript>
<!--

function listToAray(fullString, separator) {
  var fullArray = [];

  if (fullString !== undefined) {
    if (fullString.indexOf(separator) == -1) {
      fullAray.push(fullString);
    } else {
      fullArray = fullString.split(separator);
    }
  }

  return fullArray;
}
//-->
</SCRIPT>
</head>

<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">


<div align="center">


<table border=0 cellpadding=3 cellspacing=0 width="90%" class=Normal height="100">

  <tr> 
    <td align="center" height="50">


<%
             
       ' Get the folder
       sFolder = Trim(Server.MapPath(Rs1("DepotFolder")))

      set fo=fs.GetFolder(sFolder)

            for each x in fo.files  
  

      
         'Get Archive Folder
         set RsFd = server.createobject("adodb.recordset")
         RsFd.open ("Exec Get_SystemSetting 'ArchiveFolder'") ,  conn,3,1


         ' Check if distinction file exists
         If fs.FileExists(RsFd("SettingValue") & x.Name)  Then

              fs.DeleteFile(RsFd("SettingValue") & x.Name)
 
         end if
           
         response.write x.Name

         fs.movefile sFolder&"\"&x.Name , RsFd("SettingValue") 
    

 

     next

  set fs=nothing

  response.redirect "ReconCheckList.asp?depotid="&depotid&"&sid="&sessionid

 Rs1.Close
 set Rs1 = Nothing
 Conn.Close
 Set Conn = Nothing
    
  

%>


</td>

</table>
            
</div>        
</div>
<%
Conn.Close
Set Conn = Nothing
%>          
</body>
</html>
