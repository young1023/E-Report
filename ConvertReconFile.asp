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

DepotCode = Rs1("DepotCode")

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

      
      'Audit Log
      Conn.Execute "Exec AddReconLog 'converted file " & x.Name & "','" & Session("MemberID") & "'"
         

%>

          

            <!--#include file="include/remove_comma.inc.asp" -->

            

   
<%
           
            
            ' Delete imported record if exists, delete view if exists
             Conn.Execute "Exec ConvertReconFile '" & DepotID & "', '" & x.Name & "'"



     
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

     'response.write sql2 

     If Not Rs2.EoF Then

     Do While Not Rs2.EoF

     FieldName =  Rs2("fieldname") & "," & FieldName

     Rs2.MoveNext

     Loop 

     End If 

     FieldName =  Left(FieldName,Len(FieldName)-1) 
   
     sqv = "create view vw_"&DepotID&" as select top 1 depotid, ImportFileName, "&FieldName&" from StockReconciliation"

     'response.write sqv

    
     Conn.Execute(sqv)

    Sql3 = "bulk insert vw_"&DepotID

    Sql3 = Sql3 & " from '"&sFolder&"\"&x.Name&"'"

    Sql3 = Sql3 & " WITH (FIRSTROW = "& Rs1("FirstRow") &", "

    Sql3 = Sql3 & " ERRORFILE = '" & sFolder & "\Errorlog.log' , "

    Sql3 = Sql3 & " MAXERRORS = 1000 , "

    Sql3 = Sql3 & " FIELDTERMINATOR= ',',"

    Sql3 = Sql3 & " ROWTERMINATOR = '\n')"

    'response.write Sql3 & "<br/>"


    'response.end
    Conn.Execute(Sql3)


            ' Record if there is error
            If Err.Number <> 0 Then
  
  
         'Audit Log
         Conn.Execute "Exec AddReconLog 'convert error " & Err.Description & " on file " & x.Name & " ','" & Session("MemberID") & "'"
         
             End If

     On Error GoTo 0


       ' Replace Depot code for multi market files
        Conn.Execute "Exec Replace_ReconDepot '" & DepotCode & "', '" & x.Name & "'"



 
         'Get Archive Folder
         set RsFd = server.createobject("adodb.recordset")
         RsFd.open ("Exec Get_SystemSetting 'ArchiveFolder'") ,  conn,3,1


         ' Check if distinction file exists
         If fs.FileExists(RsFd("SettingValue") & x.Name)  Then

              fs.DeleteFile(RsFd("SettingValue") & x.Name)
 
         end if
           
         response.write x.Name

         'fs.movefile sFolder&"\"&x.Name , RsFd("SettingValue") 
    

     next





%>


<%


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
