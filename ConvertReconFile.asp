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

     Do While Not Rs2.EoF

     FieldName =  Rs2("fieldname") & "," & FieldName

     Rs2.MoveNext

     Loop 

     FieldName =  Left(FieldName,Len(FieldName)-1) 
   

     sqv = "create view vw_"&DepotID&" as select top 1 depotid, ImportFileName, "&FieldName&" from StockReconciliation"

     response.write sqv
Conn.Execute(sqv)

     'sqd_v = "delete from view vw_"&DepotID

     'Conn.Execute(sqd_v)


       ' Set Error situation
            Err.Clear
            On Error Resume Next
  
    'response.end

    Sql3 = "bulk insert vw_"&Rs1("DepotID")

    Sql3 = Sql3 & " from '"&sFolder&"\"&x.Name&"'"

    Sql3 = Sql3 & " WITH (FIRSTROW = "& Rs1("FirstRow") &", "

    Sql3 = Sql3 & " ERRORFILE = '" & sFolder & "\Errorlog.log' , "

    Sql3 = Sql3 & " MAXERRORS = 1000 , "

    Sql3 = Sql3 & " FIELDTERMINATOR= ',',"

    Sql3 = Sql3 & " ROWTERMINATOR = '\n')"

    response.write Sql3 & "<br/>"

    Conn.Execute(Sql3)

     If Err.Number <> 0 Then
  
  
         'Audit Log
         Conn.Execute "Exec AddReconLog 'convert error " & Err.Description & " on file " & x.Name & " ','" & Session("MemberID") & "'"
         
             End If

     On Error GoTo 0
   

         ' Check if distinction file exists
         If fs.FileExists("E:\Data\Recon\Archive\"&x.Name)  Then

              fs.DeleteFile("E:\Data\Recon\Archive\"&x.Name)
 
         end if


         fs.movefile sFolder&"\"&x.Name , "E:\Data\Recon\Archive\"



     next



%>


<%


  set fs=nothing

  response.redirect "ReconDepotFile.asp?sid="&sessionid


  

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
