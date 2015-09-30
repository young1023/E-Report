
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

response.expires=0
flag=trim(request.form("whatTodo"))

dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")

Title = "Depot File Conversion"


' Retrieve Folder
'****************

   SQL1 = "select * from ReconDepotFolder order by DepotName Asc"
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
      	If Not Rs1.EoF Then

             		Rs1.MoveFirst

 
			Do While Not Rs1.EoF

       ' Get the folder
      sFolder=Rs1("DepotFolder")

      If fs.FolderExists(sFolder)=true Then

    
      If fs.GetFolder(sFolder).Files.Count  > 0  Then


            set fo=fs.GetFolder(sFolder)

            for each x in fo.files  

       ' Check file extension
      If Right(x.Name,3) = Rs1("FileType") Then

 
     Sql = "Select count(f.depotid) as Tcount from  (ReconDepotFolder f join reconfileorder o "

     Sql = Sql & "on f.depotid = o.depotid) join ReconFile r on o.fieldid = r.fieldid "

     Sql = Sql & " and f.depotid=" &Rs1("DepotID")
  
     Set Rs = Conn.Execute(Sql)

     If Rs("Tcount") > 0 then

     Sql2 = "Select f.depotid, fieldname from  (ReconDepotFolder f join reconfileorder o "

     Sql2 = Sql2 & "on f.depotid = o.depotid) join ReconFile r on o.fieldid = r.fieldid "

     Sql2 = Sql2 & " and f.depotid=" &Rs1("DepotID")

     Sql2 = Sql2 & "order by f.depotid, o.priority"

     FieldName = ""

     Set Rs2 = Conn.Execute(Sql2)

      Do While Not Rs2.EoF

     FieldName =  Rs2("fieldname") & "," & FieldName

     Rs2.MoveNext

      Loop ' Rs2

     FieldName = Left(FieldName,Len(FieldName)-1) 

     Response.write FieldName

     'sqv = "create view vw_"&Rs1("DepotID")&" as select "&FieldName&" from StockReconciliation"

     'Conn.Execute(sqv)
  
   


  Response.write("Converting "&x.Name& "<br/><br/>")

  set f=fs.OpenTextFile(sFolder&"\"&x.Name,1)

  ' read line
  Do While Not f.AtEndOfStream


    strReadLineText = f.ReadLine

    response.Write(strReadLineText & "<br>")

      

    If strReadLineText<>"" then

        If Instr(strReadLineText,",")>0 then

            strReadLineTextArr=split(strReadLineText,",")

            response.Write(strReadLineTextArr(1)&"<br/>")

            'URLString=strReadLineTextArr(1)

         
        end if 

    end if

    

   Loop

      End if ' check depot field exist

      End If ' check file type

     'sqv_d = "drop view vw_"&Rs1("DepotID")&"" 

     'Conn.Execute(sqv_d)


  next

     
    

 End if    ' check file exists

End if     ' check folder exists


%>


<%
	Rs1.movenext 

	   loop 
 
	End If


set fs=nothing


%>




</td>
  <tr>
    <td align="center" height="50"><br><a href="<%=whatgo%>">Return</a></td>
  </tr>
  
</table>
            
</div>        
</div>
<%
Conn.Close
Set Conn = Nothing
%>          
</body>
</html>
