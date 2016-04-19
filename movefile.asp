
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

response.expires=0
flag=trim(request.form("whatTodo"))

dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")

' Retrieve Folder
'****************

   SQL1 = "select * from ReconDepotFolder order by DepotName Asc"
   Set Rs1 = Conn.Execute(SQL1)


%>
<HTML>
<HEAD>
<title>UOB Intranet</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
</HEAD>
<%


     	If Not Rs1.EoF Then

             		Rs1.MoveFirst

 
			Do While Not Rs1.EoF

       ' Get the folder
      sFolder = Trim(Server.MapPath(Rs1("DepotFolder")))

 sReadyToConvert = Trim(Rs1("ReadyToConvert"))




      If sReadyToConvert = "True" Then



       set fo=fs.GetFolder(sFolder)

            for each x in fo.files  

 

     If fs.FileExists(sFolder&"\"&x.Name) Then 

 
         ' Distinction file exists
         If fs.FileExists("E:\Data\Recon\Archive\"&x.Name)  Then

              fs.DeleteFile("E:\Data\Recon\Archive\"&x.Name)
 
         end if

         fs.movefile sFolder&"\"&x.Name , "E:\Data\Recon\Archive\"


           SQL5 = "Update ReconDepotFolder Set ReadyToConvert = 0 Where DepotID ="&Rs1("DepotID")

           Conn.Execute(SQL5)
              


     end if


    next

     end if


	Rs1.movenext 

	   loop 
 
	End If


set fs=nothing

   
Response.Redirect "ReconDepotFile.asp?sid="&sessionid
    

%>
</BODY>
</HTML>
