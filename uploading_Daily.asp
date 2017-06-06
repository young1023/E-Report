<!--#include file="include/SessionHandler.inc.asp" -->
<!-- #include file="ShadowUpload.asp" -->
<%

DepotID = Request("DepotId")

' Check folder
    
' *****************
      
       SQL1 = "select * from ReconDepotFolder where DepotId="&depotId

       Set Rs1 = Conn.Execute(SQL1)

            sFolder = Trim(Rs1("DepotFolder"))

dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")
%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<TITLE>Upload File</TITLE>
</head>
<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">


      <table width="98%" height="400" border="1" cellspacing="0" cellpadding="2" class="Normal">
        <tr>
          <td align="middle" height="50">

Upload file to <% = Rs1("Depotname") %>
          </td>

        </tr>
        <tr>
          <td align="middle">

<%

'Response.Write Trim(Rs1("FileType"))

Dim objUpload



    Set objUpload=New ShadowUpload

    If objUpload.GetError<>"" Then

        Response.Write("sorry, could not upload: "&objUpload.GetError)

    Else  


        For x=0 To objUpload.FileCount-1

            Response.Write("file name: "&objUpload.File(x).FileName&"<br /><br />")
            Response.Write("file size: "&objUpload.File(x).Size&"<br /><br />")

            
            


         ' Check if distinction file exists
         If fs.FileExists(Server.MapPath(sFolder) & objUpload.File(x).FileName)  Then

              fs.DeleteFile(Server.MapPath(sFolder) & objUpload.File(x).FileName)
 
         end if
    
        


                     Call objUpload.File(x).SaveToDisk(Server.MapPath(sFolder) , "")


                    
                   
               

    response.redirect "ConvertDailyFile.asp?depotid="&Rs1("DepotID")&"&sid="&sessionid


  

        Next


       End If
 

Set objUpload = Nothing


      

%>


</td>
  </tr>
 

     </table>

</div>

 </body>
    </html>