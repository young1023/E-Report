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
<!DOCTYPE html>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<TITLE>Upload File</TITLE>
<script type="text/javascript">
    function CloseWindow() {
        
        window.opener.location.reload();
        window.close();
    }
</script>
</head>
<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">


      <table width="600" height="400" border="0" cellspacing="0" cellpadding="2" class="Normal">
        <tr>
          <td align="middle">

Upload file to <% =SFolder %>
          </td>

        </tr>
        <tr>
          <td align="middle">

<%

'Response.Write Trim(Rs1("FileType"))

Dim objUpload

If Request("action")="1" Then

    Set objUpload=New ShadowUpload

    If objUpload.GetError<>"" Then

        Response.Write("sorry, could not upload: "&objUpload.GetError)

    Else  

        Response.Write("found "&objUpload.FileCount&" files...<br /><br/>")

        For x=0 To objUpload.FileCount-1

            Response.Write("file name: "&objUpload.File(x).FileName&"<br /><br />")
            Response.Write("file size: "&objUpload.File(x).Size&"<br /><br />")

            
        
            If (objUpload.File(x).ImageWidth>200) Or (objUpload.File(x).ImageHeight>200) Then

                Response.Write("image to big, not saving!")

            Elseif Trim(Rs1("FileType")) <> Trim(Right(objUpload.File(x).FileName, 3)) Then
              

               Response.Write("<font color=red><b>Wrong file Type!</b></font>")

            Else  

                Call objUpload.File(x).SaveToDisk(Server.MapPath(sFolder) , "")

   
            End If

        Next


    End If
End If


Set objUpload = Nothing

response.redirect "ConvertReconFile.asp?depotid="&Rs1("DepotID")&"&sid="&sessionid



%>


</td>
  </tr>
 

     </table>

</div>

 </body>
    </html>