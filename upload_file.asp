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


      <table width="96%" height="400" border="0" cellspacing="0" cellpadding="2" class="Normal">
        <tr>
          <td align="middle">

Upload file to <% = Rs1("DepotName") %> - <% = Rs1("Market") %>
          </td>

        </tr>
        <tr>
          <td align="middle">

<%

Dim objUpload

If Request("action")="1" Then

    Set objUpload=New ShadowUpload

    If objUpload.GetError<>"" Then

        Response.Write("sorry, could not upload: "&objUpload.GetError)

    Else  

        Response.Write("found "&objUpload.FileCount&" files...<br /><br/>")

        For x=0 To objUpload.FileCount-1

            Response.Write("file name: "&objUpload.File(x).FileName&"<br /><br />")
            Response.Write("file type: "&objUpload.File(x).ContentType&"<br /><br />")
            Response.Write("file size: "&objUpload.File(x).Size&"<br /><br />")
           
            If (objUpload.File(x).ImageWidth>200) Or (objUpload.File(x).ImageHeight>200) Then

                Response.Write("File to big, not saving!")

            Else  

                Call objUpload.File(x).SaveToDisk(Server.MapPath(sFolder) , "")

                Response.Write("file saved successfully!")

            End If

            Response.Write("<hr />")
        Next
        'Response.Write("thank you, "&objUpload("name"))
    End If
End If


If Request("action") <> "1" Then
%>
<form action="Uploading.asp?sid=<%=sessionid%>&action=1&DepotID=<%=DepotID%>" enctype="multipart/form-data" method="POST">


File:&nbsp;&nbsp;&nbsp;<input type="file" name="file1" />

<br /><br /><br /><br />

<button type="submit">Upload</button>
</form>


<% End If %>

</td>
  </tr>
 


<tr>
<td align="middle">



</td>
   </tr>
     </table>

</div>
 </body>
    </html>