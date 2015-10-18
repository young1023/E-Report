<!-- #include file="ShadowUpload.asp" -->
<%
Dim objUpload
If Request("action")="1" Then
    Set objUpload=New ShadowUpload
    If objUpload.GetError<>"" Then
        Response.Write("sorry, could not upload: "&objUpload.GetError)
    Else  
        Response.Write("found "&objUpload.FileCount&" files...<br />")
        For x=0 To objUpload.FileCount-1
            Response.Write("file name: "&objUpload.File(x).FileName&"<br />")
            Response.Write("file type: "&objUpload.File(x).ContentType&"<br />")
            Response.Write("file size: "&objUpload.File(x).Size&"<br />")
            Response.Write("image width: "&objUpload.File(x).ImageWidth&"<br />")
            Response.Write("image height: "&objUpload.File(x).ImageHeight&"<br />")
            If (objUpload.File(x).ImageWidth>200) Or (objUpload.File(x).ImageHeight>200) Then
                Response.Write("image to big, not saving!")
            Else  
                Call objUpload.File(x).SaveToDisk(Server.MapPath("Uploads"), "")
                Response.Write("file saved successfully!")
            End If
            Response.Write("<hr />")
        Next
        Response.Write("thank you, "&objUpload("name"))
    End If
End If
%>
<form action="<%=Request.ServerVariables( "Script_Name" )%>?action=1" enctype="multipart/form-data" method="POST">
File1: <input type="file" name="file1" /><br />
File2: <input type="file" name="file2" /><br />
File3: <input type="file" name="file3" /><br />
Name: <input type="text" name="name" /><br />
<button type="submit">Upload</button>
</form>