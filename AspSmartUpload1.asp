<HTML>
<BODY BGCOLOR="white">

<H1>aspSmartUpload : Sample 1</H1>
<HR>

<%
' Variables
' *********
Dim mySmartUpload
Dim intCount

' Object creation
' ***************
Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

mySmartUpload.MaxFileSize = 100000000000

' Upload
' ******
mySmartUpload.Upload

' Save the files with their original names in a virtual path of the web server
' ****************************************************************************
intCount = mySmartUpload.Save("Recon")

' Display the number of files uploaded
' ************************************
Response.Write(intCount & " file(s) uploaded.")
%>
</BODY>
</HTML>