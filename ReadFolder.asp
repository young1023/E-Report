<%
Set Conn = Server.CreateObject("ADODB.Connection")
'StrCnn = "Provider=vfpoledb;Data Source=Server.MapPath&"Recon\Archive\1015_bd500788.831.dbf;Collating Sequence=machine;"

StrCnn = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBC;SourceDB=E:\WebData\Elegant\Home\Intranet\Recon\Archive\1015_bd500788.831.dbf;Exclusive=No;NULL=NO;Collate=Machine;BACKGROUNDFETCH=NO;DELETED=NO;"


Conn.CommandTimeout=0
Conn.ConnectionTimeout=0

Conn.Open StrCnn

%>
<HTML>
<HEAD>
<TITLE></TITLE>
<META name="description" content="">
<META name="keywords" content="">
<META name="generator" content="CuteHTML">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0000FF" VLINK="#800080">
<%



Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("E:\WebData\Elegant\Home\Intranet\Recon\Archive\1015_bd500788.831",1)

strContents = objFile.ReadAll
objFile.Close

i = False

Do Until i = True 
    intLength = Len(strContents)
    If intLength < 28 Then
        Exit Do
    End If
    strLines = strLines & Left(strContents, 28) & "<br/>"
    strContents = Right(strContents, intLength - 28)



Loop

'response.write strContents & "<br/><br/>"
response.write strLines 

%>
</BODY>
</HTML>
