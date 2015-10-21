
<!--#include file="include/SQLConn.abc.asp" -->
<HTML>
<HEAD>
<TITLE></TITLE>
</HEAD>
<BODY>
<%

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSourceFile = objFSO.OpenTextFile(Server.MapPath("\Intranet\Recon\Commwealth.txt"),1,True)
Set objTargetFile = objFSO.CreateTextFile(Server.MapPath("\Intranet\Recon\Commwealth.csv"),True)

  k = 0
 
Do While Not objSourceFile.AtEndOfStream



	strData = objSourceFile.ReadLine

    intLength = Len(strData)

    response.write intLength & "<br>"

    strData = replace(strData,"'","")

		
	strName    = Trim(Mid(strData, 1, 28))
	strAddress = Trim(Mid(strData, 29, 37))
	strCity    = Trim(Mid(strData, 38, 79))


	objTargetFile.WriteLine(strName & " ," & strAddress & "," & strCity)


Loop

'response.write Server.MapPath("\Intranet\Recon\Commwealth.csv")

objSourceFile.Close
objTargetFile.Close

Set objFinalFile = objFSO.OpenTextFile(Server.MapPath("\Intranet\Recon\Commwealth.txt"),1,True)



Do While Not objFinalFile.AtEndOfStream



	strData1 = objFinalFile.ReadLine

    'rESPONSE.WRITE strData1 & "<br/>"
	

Loop


'i = False

'Do Until i = True 

    'intLength = Len(strContents)

    'If intLength < 28 Then
       ' Exit Do
   ' End If
   ' strLines = strLines & Left(strContents, 28) & "<br/>"
   ' strContents = Right(strContents, intLength - 28)



'Loop


%>
</BODY>
</HTML>
