Const strSourceFile = "c:\test.txt"
Const strTargetFile = "c:\test.csv"
 
Dim strData
 
Dim objFSO, objSourceFile, objTargetFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
Set objSourceFile = objFSO.OpenTextFile(strSourceFile, 1, True)
Set objTargetFile = objFSO.CreateTextFile(strTargetFile, True)
 
Do While Not objSourceFile.AtEndOfStream
	strData = objSourceFile.ReadLine
	'msgbox strData
	
	Dim strName, strAddress, strCity, strState, strZip
	
	strName = Trim(Mid(strData, 1, 20))
	strAddress = Trim(Mid(strData, 21, 20))
	strCity = Trim(Mid(strData, 41, 20))
	strState = Trim(Mid(strData, 61, 2))
	strZip = Trim(Mid(strData, 63 5))
	
	objTargetFile.WriteLine("""" & strName & """,""" & strAddress & """,""" & strCity & """,""" & strState & """,""" & strZip & """")

Loop