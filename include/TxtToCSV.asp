<%



If FileType  = "txt" then

'Const strSourceFile = sFolder & x.Name
 strTargetFile = sFolder &"\"&  Left(x.Name, Len(x.Name) - 4) & ".csv"
 
Dim strData
 
Dim objFSO, objSourceFile, objTargetFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
Set objSourceFile = objFSO.OpenTextFile(sFolder&"\"&x.Name, 1)
Set objTargetFile = objFSO.CreateTextFile(strTargetFile, True)
 
Do While Not objSourceFile.AtEndOfStream
	strData = objSourceFile.ReadLine

    intLength = Len(strData)

    response.write intLength & "<br>"
	
    If intLength = 75 then
	
	strName1 = Trim(Mid(strData, 1, 29))
	strName2 = Trim(Mid(strData, 30, 15))
	strName3 = Trim(Mid(strData, 65, 15))


    Elseif intLength = 80 then


    strName1 = Trim(Mid(strData, 1, 29))
	strName2 = Trim(Mid(strData, 38, 15))
	strName3 = Trim(Mid(strData, 50, 15))

    End If
  	
    sql_i1 = "Insert into StockReconciliation (DepotID, ImportFileName, ISINCode, UnitHeld) Values (" & DepotID & ", '" & x.Name & "' , '" & strName2 & "' , '" & strName3 &"')"

    response.write sql_i1
    Conn.Execute(sql_i1)

	'objTargetFile.WriteLine(  DepotID & "," & x.Name & "," & strName1 & "," & strName2 & "," & strName3  )

    

Loop

'response.end
End If


%>