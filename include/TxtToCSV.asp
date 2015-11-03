<%



If FileType  = "txt" then


  'Get ArchiveFolder
   
 
Dim strData
 
Dim objFSO, objSourceFile, objTargetFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
Set objSourceFile = objFSO.OpenTextFile(sFolder&"\"&x.Name, 1)
 
Do While Not objSourceFile.AtEndOfStream
	strData = objSourceFile.ReadLine

    intLength = Len(strData)

    response.write intLength & "<br>"
	
    If intLength = 75 then
	
	strName1 = Trim(Mid(strData, 1, 29))
	strName2 = Trim(Mid(strData, 30, 15))
	strName3 = Trim(Mid(strData, 65, 15))

    sql_c1 = "Select * from InstrumentMapTable where InstrumentName = '"&Trim(strName1)&"'"
    Set rs_c1 = Conn.Execute(sql_c1)

     If rs_c1.EoF Then

    sql_c2 = "Insert into InstrumentMapTable (InstrumentName) Values ('"&Trim(strName1)&"')"
    Conn.execute(sql_c2)

     End if

    Elseif intLength = 80 then


    strName1 = Trim(Mid(strData, 1, 29))
	strName2 = Trim(Mid(strData, 38, 15))
	strName3 = Trim(Mid(strData, 50, 15))

    End If
  	
    sql_i1 = "Insert into StockReconciliation (DepotID, ImportFileName, ISINCode, UnitHeld) Values (" & DepotID & ", '" & x.Name & "' , '" & strName2 & "' , '" & strName3 &"')"

    'response.write sql_i1
    Conn.Execute(sql_i1)


    

Loop

'response.end
End If


%>