<%

Set objFile = fs.OpenTextFile(sFolder&"\"&x.Name, 1)

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    strLine = replace(strLine,"'","")
    strLine = replace(strLine,"/","")
   
    intLength = Len(strLine)

    blnStart = False

    For i = 1 to intLength
        strCharacter = Mid(strLine, i, 1)
        If strCharacter = Chr(34) Then
            If blnStart = True Then
                blnStart = False
            Else
                blnStart = True
            End If
        End If

        If strCharacter = "," Then
            If blnStart = "False" Then
                strNewCharacters = strNewCharacters & strCharacter
            Else 
                strnewCharacters = strNewCharacters & " "
            End If
        Else
            strNewCharacters = strNewCharacters & strCharacter
        End If
    Next

    strNewContents = strNewContents & strNewCharacters & vbCrLf
    strNewCharacters = ""
    strNewContents = replace(strNewContents,"""","")

    'response.Write(i & ". " &strNewContents & "<br/><br/>")

    

Loop
     
objFile.Close

Set objFile = fs.OpenTextFile(sFolder&"\"&x.Name, 2)
objFile.Write strNewContents
objFile.Close

%>