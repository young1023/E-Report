<% Response.Buffer = False %>
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

response.expires=0


dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")

Title = "Depot File Conversion"

DepotID = trim(Request("DepotID"))


' Retrieve Folder
'****************

   SQL1 = "select * from ReconDepotFolder where depotid="&DepotID
   Set Rs1 = Conn.Execute(SQL1)

FileType = Rs1("FileType")
DepotCode = Rs1("DepotCode")



%>   

<html>
<head>
<title>UOB Intranet</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />

<SCRIPT language=JavaScript>
<!--

function listToAray(fullString, separator) {
  var fullArray = [];

  if (fullString !== undefined) {
    if (fullString.indexOf(separator) == -1) {
      fullAray.push(fullString);
    } else {
      fullArray = fullString.split(separator);
    }
  }

  return fullArray;
}
//-->
</SCRIPT>
</head>

<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">


<div align="center">


<table border=0 cellpadding=3 cellspacing=0 width="90%" class=Normal height="100">

  <tr> 
    <td align="center" height="50">


<%
             
       ' Get the folder
       sFolder = Trim(Server.MapPath(Rs1("DepotFolder")))

      set fo=fs.GetFolder(sFolder)

            for each x in fo.files  
  
      'Audit Log
      Conn.Execute "Exec AddReconLog 'converted file " & x.Name & "','" & Session("MemberID") & "'"
      
         


Set objFile = fs.OpenTextFile(sFolder&"\"&x.Name, 1)

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    strLine = replace(strLine,"'","-")
 
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
            End If
        Else
            strNewCharacters = strNewCharacters & strCharacter
        End If
    Next

    strNewContents = strNewContents & strNewCharacters & vbCrLf
    strNewCharacters = ""
    strNewContents = replace(strNewContents,"""","") 

Loop
      
objFile.Close

Set objFile = fs.OpenTextFile(sFolder&"\"&x.Name, 2)
objFile.Write strNewContents
objFile.Close

           
            
            ' Delete imported record if exists, delete view if exists
             Conn.Execute "Exec ConvertReconFile '" & DepotID & "', '" & x.Name & "'"



 
Dim strData
 
Dim objFSO, objSourceFile, objTargetFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
Set objSourceFile = objFSO.OpenTextFile(sFolder&"\"&x.Name, 1)
 
Do While Not objSourceFile.AtEndOfStream
	strData = objSourceFile.ReadLine

    intLength = Len(strData)

    response.write intLength & "<br>"
    response.write strData
    
    
    If intLength = 75 and DepotCode = 605 then
	
	strName2 = Trim(Mid(strData, 1, 29))
	strName3 = replace(replace(Trim(Mid(strData, 65, 15)),",",""),".","")

    sql_c1 = "Select * from InstrumentMapTable where InstrumentName = '"&Trim(strName2)&"'"
    Set rs_c1 = Conn.Execute(sql_c1)

     If rs_c1.EoF Then

     sql_c2 = "Insert into InstrumentMapTable (InstrumentName) Values ('"&Trim(strName2)&"')"  

     Conn.execute(sql_c2)
     
     Else
     
      If rs_c1("InstrumentCode") <> "" then
      
         InstrumentCode = rs_c1("InstrumentCode")
         
      Else
      
         InstrumentCode = "NA"
         
      End if
     
     sql_i1 = "Insert into StockReconciliation (DepotID, ImportFileName, Instrument, UnitHeld) Values (" & DepotID & ", '" & x.Name & "' , '" & InstrumentCode & "' , '" & strName3 &"')"

     
     End If
 
     ElseIf intLength = 80  and DepotCode = 202 then


    strName1 = Trim(Mid(strData, 1, 9))
	strName2 = Trim(Mid(strData, 37, 15))
	strName3 = replace(replace(Trim(Mid(strData, 50, 15)),",",""),".","")
	
	
	
   sql_i1 = "Insert into StockReconciliation (DepotID, ImportFileName, ISINCode, UnitHeld) Values (" & DepotID & ", '" & x.Name & "' , '" & strName2 & "' , '" & strName3 &"')"


    End If
    
  	
  
    response.write sql_i1 & "<br>"
    Conn.Execute(sql_i1)
  

Loop

     next
     
  'response.end
	
  set fs=nothing

  response.redirect "ConvertTxtFile2.asp?depotid="&depotid&"&sid="&sessionid

 Rs1.Close
 set Rs1 = Nothing
 Conn.Close
 Set Conn = Nothing
    
  

%>


</td>

</table>
            
</div>        
</div>
<%
Conn.Close
Set Conn = Nothing
%>          
</body>
</html>
