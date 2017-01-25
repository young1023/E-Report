<% Response.Buffer = False %>
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

response.expires=0


dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")

Title = "Asia Mile File Conversion"

DepotID = trim(Request("DepotID"))


' Retrieve Folder
'****************

   SQL1 = "select * from ReconDepotFolder where depotid="&DepotID
   Set Rs1 = Conn.Execute(SQL1)

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




<table border=0 cellpadding=3 cellspacing=0 width="90%" class=Normal height="100">

  <tr> 
    <td align="center">


<%
      

       ' Retrieve folder information from database
       sFolder = Trim(Server.MapPath(Rs1("DepotFolder")))


            set fo=fs.GetFolder(sFolder)

            for each x in fo.files  

' *********************************************
'         
' Remove Semicolon and strange characters
'
' *********************************************

    Set objFile = fs.OpenTextFile(sFolder&"\"&x.Name, 1)

   'Initiate line number
    lineNo = 0

    Do Until objFile.AtEndOfStream

    ' read each line of the file
    strLine = objFile.ReadLine

    strLine = replace(strLine,"""","")

    ' remove unwant character from line
    strLine = replace(strLine,"'","")

    lineNo = lineNo + 1

    ' The first line is not retrieved
    If lineNo > 1 Then
 
    ' set intLength to be the length of the line
    intLength = Len(strLine)

    'Initiate delimiter number
     delimiterNo = 0


    ' For each character of the line
    For i = 1 to intLength

        ' Read every single character
        strCharacter = Mid(strLine, i, 1)          

        
        If strCharacter = ";" Then
            
          delimiterNo = delimiterNo + 1

        End If

        If i = 1 Then

         strNewContents =  strNewContents & "AC"

        End if

        'Retrieve Family Name
        If delimiterNo = 5 Then
      
          strNewContents = strNewContents & strCharacter 

          'strFamilyName  = strFamilyName & strCharacter 

        End if

        'Retrieve Activity Date
        If delimiterNo = 68 Then
      
          strNewContents = strNewContents & strCharacter

           'strActivityDate = strActivityDate & strCharacter

        End if
    
    Next

    strNewContents =  replace(strNewContents,";","")  & space(75) & "." & vbCrLf 

    ' reset character
    strNewCharacters = ""

    


    ' end of line number > 1
    End if
  

Loop

   ' Formation of output file

   'Retrieve Month
   TapMonth = month(Now)
   If len(TapMonth) = 1 Then
      TapMonth = "0" & TapMonth
   End if

   'Retrieve Day
   TapDay = day(Now)
   If len(TapDay) = 1 Then
      TapDay = "0" & TapDay
   End if

   ' Retrieve Tape Creation Date
   TapeCreationDate = year(Now) & TapMonth & TapDay

   ' First Line
   strFirstLine = "HDNONAIR    "  & TapeCreationDate & "XX1 " & space(175) & "." & vbCrLF 

   strNewContents1 = strFirstLine & strNewContents

       
   objFile.Close

   Set objFile = fs.OpenTextFile(sFolder&"\"&x.Name, 2)
   objFile.Write strNewContents1
   objFile.Close


         ' Record if there is error
         If Err.Number <> 0 Then
  
         'Audit Log
         Conn.Execute "Exec AddReconLog 'convert error " & Err.Description & " on file " & x.Name & " ','" & Session("MemberID") & "'"
         
         End If


' **************************************************
'
' Copy csv file into required file name and format
'
' **************************************************

       fs.CopyFile sFolder&"\"&x.Name, sFolder&"\001.txt"
  

     next


       ' Get current url
        curPageURL = "http://" & Request.ServerVariables("SERVER_NAME") & "/intranet/recon/001.txt" 

    
       
%>


<a href="<% = curPageURL %>">Download File Here</a> 

<%



 set fs=nothing
 Rs1.Close
 set Rs1 = Nothing
 Conn.Close
 Set Conn = Nothing
    
       
%>



</td>
</tr>
</table>
            
    
</div>
       
</body>
</html>
