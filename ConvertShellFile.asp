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


%>

 <!--#include file="include/remove_Shell_comma.inc" -->

<%

    

    strNewContents = ""
    strCharacter   = ""

    Set objFile = fs.OpenTextFile(sFolder&"\"&x.Name, 1)

   'Initiate line number
    lineNo = 0

    Do Until objFile.AtEndOfStream

    ' read each line of the file
    strLine = objFile.ReadLine

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

        
        If strCharacter = Chr(44) Then
            
          delimiterNo = delimiterNo + 1

        End If


        'Retrieve Given Name
        If delimiterNo = 4 Then
      
         strGivenName  = replace(strGivenName,",","") & strCharacter 

        End if


        'Retrieve Family Name
        If delimiterNo = 5 Then
      
         strFamilyName  = replace(replace(strFamilyName,",","")," ","_") & strCharacter 

        End if

        'Retrieve mile
        If delimiterNo = 14 Then
      
         strMile  = replace(strMile,",","") & strCharacter 

        End if

        'Retrieve Activity Date
        If delimiterNo = 17 Then
      
          'strNewContents = strNewContents & strCharacter

          strActivityDate = replace(strActivityDate,",","") & strCharacter

        End if

       ' Retrieve Membership number
       If delimiterNo = 18 Then
      
          'strNewContents = strNewContents & strCharacter

           strMembership = replace(strMembership,",","") & strCharacter

        End if

    
    Next


     SQL2 = "Insert into AsiaMileData (LineNumber, Membership, FamilyName,  ActivityDate, Miles) Values "

     SQL2 = SQL2 & "( '" & LineNo &"' , '" & strMembership &"' , '" & strFamilyName &"' ,'" & strActivityDate & "', "

     SQL2 = SQL2 & " ' " & strMile & "' )"

     'Response.write "Write into database :" & SQL2 & "<br/>"

     Conn.Execute(SQL2)


    ' reset character
    strNewCharacters  = ""
    strMembership     = ""
    strFamilyName     = ""
    strGivenName     = ""
    strActivityDate   = ""
    strMile           = ""

    


    ' end of line number > 1
    End if
  

Loop


   ' ****************************************************
   '
   ' Formation of output file
   '
   ' ****************************************************

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

   ' **********************************************************
   '
   ' Formation of First Line
   '
   ' **********************************************************

   strFirstLine = "HDNONAIR    "  & TapeCreationDate & "XX1 " & space(175) & "." & vbCrLF 

   ' **********************************************************
   ' 
   '  Formation of content
   '
   ' ***********************************************************

   ' Retrieve System parameters
   SQL3 = "Select * from AsiaMileSetup"

   Set Rs3 = Conn.Execute(SQL3)

   ' ***********************************************************
   '
   ' Retrieve content
   '
   ' ***********************************************************

   SQL4 = "Select * from AsiaMileData order by LineNumber asc"
   
   Set Rs4 = Conn.Execute(SQL4)

   If Not Rs4.EoF Then

       Rs4.MoveFirst

     Do While Not Rs4.EoF

       
       ' ********************************************************
       '
       ' Membership
       '
       ' ********************************************************
       strMembership = Rs4("Membership")

       iSpace = 10 - Len(strMembership)

       strMembership = strMembership & space(iSpace) 

      ' **********************************************************
      '
      ' handle space of Family Name
      '
      ' **********************************************************
       strFamilyName = Rs4("FamilyName")

       iSpace = 25 - Len(strFamilyname)

       strFamilyName =  strFamilyName & space(iSpace) & "GivenName" & space(41)

       ' ***********************************************************
       '
       '  Handle Activity Date
       '
       ' ***********************************************************
       strActivityDate = Rs4("ActivityDate")

       strActivityDate = Right(strActivityDate,4) & mid(strActivityDate, 4, 2) & Left(strActivityDate,2) & space (30)
       
       ' ***********************************************************
       '
       '  Handle Mile
       '
       ' ***********************************************************
       strMile = Rs4("Miles")

       strMile = strMile * 400 
   

       strNewContents = strNewContents & "AC" & strMembership & strFamilyName & strActivityDate & strMile & space(58) & "." &vbCrLF 

     
       Rs4.MoveNext

     Loop

   End if


       ' ***********************************************************
       '
       '  Handle Last Line
       '
       ' ***********************************************************

   ' Handle total number of records
   iSpace = Len(LineNo)  

   ' Number of zero
   ' ***************
   Select case iSpace
  
     Case 1
     
        strZero = "00000"

     Case 2

        strZero = "0000"

     Case 3

        strZero = "000"

     Case 4

        strZero = "00"

     Case 5

        strZero = "0"

     Case Else
      
        strZero = ""

   End Select

 

   LineNo = strZero & LineNo - 1 

   strLastLine = "$$" & LineNo & LineNo & "000000000000000000000000" & space(161) & "."
    
   'response.write  strNewContents

   strNewContents1 = strFirstLine & strNewContents & strLastLine

       
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

       fs.MoveFile sFolder&"\"&x.Name, sFolder&"\001.txt"
  

     next


       ' Get current url
        curPageURL = "http://" & Request.ServerVariables("SERVER_NAME") & "/intranet/recon/001.txt" 


     SQL5 = "Delete From AsiaMileData"

     Conn.Execute(SQL5)

    
       
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
