<% Response.Buffer = False %>
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

response.expires=0


dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")

Title = "Daily Sales File Conversion"

DepotID = trim(Request("DepotID"))


' Retrieve Folder
'****************

   SQL1 = "select * from ReconDepotFolder where DepotId = 69"

   Set Rs1 = Conn.Execute(SQL1)

   If Not Rs1.EoF Then

   DepotFolder          = trim(Rs1("DepotFolder"))

   End if

  
%>   

<html>
<head>
<title></title>
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
       sFolder = Trim(Server.MapPath(DepotFolder))


            set fo=fs.GetFolder(sFolder)

            for each x in fo.files  

             FileName = x.Name

        sql_d = "Delete from DailySales where FileName like '%"&FileName&"'"

        Conn.execute(sql_d)
            
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


        'Retrieve Year
        If delimiterNo = 1 Then
      
        strYear  = replace(strYear,",","") & strCharacter 

        End if


        'Retrieve Month
        If delimiterNo = 2 Then
      
        strMonth  = replace(replace(strMonth,",","")," ","_") & strCharacter 

        End if

        'Retrieve Product ID
        If delimiterNo = 6 Then
      
         strProductID  = replace(replace(strProductID,",",""),".","/") & strCharacter 

        End if

        'Retrieve Product EAN Code
        If delimiterNo = 7 Then
      
          strEANCode = replace(replace(strEANCode,",","")," ","") & strCharacter

        End if

       ' Retrieve QTY
       If delimiterNo = 13 Then
      
           strQTY = replace(strQTY,",","") & strCharacter

        End if

       ' Retrieve Total of Sales
       If delimiterNo = 15 Then
      
           strSales = replace(strSales,",","") & strCharacter

        End if

    
    Next

     If Len(strMonth) = 1 then

         strMonth = "0" & strMonth

     End If


     If QTY <> "," Then


     SQL2 = "Insert into DailySales (Year, Month, ProductID, EANCode, SaleQTY, TotalSale, FileName) Values "

     SQL2 = SQL2 & "( '" & strYear &"' , '" & strMonth &"' , '" & strProductID &"' , '" & strEANCode &"'  "

     SQL2 = SQL2 & " , ' " & trim(strQTY) & "', ' " & trim(strSales) & "' , ' " & trim(FileName) & "' )"

     'Response.write "Write into database :" & SQL2 & "<br/>"

     Conn.Execute(SQL2)

    End If 


    ' reset character
    strYear          = ""
    strMonth         = ""
    strProductID     = ""
    strEANCode       = ""
    strQTY           = ""
    strSales         = ""
    
    


    ' end of line number > 1
    End if
  

Loop


  

         ' Record if there is error
         If Err.Number <> 0 Then
  
         'Audit Log
         Conn.Execute "Exec AddReconLog 'convert error " & Err.Description & " on file " & x.Name & " ','" & Session("MemberID") & "'"
         
         End If

    Next

 set fs=nothing
 Rs1.Close
 set Rs1 = Nothing
 Conn.Close
 Set Conn = Nothing
    
         response.redirect "DailyCheckList.asp?depotid="&depotid&"&sid="&sessionid



%>



</td>
</tr>
</table>
            
    
</div>
       
</body>
</html>
