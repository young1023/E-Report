<!--#include file="include/SessionHandler.inc.asp" -->
<!-- #include file="ShadowUpload.asp" -->
<%

DepotID = Request("DepotId")

Response.write Month(Session("DBLastModifiedDate")) 

' Check folder
    
' *****************
      
       SQL1 = "select * from ReconDepotFolder where DepotId="&depotId

       Set Rs1 = Conn.Execute(SQL1)

            sFolder = Trim(Rs1("DepotFolder"))

dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")
%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<TITLE>Upload File</TITLE>
<SCRIPT language=JavaScript>
<!--

  function doReturn(){
  alert("You must select a record!");
  document.fm1.action="ReconDepotFile.asp?sid=<%=sessionid%>;
  document.fm1.submit();
  }

//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">


      <table width="98%" height="400" border="1" cellspacing="0" cellpadding="2" class="Normal">
        <tr>
          <td align="middle" height="50">

Upload file to <% = Rs1("Depotname") %>
          </td>

        </tr>
        <tr>
          <td align="middle">

<%

'Response.Write Trim(Rs1("FileType"))

Dim objUpload

If Request("action")="1" Then

    Set objUpload=New ShadowUpload

    If objUpload.GetError<>"" Then

        Response.Write("sorry, could not upload: "&objUpload.GetError)

    Else  


        For x=0 To objUpload.FileCount-1

            Response.Write("file name: "&objUpload.File(x).FileName&"<br /><br />")
            Response.Write("file size: "&objUpload.File(x).Size&"<br /><br />")

            
        
            If (objUpload.File(x).ImageWidth>200) Or (objUpload.File(x).ImageHeight>200) Then

                Response.Write("image to big, not saving!")

            Elseif Trim(Rs1("FileType")) <> Trim(Right(objUpload.File(x).FileName, 3)) Then

%>
              

       <font color=red><b>Wrong file Type!</b></font><br/><br/>

      <a href="ReconDepotFile.asp?sid=<%=sessionid%>">Return</a>
              
<%

            'Elseif  CInt(Left(objUpload.File(x).FileName, 2)) <>  CInt(Month(now())) - 3 Then
 
               'Response.Write "Only file from last month is allowed.<br/><br/>"


            Else  

              


         ' Check if distinction file exists
         If fs.FileExists(Server.MapPath(sFolder) & objUpload.File(x).FileName)  Then

              fs.DeleteFile(Server.MapPath(sFolder) & objUpload.File(x).FileName)
 
         end if
    
                'txt file. use different method
                If Trim(Right(objUpload.File(x).FileName, 3)) = "txt" Then


                      Call objUpload.File(x).SaveToDisk(Server.MapPath(sFolder) , "")

  response.redirect "ConvertTxTFile.asp?depotid="&Rs1("DepotID")&"&sid="&sessionid


                Else

                       Call objUpload.File(x).SaveToDisk(Server.MapPath(sFolder) , "")

  response.redirect "ConvertReconFile.asp?depotid="&Rs1("DepotID")&"&sid="&sessionid


                End If

    
   
            End If

        Next


    End If
End If


Set objUpload = Nothing


      

%>


</td>
  </tr>
 

     </table>

</div>

 </body>
    </html>