
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")

Title = "Depot File Reconciliation"
%>

<html>
<head>
    <style type="text/css">
    <!-- Hide from legacy browsers
    .print { 
    display: none;
    }
    @media print {
    	.noprint {
    	 display: none;
    	}
    }  -->
    
    </style>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<SCRIPT language=JavaScript>
<!--
function doCovert(){
document.fm1.action="ConvertReconFile.asp?sid=<%=sessionid%>";
document.fm1.submit();
}


function doUpload(what)
{
 window.open('upload_file.asp?sPath='+what,'user','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=20,left=30,width=600,height=400');
}
//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0">



<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">
<%
'-----------------------------------------------------------------------------
'
'      Main Content of the page is inserted here
'
'-----------------------------------------------------------------------------

%>
  <TABLE border=0 cellPadding=0 cellSpacing=0 height=100% width=99%>
    <TBODY> 
    <TR>
      <TD vAlign=top>
        <table width="100%" border="0" cellpadding="1" cellspacing="1" height="100%">
          <tr> 
            <td bgcolor="#000000">
              <table width="100%" border="0" cellpadding=0 cellspacing="0" bgcolor="#FFFFFF" height="100%">
                <tr>
                  <td valign="top" align="center" bgcolor="#E6EBEF">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E6EBEF" class="normal">
                      <tr> 
                          <td valign="top" align="center">
                            <form name=fm1 method=post>
<input name="Location_id" type="hidden" value="" >
                            <table width="99%" border="0" cellspacing="1" bgcolor="#FFFFFF" class="normal">

                              <tr> 
                                <td height="28"> 
                                  <%
		pageid=trim(request.form("pageid"))
		if pageid="" then
		  pageid=1
		end if

' Start the Queries
    
' *****************
      
       SQL1 = "select * from ReconDepotFolder order by DepotName Asc"
       Set Rs1 = Conn.Execute(SQL1)


	  %>
                                </td>
                              </tr>
                              <tr> 
                                <td valign="top" align="center" height="28"> 
   
<table border="1" cellpadding="5" cellspacing="1" class="normal" width="99%">
<tr bgcolor="#006699">
<td width="20%"><font color="#FFFFFF">Depot</font></td>
<td width="30%"><font color="#FFFFFF">Folder</font></td>
<td width="50%"><font color="#FFFFFF">Status</font></td>     
</tr>

<%
      	If Not Rs1.EoF Then
             		
			Do While Not Rs1.EoF
%>

<tr>
<td width="20%">
<% = Rs1("Market") %> - <% = Rs1("DepotName") %>
</td>
<td>
<% = Rs1("DepotFolder") %>
</td>
<td>


<%

 ReadyToConvert = True

 sFolder = Rs1("DepotFolder")

' ---------------------------------------------------------
'                                                          
' Check folder exists                    
'
' ---------------------------------------------------------

 If fs.FolderExists(sFolder)=True then


' ---------------------------------------------------------
'                                                          
' Check file exists when folder exists                   
'
' ---------------------------------------------------------

  If fs.GetFolder(sFolder).Files.Count  = 0 then

    'sFolder1 = replace(sFolder,"\","\\")

     Response.write "Folder is empty." 

     ReadyToConvert = False

  Elseif fs.GetFolder(sFolder).Files.Count  > 1 then

     Response.write "There are more than one file in the folder" 

     ReadyToConvert = False

  Else


    set fo=fs.GetFolder(sFolder)

  for each x in fo.files


' ---------------------------------------------------------
'                                                          
' Check if Depot has set up                  
'
' ---------------------------------------------------------

  
     Sql = "Select count(f.depotid) as Tcount from  (ReconDepotFolder f join reconfileorder o "

     Sql = Sql & "on f.depotid = o.depotid) join ReconFile r on o.fieldid = r.fieldid "

     Sql = Sql & " and f.depotid=" &Rs1("DepotID")
  
     Set Rs = Conn.Execute(Sql)

     If Rs("Tcount") > 0 then


' ---------------------------------------------------------
'                                                          
' Check file extension                  
'
' ---------------------------------------------------------

     If trim(Right(x.Name,3)) = Trim(Rs1("FileType")) Then


' ---------------------------------------------------------
'                                                          
' Check file date                 
'
' ---------------------------------------------------------






          'Print the name of file in the test folder
           Response.write("<b>File Name:</b><br/> "&x.Name& "<br/><br/>")

           set f=fs.OpenTextFile(sFolder&"\"&x.Name,1)

           Response.Write("<b>First line of file:</b><br/>"&f.ReadLine)








     








' ---------------------------------------------------------
'                                                          
' End of checking file type                 
'
' ---------------------------------------------------------

     Else

        Response.Write "Wrong file type."

        ReadyToConvert = False

     End if


' ---------------------------------------------------------
'                                                          
'  End of checking depot's field                
'
' ---------------------------------------------------------


     Else 

        Response.Write "Depot's field setup is not completed."

        ReadyToConvert = False

     End if

' ---------------------------------------------------------
'                                                          
'  Next file in folder               
'
' ---------------------------------------------------------
   
  next

' ---------------------------------------------------------
'                                                          
'  End of Check if file exists              
'
' ---------------------------------------------------------

 End If 

' ---------------------------------------------------------
'                                                          
'  End of folder does not exist            
'
' ---------------------------------------------------------
 
else

  response.write("Folder "& Rs1("DepotFolder") &" does not exist!")

  ReadyToConvert = False

end if

   Response.write "<br/>" & ReadyToConvert


   If ReadyToConvert = False Then

       SQL4 = "Update ReconDepotFolder Set ReadyToConvert = 0 Where DepotID ="&Rs1("DepotID")

   Else
       
     SQL4 = "Update ReconDepotFolder Set ReadyToConvert = 1 Where DepotID ="&Rs1("DepotID")

   End If

       Conn.Execute(SQL4)

%>

</td>
</tr>

<%

    

	Rs1.movenext 

	   loop 
 
	End If


set fs=nothing


%>

                                  </table>
                              
                                </td>
                              </tr>
               <tr> 
                                <td height="28" align="center"> 
<%
			  Rs1.close
			  set Rs1=nothing
            
			  Conn.close
			  set Conn=nothing
%>
                                                         
 </td>
     </tr>
      <tr> 
         <td align="center">
          <input type="Button" value=" Convert" onClick="doCovert();" class="Normal"></td>
                             
</tr>
                            </table>
                          </form>


                          </td>
                        </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
</TABLE>
 </div>
   </body>
</html>