<% Response.Buffer = False %>
<!--#include file="include/SessionHandler.inc.asp" -->
<%

if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

dim fs, fo, ts, f
set fs=Server.CreateObject("Scripting.FileSystemObject")
Title = "Shell File Conversion"


' Delete File
'***************
if trim(request("action_button")) = "deleteFile" then
	delete_depotid = trim(request("depotid"))
       'response.write depotid
        sql = "Select DepotFolder from ReconDepotFolder where DepotId="&delete_depotid
        Set Rs = Conn.Execute(sql)
         'response.write Rs("DepotFolder")
  	     set fo=fs.GetFolder(Server.MapPath(Rs("DepotFolder")))
               for each x in fo.files
 
                fs.DeleteFile(Server.MapPath(Rs("DepotFolder"))&"\"&x.Name)
               next
	
end if
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

function doDelete(what){
document.fm1.action="ShipConvert.asp?sid=<%=sessionid%>&depotid="+what;
document.fm1.action_button.value="deleteFile";
document.fm1.submit();
}



function doUpload(what)
{
  document.fm1.action="upload_Daily_file.asp?sid=<%=sessionid%>&depotid="+what;
  document.fm1.submit();
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
  
          

                       <form name=fm1 method=post>

                            <table width="100%" border="0" cellspacing="2" bgcolor="#FFFFFF" class="normal">

                                 
<%




      ' Start the Queries
      '
      ' *************************************
      
       SQL1 = "select * from ReconDepotFolder where DepotCode = '2016' order by Depotcode Asc"

       Set Rs1 = Conn.Execute(SQL1)

%>
                               
               <tr> 
                                
                    <td valign="top" align="center" height="28"> 

   
             <table border="1" cellpadding="10" cellspacing="0" class="normal" width="99%">

                   <tr bgcolor="#006699">

                         <td width="20%"><font color="#FFFFFF">Name</font></td>

                            <td width="50%"><font color="#FFFFFF">Status</font></td> 

                                <td width="20%"><font color="#FFFFFF">Action</font></td>     
                    </tr>

<%
      	If Not Rs1.EoF Then
             		
			Do While Not Rs1.EoF
 
%>
             <tr <% If Trim(delete_depotid) = Trim(Rs1("DepotID")) then%>bgcolor="#ffccff"<% end if%>>

                       <td width="20%">

                            <% = Rs1("DepotName") %>
                         </td>

                         <td>

<%

 FolderIsEmpty = False
 FileExists = False
 ProFileSet = True

 sFolder = Server.MapPath(Rs1("DepotFolder"))
' ---------------------------------------------------------
'                                                          
' Check folder exists                    
'
' ---------------------------------------------------------
 If fs.FolderExists(sFolder) = False then

  response.write("Folder "& Rs1("DepotFolder") &" does not exist!")

  
   
 Else



' ---------------------------------------------------------
'                                                          
' Check file exists              
'
' ---------------------------------------------------------
  If fs.GetFolder(sFolder).Files.Count  > 0 then


        

      set fo=fs.GetFolder(sFolder)
  For each x in fo.files

       
' ---------------------------------------------------------
'                                                          
' Check file date                 
'
' ---------------------------------------------------------
          FileExists = True
          'Print the name of file in the test folder
          'Response.write(x.Name)
          
   ' Browse Error Log     
      
           If Left(x.Name,3) = "Err" Then
    'response.write (sFolder & x.Name)
     Set objFile = fs.OpenTextFile((sFolder & "\" & x.Name),1,True)
	Err_Msg = objFile.ReadAll
    response.write Err_Msg & "<br/>"
   
    'Audit Log
    Conn.Execute "Exec AddReconLog 'Error message : <b>" & Err_Msg & "</b> on converting file <b>" & x.Name & "</b>','" & Session("MemberID") & "'"
    
            End if

      fs.DeleteFile(sFolder&"\"&x.Name)


  Next





   Else   ' If no file exists


     Sql2 = "Select top 1 ImportFileName, CreateDate from StockReconciliation where depotid ="&Rs1("DepotID")&" order by CreateDate desc"
     Set Rs2 = Conn.Execute(Sql2)

     If Not Rs2.EoF Then
     Response.Write "Latest import file " & Rs2("ImportFileName") & " was imported on "& Rs2("CreateDate")
     Else
     Response.Write "No File was imported before."
     End If

     FolderIsEmpty = True


 End If 


End If ' FolderExist
 

 
%>

</td>
<td>
<% If FolderIsEmpty = True and ProFileSet = True then %>
           
               <input type="Button" value=" Upload " onClick="doUpload(<%=Rs1("DepotID")%>);" class="Normal">

<% End If %>

<% If FileExists = True then %>

               <input type="Button" value=" Upload " onClick="doUpload(<%=Rs1("DepotID")%>);" class="Normal">
           
               <input type="Button" value=" Delete " onClick="doDelete(<%=Rs1("DepotID")%>);" class="Normal">&nbsp;

               <input type="hidden" name="DepotID" value="<% = Rs1("DepotID") %>">  

<% End If %>
<input type="hidden" name="action_button" value="">   

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
                            </table>
                          </form>


            
       
  

<%
			  
Rs1.close			  
set Rs1=nothing
            			  
Conn.close			  
set Conn=nothing

%>

 </div>
   </body>
</html>