<% Response.Buffer = False %>
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")

Title = "Depot File Reconciliation"



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

    sql1 = "Update ReconDepotFolder set ReadyToConvert = 0 where DepotID = "&delete_depotid

    Conn.Execute(sql1)

  

	
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
function doCovert(Flat){
document.fm1.action="ConvertReconFile.asp?sid=<%=sessionid%>";
document.fm1.submit();
}

function doDelete(what){
document.fm1.action="ReconDepotFile.asp?sid=<%=sessionid%>&depotid="+what;
document.fm1.action_button.value="deleteFile";
document.fm1.submit();
}

function setFocus(){
    document.getElementById("getFocus").focus();
}

function doUpload(what)
{
  document.fm1.action="upload_file.asp?sid=<%=sessionid%>&depotid="+what;
  document.fm1.submit();
}
//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" onload='setFocus()'>



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
   
<table border="1" cellpadding="6" cellspacing="0" class="normal" width="99%">
<tr bgcolor="#006699">
<td width="20%"><font color="#FFFFFF">Depot</font></td>
<td width="70%"><font color="#FFFFFF">Status</font></td> 
<td width="10%"><font color="#FFFFFF">Action</font></td>     
</tr>

<%
      	If Not Rs1.EoF Then
             		
			Do While Not Rs1.EoF

 
%>
<tr <% If Trim(delete_depotid) = Trim(Rs1("DepotID")) then%>bgcolor="#ffccff"<% end if%>>
<td width="20%">
<% = Rs1("DepotID") %>. <% = Rs1("DepotName") %>
</td>

<td>

<%

 FileIsEmpty = False

 FileExists = False


 sFolder = Server.MapPath(Rs1("DepotFolder"))


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

     Response.write "Folder is empty." 

     FileIsEmpty = True

     ReadyToConvert = False


  Else


    set fo=fs.GetFolder(sFolder)

  for each x in fo.files


' ---------------------------------------------------------
'                                                          
' Check if Depot has set up                  
'
' ---------------------------------------------------------

  
     ' Check if view exist
     sql = "Select count(*) as count1 FROM sys.views WHERE name = 'vw_"&Rs1("DepotID")&"'"

     'response.write sql

     Set Rs = Conn.Execute(sql)

     If Rs("count1") = 1 then


' ---------------------------------------------------------
'                                                          
' Check file date                 
'
' ---------------------------------------------------------

' Remove comma in sting

  'If Trim(Rs1("ReadyToConvert")) = "True" then

   ' <!--#include file="include/remove_comma.inc.asp" -->

%>

   

<%  
    
     'response.redirect "ConvertReconFile.asp?depotid="&Rs1("DepotID")&"&sid="&sessionid

   'End If

 

           FileExists = True

          'Print the name of file in the test folder
           Response.write(x.Name)

 


' ---------------------------------------------------------
'                                                          
'  End of checking depot's field                
'
' ---------------------------------------------------------


     Else 

        Response.Write "Depot's profile not created."

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

 

%>

</td>
<td>
<% If FileIsEmpty = True then %>
           
               <input type="Button" value=" Upload " onClick="doUpload(<%=Rs1("DepotID")%>);" class="Normal">

<% End If %>

<% If FileExists = True then %>
           
               <input type="Button" value=" Delete " onClick="doDelete(<%=Rs1("DepotID")%>);" class="Normal">

               <input type="hidden" name="DepotID" value="<% = Rs1("DepotID") %>">  

<% End If %>
</td>
</tr>

<%

    

	Rs1.movenext 

	   loop 
 
	End If


set fs=nothing


%>

<tr bgcolor="#006699">
<td width="20%"><font color="#FFFFFF"></font></td>
<td width="70%"><font color="#FFFFFF"></font></td> 
<td width="10%">



</td>     
</tr>

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
          <input type="Button" value=" Convert" onClick="doCovert();" class="Normal">
              <input type="hidden" name="action_button" value="">   
</td>
                             
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