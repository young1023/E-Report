<% 
Title = "Upload File"

if session("shell_power")="" then
  response.redirect "default.asp"
end if
%>

<!--#include file="include/conne.inc" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<SCRIPT language=JavaScript>
<!--
function dosubmit(){
 document.fm1.action="hsemis.asp?page_id=execute";
 if (document.fm1.Location.value == "")

document.fm1.submit();
}

function doReturn(){
document.fm1.action="sa_group.asp";
document.fm1.submit();
}


function dosave(){
document.fm1.action="hsemis.asp?page_id=execute";
document.fm1.whatdo.value="Modify Group";
document.fm1.submit();
}


//-->
</SCRIPT>


</head>
<body leftmargin="0" topmargin="0" OnLoad="document.fm1.Name.focus();">

<!--#include file="include/logo.inc" -->

<table width="100%" border="0" cellspacing="0" cellpadding="0" height="60%">
    <tr>
      <td width="160" valign="top">

<!-- #include file ="include/menu.inc" -->

      </td>
      <td width="1" class="HSEBlue"></td>
      <td valign="top">
      <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="4" class="HSEBlue"></td>
        </tr>
        <tr valign="top">
          <td height="25"><img src="images/Curve.gif" width="22" height="16" /></td>
        </tr>
        <tr valign="top">
          <td height="100%" align="middle">
<%
'-----------------------------------------------------------------------------
'
'      Main Content of the page is inserted here
'
'-----------------------------------------------------------------------------


submit_date = datevalue(date())
submit_time = hour(now())


if id <> "" then
step = 2
         sql = " Select * from UserGroup where GroupID="&id
         'response.write sql
         set rs = conn.execute(sql)

Name = rs("Name")
Description = rs("Description")
end if




%>




<form name="fm1" method="post" action="">

<input name="Location_id" type="hidden" value="" ><input name="Location_id" type="hidden" value="<% = site %>" >

  <table width="99%" border="0" class="normal">
    <tr> 
      <td class="BlueClr"></td> 
    </tr>
    <tr> 
      <td></td>
    </tr>
  <tr> 
      <td  align="right">
<font color="red">*</font>	 Denotes a mandatory field</td>
    </tr>
  <tr> 
      <td  align="right">
　</td>
    </tr>
    
 <tr> 
      <td>
　</td>
    </tr>
    
 <tr> 
      <td align="center">
Please upload the Excel file here <br>
<FORM METHOD="POST" ACTION="upphoto.asp" ENCTYPE="multipart/form-data">
   <input type="hidden" name="flag" size="23" value="<% = id %>" ><br><br>
   <INPUT TYPE="FILE" NAME="FILE1" SIZE=30"><BR><br>
   <INPUT TYPE="SUBMIT" VALUE="      Upload   ">
</FORM>
</td>
    </tr>
    
 <tr> 
      <td>
      	     
　</td>
    </tr>
    
 <tr> 
      <td>
      	     
　</td>
    </tr>
</table>
</form>

<%
'-----------------------------------------------------------------------------
'
'      End of the main Content 
'
'-----------------------------------------------------------------------------
%>
</td>
              </tr>
                </table>
                </td>
                </tr>
              </table>
              
              <center>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td colspan="2" height="1"></td>
                </tr>
                <tr class="HSEBlue">
                  <td colspan="2" height="1"></td>
                </tr>
              </table>
              </center>
            
            </td>
                </tr>
              </table>
              
              </body>

              </html>