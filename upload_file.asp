<%

dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")

sFolder = request("sPath")

response.write sFolder

'id = session("id")
%>
<!DOCTYPE html>
<HTML>
<HEAD>
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>Upload File</TITLE>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<script type="text/javascript">
    function CloseWindow() {
        window.close();
        window.opener.location.reload();
    }
</script>
</head>
<body leftmargin="0" topmargin="0">

      <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="2" class="Normal">
        <tr>
          <td align="middle">
upload<br>
<FORM METHOD="POST" ACTION="upphoto.asp" ENCTYPE="multipart/form-data">
   <input type="hidden" name="flag" size="23" value="<% = id %>" >
   <p><INPUT TYPE="FILE" NAME="FILE1" SIZE="30"></p>
   <p>
   <INPUT TYPE="submit" VALUE="     Upload    ">
</p>
</FORM>
</td>
  </tr>
 
<%
   set fo=fs.GetFolder(sFolder)

  for each x in fo.files

%>

<tr>
<td align="middle">

<%

  'Print the name of all files in the test folder

  Response.write("<b>File Name:</b><br/> "&x.Name& "<br/><br/>")

  set f=fs.OpenTextFile(sFolder&"\"&x.Name,1)

  Response.Write("<b>First line of file:</b><br/>"&f.ReadLine)

  
  next

%>

</td>
   </tr>

<tr><td align=center valign=center>
   <INPUT TYPE="BUTTON" VALUE="      Close Window" onclick="javascript: return CloseWindow();" /></td></tr>
     </table>
 </body>
    </html>