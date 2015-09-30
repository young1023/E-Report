
<%
Title = "Forgot password"
%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<script language="JavaScript">
    function dosubmit() {
	document.fm1.action="hsemis.asp?page_id=execute";
	document.fm1.whatdo.value="forgot passord";
       
        if (document.fm1.num.value == "") {
            alert("Please enter your login name.");
            document.fm1.num.focus();
            return false;
        }

      document.fm1.submit();
    }
//-->
</script>
</head>
<body leftmargin="0" topmargin="0" OnLoad="document.fm1.num.focus();">

<!-- #include file ="include/Master.inc.asp" -->
<div id="Content">


<form name="fm1" method="post" action="">
		<table border="0" width="80%" class="Normal" id="table1" height="300">
		<tr>
		<td height="89">If you are having any difficulty on logging in, 
		please contact client service department.</td>
	</tr>

</table>
</form>
</div>
            
              </body>

              </html>