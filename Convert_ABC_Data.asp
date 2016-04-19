<!--#include file="include/SessionHandler.inc.asp" -->
<%


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

                     Converting in Process. Please wait.
                     
          </td>

        </tr>
        <tr>
          <td align="middle">

<%


   set RsABC = server.createobject("adodb.recordset")
   RsABC.open ("Exec Process_ReconMonthly") ,  StrCnn,3,1


   response.redirect "ReconReport.asp?sid="&sessionid
      

%>


</td>
  </tr>
 

     </table>

</div>

 </body>
    </html>