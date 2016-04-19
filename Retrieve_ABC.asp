<!--#include file="include/SessionHandler.inc.asp" -->
<%
  
        On Error resume Next
        
        '  Collect Date values
        Search_From_Month       = Request.form("FromMonth")
        Search_From_Year        = Request.form("FromYear")
        action_button           = Request.Form("action_button")
        
      

  
%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<TITLE>Retrieve ABC Position</TITLE>

<SCRIPT language=JavaScript>
<!--

    function dosubmit(what) {
        document.fm1.action = "Retrieve_ABC.asp?sid=<%=SessionID%>";
        document.fm1.action_button.value = 1;
        document.fm1.submit();
    }
 //-->
</script>

</head>
<body leftmargin="0" topmargin="0">

<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">
<form name="fm1" method="post" action="">

      <table width="98%" height="400" border="0" cellspacing="0" cellpadding="2" class="Normal">
        
        <tr>
          <td align="middle">

<%

                          
                         
             if action_button <> "" then
             
              
                  
              set Rs1 = server.createobject("adodb.recordset")
              
             ' response.Write  ("Exec Process_ReconMonthly2 '"&Search_From_Month&"', '"&Search_From_Year&"' ")

              
	         Rs1.open ("Exec Process_ReconMonthly2 '"&Search_From_Month&"', '"&Search_From_Year&"' ") ,  conn,3,1
	         
	
	          
	              Response.Write "Success."
	              
	              Response.Write "<br/>"
	              
	              Response.Write Rs1("Description")
	              
	              Response.Write Rs1("ERROR_MESSAGE")
	         
      
	         
	        
	         else 
	        
	        
	        
	       Response.Write "ABC record of " &  Search_From_Month & "-" &  Search_From_Year &" will be retrieved."
	       
	       
	       Response.Write "<p>" &  "It may take several minutes to process." & "</p>"
	       
	     
	       
	       Response.Write "<p>" & "Press ok to confirm." & "</p>"
	       
	       
%>

      <tr>
 <td align="center">
    <input id="Button1" type="button" value="OK" onClick="dosubmit();">

 
 </td>
 </tr>
 
<%        


              end if

%>


</td>
  </tr>
   <tr>
 <td align="center">
   
 	<input type=hidden   value=""   name="action_button"> 
 	<input type=hidden   value="<% = Search_From_Month %>"   name="FromMonth"> 
 	<input type=hidden   value="<% = Search_From_Year %>"   name="FromYear"> 
 
 </td>
 </tr>
 <tr>
 <td align="center"><a href = "ReconReport.asp?sid=<%=SessionID%>" style="TEXT-DECORATION: none">Return</a>
 </td>
 </tr>

     </table>
  </form>
</div>

 </body>
    </html>