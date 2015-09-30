
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


Title = "Maintenance"
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
function doDelete(){
document.fm1.action="execute.asp?sid=<%=sessionid%>";
document.fm1.whatdo.value = "deleted db";

    if (document.fm1.DeleteMonth.value == "") {
            alert("Please enter the value of month.");
            document.fm1.DeleteMonth.focus();
            return false;
        }
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
    ' Start the queries
' *****************
      
       SQL1 = "select Top 1 StatementDate from ClientStatement order by StatementDate Asc"
       Set Rs1 = Conn.Execute(SQL1)

       SQL2 = "select Top 1 StatementDate from ClientStatement order by StatementDate Desc"
       Set Rs2 = Conn.Execute(SQL2)

       SQL3 = "select Top 1 ConfirmationDate from DetailTrade order by ConfirmationDate Asc"
       Set Rs3 = Conn.Execute(SQL3)

       SQL4 = "select Top 1 ConfirmationDate from DetailTrade order by ConfirmationDate Desc"
       Set Rs4 = Conn.Execute(SQL4)

       SQL5 = "select Top 1 TradeDate from TransactionHistory order by TradeDate  Asc"
       Set Rs5 = Conn.Execute(SQL5)

       SQL6 = "select Top 1 TradeDate  from TransactionHistory order by TradeDate  Desc"
       Set Rs6 = Conn.Execute(SQL6)

	  %>
                                </td>
                              </tr>
                              <tr> 
                                <td valign="top" align="center" height="28"> 
   
<table border="0" cellpadding="5" cellspacing="1" class="normal" width="99%">
<tr bgcolor="#006699">
<td width="20%"><font color="#FFFFFF">Database</font></td>
<td width="40%"><font color="#FFFFFF">Trade Date (From)</font></td>
<td width="40%"><font color="#FFFFFF">Trade Date (To)</font></td>     
</tr>
<tr>
<td width="20%">
Client Statement
</td>
<td>
<%= Rs1("StatementDate") %>
</td>
<td>
<%= Rs2("StatementDate") %>
</td>
</tr>

<tr>
<td width="20%">
Detail Trade
</td>
<td>
<%= Rs3("ConfirmationDate") %>
</td>
<td>
<%= Rs4("ConfirmationDate") %>
</td>
</tr>

<tr>
<td width="20%">
Transaction History
</td>
<td>
<%= Rs5("TradeDate") %>
</td>
<td>
<%= Rs6("TradeDate") %>
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
              Rs2.close
			  set Rs2=nothing
              Rs3.close
			  set Rs3=nothing
              Rs4.close
			  set Rs4=nothing
              Rs5.close
			  set Rs5=nothing
              Rs6.close
			  set Rs6=nothing
			  Conn.close
			  set Conn=nothing
%>
                                                         
 </td>
     </tr>
      <tr> 
         <td valign="top">Delete the transaction records which are longer than　<input name="DeleteMonth" type=text value="" size="2" class="Normal"> 
								months   
          <input type="Button" value=" Submit" onClick="doDelete();" class="Normal">&nbsp;&nbsp; </td>
                              <input type="hidden" value="" name="whatdo">
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