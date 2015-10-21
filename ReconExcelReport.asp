<!--#include file="include/SQLConn.inc.asp" -->
<%
' Tells the browser to open excel
'Response.ContentType = "application/vnd.ms-excel" 
'Response.addHeader "content-disposition","attachment;filename=Transaction.xls"


if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


Search_From_Month       = Request("From_Month")
Search_From_Year        = Request("From_Year")
Search_Market           = Request("Market")
Search_Instrument       = Request("Instrument")


On Error resume Next


%>

<html>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<body>
<Head>
<STYLE TYPE="text/css">
<!--

TD 
{
  color: black;
  font-family: verdana, Garamond, Times, sans-serif;
  FONT-SIZE: 9px;
  TEXT-ALIGN: left 
}

TD.caption
{
  color: red;
  font-family: verdana, Garamond, Times, sans-serif;
  FONT-SIZE: 9px;
  TEXT-ALIGN: left 
}
-->
</STYLE>
</head>

<div id="Content">


<%


' Start the Queries
' *****************
      
       fsql = "select * from StockReconciliation s left join ReconDepotFolder r on s.depotid = r.depotid "

       fsql = fsql & "left join UOBKHHKEQPRO.dbo.Instrument I on S.ISINCode = I.ISIN "

       fsql = fsql & " Where 1 = 1 "

   
  ' Search by Date
  ' **************


        fsql = fsql & " and left(Importfilename,4) =   '" &Search_From_Month & Right(Search_From_Year,2)& "' " 

        fsql = fsql & " order by s.CreateDate desc"

        response.write fsql
        set frs=conn.execute(fsql)
        'set frs=createobject("adodb.recordset")
		'frs.cursortype=1
		'frs.locktype=1
        'frs.open fsql,conn
 
%>   
  

<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

<tr bgcolor="#ADF3B6" align="center">

      <td width="10">Depot</td>
      <td width="10">Market</td>
      <td>Custodian ID</td>
      <td>Trade Date</td>
      <td>ISINCode</td>
      <td>Common Code</td>
      <td>Security Name</td>
      <td>Description</td>
      <td>Unit Held</td>
      <td>Total Amount</td>
 
</tr>

<%
    i =0

   If Not frs.Eof Then

  do while not frs.EoF

     i=i+1
   
%>


<tr bgcolor="#FFFFCC"> 


<td><% = i & ". " & frs("DepotName") %></td>
<td><% = frs("iMarket") %></td>
<td><% = frs("CustodianID") %></td>
<td><% = frs("TradeDate") %></td>
<td><% = frs("ISINCode") %></td>
<td><% = frs("CommonCode") %></td>
<td><% = frs("SecurityName") %></td>
<td><% = frs("Description") %></td>
<td><% = frs("UnitHeld") %></td>
<td><% = frs("TotalAmount") %></td>
 



</tr>


<%

				
					
				frs.movenext
				
		loop
	
end if	

%>
         
</table>
              </center>

              </body>

              </html>
 