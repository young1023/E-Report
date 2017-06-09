<!--#include file="include/SQLConn.inc.asp" -->
<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=SaleInOutReport_"&Request("From_Month")&Right(Request("From_Year"),2)&".xls"


if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


Search_From_Month       = Request("From_Month")
Search_From_Year        = Request("From_Year")




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
     set Rs1 = server.createobject("adodb.recordset")

    'response.write ("Exec Retrieve_InOutReport '"&Search_From_Month&"', '"&Search_From_Year&"' , '"&iPageCurrent&"' ") 
              
	Rs1.open ("Exec Retrieve_InOutReport '"&Search_From_Month&"', '"&Search_From_Year&"' , '"&iPageCurrent&"' ") ,  conn,3,1

     Set Rs1 = Rs1.NextRecordset() 
 
%>   
  
<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

<tr bgcolor="#ADF3B6" align="center">
      
      <td></td>
      <td>Material Code</td>
      <td>Product Name</td>
      <td>Retailer</td>
      <td>Month/Year</td>
      <td>Sale In Volume</td>
      <td>Sale In Amount</td>
      <td>Sale Out Volume</td>
      <td>Sale Out Amount</td>
      <td>Difference</td> 
      
        
</tr>
<%		
    i=0
  do while Not Rs1.EoF

   if Rs1.eof then exit do

   i=i+1
		
%>
<tr bgcolor="#FFFFCC"> 
<td>
<%
        Response.write i 
 
%>
</td>

<td><%  = Rs1("Material")  %></td> 

<td>
<%  = Rs1("ProductName") %>
</td>

<td><%  = Rs1("Retailer")  %></td> 

<td><%  = Search_From_Month & "/" & Search_From_Year %></td> 


<td>
<% = FormatNumber(Rs1("SaleInQTY"),0) %>
</td>

<td>

<% = FormatNumber(Rs1("SaleInAmount"),2) %>

</td>

<td>
<% = Rs1("SaleOutQTY") %>
</td>


<td><% = FormatNumber(Rs1("SaleOutAmount"),2) %></td>
<td ></td>



</tr>


<%

				
					
				Rs1.movenext
				
		loop
	
		

    If Search_Match => 2 Then

   Set rs1 = rs1.NextRecordset() 
    i=0

  do while Not Rs1.EoF

   if Rs1.eof then exit do

   i=i+1
		
%>
<tr bgcolor="#FFFFCC"> 
<td>
<%
   
        Response.Write Rs1("DepotName")


 
%>
</td>
<td>
<%
   
        Response.Write Rs1("DepotCode")


 
%>
</td>


<td><%
       If Rs1("Instrument") <> "" Then
       
        Response.Write Rs1("Instrument")
        
       Elseif Rs1("ISIN") <> "" Then
       
        Response.Write Rs1("ISIN")
        
       Else
       
        Response.Write Rs1("Sedol")
        
       End If


 
%></td>

<td>
<% 

        sql2 = "Select ShortName from UOBKHHKEQPRO.dbo.Instrument where "
        
        If Trim(Rs1("Instrument")) <> "" then

        sql2 = sql2 & " Instrument = '" & Trim(Rs1("Instrument")) & "'"
        
        Elseif Trim(Rs1("ISIN")) <> "" then
        
        sql2 = sql2 & " ISIN = '" & Trim(Rs1("ISIN")) & "'"
        
        Else 
        
        sql2 = sql2 & " Sedol = '" & Trim(Rs1("Sedol")) &"'"
        
        End if
             
        'response.write sql2

        Set Rs2 = Conn.execute(sql2)
        
        If not Rs2.Eof then

        Response.Write Rs2("ShortName")
        
        Else
        
        Response.write "Instrument name cannot be found."
        
        End if

%>
</td>


<td><% = formatnumber(Rs1("UnitHeld"),0) %></td>
<td ><% = formatnumber(Rs1("TotalQTY"),0)%></td>
<td><% = formatnumber((formatnumber(Rs1("UnitHeld"),0) - formatnumber(Rs1("TotalQTY"),0)),0)   %></td> 



</tr>


<%

				
					
				Rs1.movenext
				
		loop
	


 Set rs1 = rs1.NextRecordset() 

    i=0

  do while Not Rs1.EoF

   if Rs1.eof then exit do

   i=i+1
		
%>
<tr bgcolor="#FFFFCC"> 
<td>
<%
   
        Response.Write Rs1("DepotName")


 
%>
</td>
<td>
<%
   
        Response.Write Rs1("DepotCode")


 
%>
</td>

<td><%
       Response.Write Rs1("StockCode")


 
%></td>

<td>
<% 



        Response.Write Rs1("InstrumentName")

%>
</td>



<td><% = formatnumber(Rs1("UnitHeld"),0) %></td>
<td ><% = formatnumber(Rs1("TotalQTY"),0)%></td>
<td><% = formatnumber((formatnumber(Rs1("UnitHeld"),0) - formatnumber(Rs1("TotalQTY"),0)),0)   %></td> 

</tr>


<%

				
					
				Rs1.movenext
				
		loop
	
	
'End if

End If

%>         
</table>
              </center>

              </body>

              </html>
 