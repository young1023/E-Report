<!--#include file="include/SQLConn.inc.asp" -->
<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=Convert_Report_"&Request("From_Month")&Right(Request("From_Year"),2)&".xls"


if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


Search_From_Month       = Request("From_Month")
Search_From_Year        = Request("From_Year")
Search_Market           = Request("Search_Market")
Search_Instrument       = Request("Instrument")
Search_ISIN             = Request("ISIN")
Search_Sedol            = Request("Sedol")
Search_Match            = Request("Search_Match")
Search_DepotCode        = Request("DepotCode")



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

     'response.write ("Exec Retrieve_MonthReport '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_Market&"', '"&Search_DepotCode&"', '"&Search_Match&"' , '1' ") 

              
	 Rs1.open ("Exec Retrieve_MonthReport '"&Search_From_Month&"', '"&Search_From_Year&"', '"&Search_Market&"', '"&Search_DepotCode&"', '"&Search_Match&"' , '1' ") ,  conn,3,1

     Set Rs1 = Rs1.NextRecordset() 
 
%>   
  
<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

<tr bgcolor="#ADF3B6" align="center">
      
      <td width="20%" >Depot</td>
      <td width="7%" >Depot Code</td>
      <td width="7%" >Location</td>
      <td width="7%" >STOCK Code</td>
      <td width="7%">Local Code</td>
      <td width="30%">Instrument Name</td>
      <td>Status</td>
      <td>ABC Position</td>
      <td>Custodian Position</td>
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
   
        Response.Write Rs1("DepotName")


 
%>
</td>
<td>
<%
   
        Response.Write Rs1("DepotCode")


 
%>
</td>

<td>
<%
   
        Response.Write Rs1("Location")


 
%>
</td>
<td>
<%
       If Rs1("Instrument") <> "" Then
       
        Response.Write Rs1("Instrument")
        
       Elseif Rs1("ISIN") <> "" Then
       
        Response.Write Rs1("ISIN")
        
       Else
       
        Response.Write Rs1("Sedol")
        
       End If


 
%>
</td>
<td><% = Rs1("Local Code") %></td>
<td>
<% 



        Response.Write Rs1("InstrumentName")

%>
</td>

<td></td>
<td><% = formatnumber(Rs1("UnitHeld"),0) %></td>
<td ><% = formatnumber(Rs1("TotalQTY"),0)%></td>
<td><% = formatnumber((formatnumber(Rs1("UnitHeld"),0) - formatnumber(Rs1("TotalQTY"),0)),0)   %></td>
 



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

<td></td>

<td><%
       If Rs1("Instrument") <> "" Then
       
        Response.Write Rs1("Instrument")
        
       Elseif Rs1("ISIN") <> "" Then
       
        Response.Write Rs1("ISIN")
        
       Else
       
        Response.Write Rs1("Sedol")
        
       End If


 
%></td>
<td></td>
<td>
<% 

        sql2 = "Select InstrumentName from ReconMonthly where "

        sql2 = sql2 & "(ISIN = '" & Rs1("ISIN") &"' or Instrument = '" & Rs1("Instrument") & "' )"

        Set Rs2 = Conn.execute(sql2)

        Response.Write Rs2("InstrumentName")

%>
</td>
<td></td>

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
<td></td>
<td><%
       If Rs1("Instrument") <> "" Then
       
        Response.Write Rs1("Instrument")
        
       Elseif Rs1("ISIN") <> "" Then
       
        Response.Write Rs1("ISIN")
        
       Else
       
        Response.Write Rs1("Sedol")
        
       End If


 
%></td>
<td></td>
<td>
<% 



        Response.Write Rs1("InstrumentName")

%>
</td>


<td></td>

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
 