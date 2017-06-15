<!--#include file="include/SQLConn.inc.asp" -->
<%
' Tells the browser to open excel
Response.ContentType = "text/csv"
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Content-Disposition", "attachment; filename=SaleInOutReport_"&Month(Now())&"_"&Year(now())&".csv" 

if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


Search_From_Month       = Request("From_Month")
Search_From_Year        = Request("From_Year")




'On Error resume Next



     
' Start the Queries
' *****************
     set Rs = server.createobject("adodb.recordset")
              
	Rs.open ("Exec Retrieve_InOutReport '"&Search_From_Month&"', '"&Search_From_Year&"' , '"&iPageCurrent&"' ") ,  conn,3,1

    Set Rs = Rs.NextRecordset() 

        

Response.Write "Material Code, Product, Retailer,  Sale In Volume, Sale In Amount, Sale Out Volume, Sale Out Amount, Difference" & vbcrlf 

If not rs.eof then

' Move to the first record

'rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
Response.Write """" & rs("Material") & """," 
Response.Write """" & rs("ProductName") & """," 
Response.Write """" & rs("Retailer") & """," 
'Response.Write """" & Search_From_Month &"/"& Search_From_Year & """," 
Response.Write """" & rs("SaleInQTY") & """," 
Response.Write """" & rs("SaleInAmount") & """," 
Response.Write """" & rs("SaleOutQTY") & """," 
Response.Write """" & rs("SaleOutAmount") & """," 
Response.Write """" & FormatNumber(Rs("SaleInQTY"),0) - FormatNumber(Rs("SaleOutQTY"),2)  & """" & vbCrLf  


' Move to the next record
rs.movenext
' Loop back to the do statement
loop 

end if

 
%>   
  

 