<%
'Check to see if title has been entered or not
u_title=request.form("u_title")
if u_title = "" then
%>
<html>
<body bgcolor="#FFFFFF">
<!-- Input form area - This will only display when no Title has been entered -->
<form method="POST" action="<%= request.servervariables("script_name") %>">
Document Title<br>
<input type="text" name="u_title" size="35">
<br><br>
Cell 1
<br>
<textarea rows="2" name="u_cell1" cols="35"></textarea>
<br><br>
Cell 2
<br>
<textarea rows="2" name="u_cell2" cols="35"></textarea>
<input type="submit" value="Submit" ></p>
</form>
<%
else

' If there is a user inputted title
' get all of the user inputed values
u_title=request.form("u_title")
u_cell1=request.form("u_cell1")
u_cell2=request.form("u_cell2")

' Varible created fo excel file name. Speces are changed to underscores
' and later the current date is added in attempts to create a unique file
' Users are not prevented from entering characters !@#$%^&*()+= that are 
' invlaid file names in this example
g_filename=replace(u_title," ","_")


set fso = createobject("scripting.filesystemobject")
' create the text (xls) file to the server adding the -mmddyyyy after the g_title value
Set act = fso.CreateTextFile(server.mappath(""&g_filename & "-"& month(date())& day(date())& year(date()) &".xls"), true)

' write all of the user input to the text (xls) document 
' The .xls extension can just as easily be .asp or .inc whatever best suits your needs
' Providing that you remove the info contained in the header and remove the xml
' reference in the html tag that starts the page/excel file. It is to add gridlines and
' a title to the excel worksheet
act.WriteLine "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
act.WriteLine "<head>"
act.WriteLine "<!--[if gte mso 9]><xml>"
act.WriteLine "<x:ExcelWorkbook>"
act.WriteLine "<x:ExcelWorksheets>"
act.WriteLine "<x:ExcelWorksheet>"
act.WriteLine "<x:Name>"& u_title &"</x:Name>"
act.WriteLine "<x:WorksheetOptions>"
act.WriteLine "<x:Print>"
act.WriteLine "<x:ValidPrinterInfo/>"
act.WriteLine "</x:Print>"
act.WriteLine "</x:WorksheetOptions>"
act.WriteLine "</x:ExcelWorksheet>"
act.WriteLine "</x:ExcelWorksheets>"
act.WriteLine "</x:ExcelWorkbook>"
act.WriteLine "</xml>"
act.WriteLine "<![endif]--> "
act.WriteLine "</head>"
act.WriteLine "<body>"
act.WriteLine "<table>"
act.WriteLine "<tr>"
act.WriteLine "<td>"
act.WriteLine u_cell1
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine u_cell2
act.WriteLine "</td>"
act.WriteLine "</tr>"
act.WriteLine "</table>"
act.WriteLine "</body>"
act.WriteLine "</html>"
' close the document 
act.close
%>
Your excel has been successfully create and can be viewed by clicking 
<a href="<%= g_filename &"-"& month(date())& day(date())& year(date()) %>.xls" target="_blank">here</a>
<%
' end check of form input
end if
%>
</body>
</html> 
 

