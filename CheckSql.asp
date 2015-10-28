
<!--#include file="include/SQLConn.inc.asp" -->

<%


sql = "SELECT * FROM " & tablename

RS = Conn.Execute(sql)

Do until oRs.EOF
   Response.Write(ucase(fieldname) & ": " & oRs.Fields(fieldname))
   oRS.MoveNext
Loop
oRs.Close


oRs = nothing
oConn = nothing
%>