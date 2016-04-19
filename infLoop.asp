<%
On Error resume Next
%>

<HTML><BODY>
<%

Server.ScriptTimeout=60

Dim colErrs, objError, timeout, adoConn, adoCmd
timeout=False
Set adoConn = Server.CreateObject("ADODB.Connection")
'adoConn.CursorLocation = adUseClient
adoConn.Open ("Provider=sqloledb; User ID=intranet; Password=intranet; Initial Catalog=UOBIntranet; Data Source=localhost")
'adoConn.IsolationLevel = adXactReadUncommitted
'The next line starts a transaction
adoConn.BeginTrans
Set adoCmd = Server.CreateObject("ADODB.Command")
'I added this to give the command 45 seconds to execute.
adoCmd.CommandTimeout = 2
adoCmd.ActiveConnection = adoConn
adoCmd.CommandText = "WAITFOR  delay '00:10:01'"
'adoCmd.CommandType = adCmdStoredProc
adoCmd.Execute

Set colErrs=adoConn.Errors


			If adoConn.Errors.Count <> 0 then
			For Each objError In colErrs
			'This is the error number for a timeout.
			If objError.Number=-2147217871 Then
			adoConn.RollbackTrans
			Response.Write "The query timed out before finishing. Please try again."
			timeout=True
			adoConn.Errors.Clear
			Exit For
			End If
			Next
			End If

If Not timeout Then
adoConn.CommitTrans
response.write ("AAA")
End If

Set adoCmd = Nothing

adoConn.Close
Set adoConn = Nothing 
%>

</BODY></HTML>