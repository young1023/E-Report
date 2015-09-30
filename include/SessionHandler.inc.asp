<!--#include file="SQLconn.inc.asp" -->
<%
'Authorisation


'Response.ExpiresAbsolute = #2010-01-01# 
'Response.AddHeader "cache-control", "private, no-cache, must-revalidate" 
'Response.AddHeader "pragma", "no-cache"
'Response.CacheControl = "private, no-cache, must-revalidate"
'Response.Expires = -1

Response.Buffer=true
Response.Expires=-1
Response.ExpiresAbsolute=now()-1
Response.CacheControl="private, no-cache, must-revalidate"
Response.AddHeader "cache-control", "private, no-cache, must-revalidate" 
Response.AddHeader "pragma", "no-cache"

ipaddress  = Request.ServerVariables("REMOTE_ADDR")
SessionID   = Request.QueryString("sid").item



set RsSession = server.createobject("adodb.recordset")
'response.write ("Exec Validate_User '"&Session("id")&"',  '"&Sessionid&"','"&ipaddress&"'") & "<BR>"

RsSession.open ("Exec Validate_User '"&Session("id")&"',  '"&Sessionid&"','"&ipaddress&"'") ,  StrCnn,3,1
'RsSession.open ("Exec Validate_User '"&Session("id")&"',  '"&Session("Sessionid")&"','"&ipaddress&"'") ,  StrCnn,3,1


'Session.Abandon

	If (not RsSession.EoF) Then
			
			' Invalid Login Session
			if (rsSession("Sessionid") = "") then

					' Terminate All session
					'Session.Abandon 	
					'Session("SystemMessage") = "Session Expired. User logged out automatically"
					session("id") = ""
					response.redirect "logout.asp?r=1"
'for each x in session.Contents
' Response.Write(x & "=" & session.Contents(x) & "   ")
'next					
			else		
				'Session("SessionID") = RsSession("SessionID")
			SessionID = RsSession("SessionID")
			'Session("Sessionid") = RsSession("SessionID")
'			response.write ("Exec Validate_User '"&Session("id")&"',  '"&Sessionid&"','"&ipaddress&"'")
'for each x in session.Contents
' Response.Write(x & "=" & session.Contents(x) & "   ")
'next
			end if
	end if
	

%>