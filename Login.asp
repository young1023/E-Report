<%
' Get Login and Password
num = replace(trim(request.form("num")),"'","''")
inkey = replace(trim(request.form("inkey")),"'","''")
ipaddress  = Request.ServerVariables("REMOTE_ADDR")
' Reset System Message
Session("SystemMessage") = ""
%>
<!--#include file="include/SQLConn.inc.asp" -->

<%
	Session.Timeout=600
	response.expires=0

	' Check if the Login Name Exist
	' --------------------------------------------------------------------------------------------------------------------	
	
	set Rs1 = server.createobject("adodb.recordset")
	Rs1.open ("Exec Checkpwd '"&num&"' , '"&inkey&"', '"&ipaddress&"' ") ,  StrCnn,3,1

       
	If (not Rs1.EoF) Then
			If (cint(rs1("memberid") <0)) then


					
					'Unsucess login attempt
					'**********************
					
					
					'Session.Abandon
					session("shell_power")= 0
					Session("SystemMessage") = rs1("Message")
	
          			Response.Redirect "default.asp"		
			Else
			
					'Sucessful Login
					'***********************
					
					
			        Session("id") = num
			        Session("MemberID") = rs1("MemberID")
			        Session("shell_power") = Rs1("UserLevel")
			        Session("name") = Rs1("Name")					
					Session("SessionID") = rs1("SessionID")
					Session("GroupID") = rs1("GroupID")


					' shared group handler
					set RsSharedGroup = server.createobject("adodb.recordset")
					RsSharedGroup.open ("Check_SharedGroup  '"&num&"' ") ,  StrCnn,3,1

					If (not RsSharedGroup.EoF) Then
						session("SharedGroup") = cint(RsSharedGroup(0))
					end if
					 
				
					
					'Check DB Last modified Date
					set DBDate= Conn.Execute ("Exec Get_SystemSetting 'DBLastModifiedDate' ") 
					If Not DBDate.EoF Then
							Session("DBLastModifiedDate") = DBDate("SettingValue")
					end if

					set DBDateValue= Conn.Execute ("Exec Get_SystemSetting 'DBLastModifiedDateValue' ") 
					If Not DBDate.EoF Then
							Session("DBLastModifiedDateValue") = DBDateValue("SettingValue")
					end if
	
                    ' Check the warning display before certain days
                    set DWValue = Conn.Execute ("Exec Get_SystemSetting 'DisplayWarning' ") 
					If Not DWValue.EoF Then
							DisplayWarning = DWValue("SettingValue")
					end if
					
					if clng(rs1("PasswordAge")) > 0 then

                 'Password not yet expired
                 PasswordAge = rs1("PasswordAge")
                 
                 ' Display Warning Message centain day before expired
                 If (DisplayWarning - PasswordAge) >= 0 Then

                   Warning = "<font color='red'>Password will be expired " & PasswordAge & " days later.</font>"

                 End If

						Session("SystemMessage") = rs1("Message") & "<br><br>" & Warning 

   							Response.Redirect "UserDetail.asp?sid=" & Session("SessionID")
   				else
   							'Password expired
   							Session("SystemMessage") = "Password Expired. Please change your password now"
   							Response.Redirect "PasswordExpired.asp?id=" & Session("ID")
   				end if
			End if
	end if


wait=2




%>
