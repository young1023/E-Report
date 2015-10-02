
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



Title = "System Setup"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<SCRIPT language=JavaScript>
<!--
function doSubmit(){
document.fm1.action="<%= strURL %>?sid=<%=SessionID%>";

    if (document.fm1.SystemIdleTimeout.value == "") {
            alert("Please enter idle timeout value");
            document.fm1.SystemIdleTimeout.focus();
            return false;
        }
 
     if (document.fm1.PasswordMaximumAge.value == "") {
            alert("Please enter the password expired days.");
            document.fm1.PasswordMaximumAge.focus();
            return false;
        }
        
    if (document.fm1.PasswordMinimumLength.value == "") {
            alert("Please enter the minimum password length.");
            document.fm1.PasswordMinimumLength.focus();
            return false;
        }
        
   if (document.fm1.EmailSMTPServer.value == "") {
            alert("Please enter the SMTP Server.");
            document.fm1.EmailSMTPServer.focus();
            return false;
        }

   if (document.fm1.EmailSenderAddress.value == "") {
            alert("Please enter Sender address.");
            document.fm1.EmailSenderAddress.focus();
            return false;
        }        
document.fm1.submit();
}
//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" OnLoad="document.fm1.SystemIdleTimeout.focus();">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">
<%
'-----------------------------------------------------------------------------
'
'      Main Content of the page is inserted here
'
'-----------------------------------------------------------------------------

	Dim SystemIdleTimeout
	Dim PasswordMinimumLength
	Dim PasswordMaximumAge
	Dim EmailSMTPServer
	Dim EmailSenderName
	Dim EmailSenderAddress
	 
  SystemIdleTimeout        = Request.form("SystemIdleTimeout")
  SystemIdleTimeout2       = Request.form("SystemIdleTimeout2")
  
  PasswordMinimumLength    = Request.form("PasswordMinimumLength")
  PasswordMinimumLength2   = Request.form("PasswordMinimumLength2")
  
  PasswordMaximumAge       = Request.form("PasswordMaximumAge")
  PasswordMaximumAge2      = Request.form("PasswordMaximumAge2")

  PasswordMaximumReuse     = Request.Form("PasswordMaximumReuse")
  PasswordMaximumReuse2    = Request.Form("PasswordMaximumReuse")

  DisplayWarning           = Request.Form("DisplayWarning")
  LockAccount              = Request.Form("LockAccount")
  
  EmailSMTPServer          = Request.form("EmailSMTPServer")
  EmailSMTPServer2         = Request.form("EmailSMTPServer2")
  
  EmailSenderName          = Request.form("EmailSenderName")
  EmailSenderName2         = Request.form("EmailSenderName2")
    
  EmailSenderAddress       = Request.form("EmailSenderAddress")
  EmailSenderAddress2      = Request.form("EmailSenderAddress2")
  
  EmailRecipientAddress    = Request.form("EmailRecipientAddress")
  EmailRecipientAddress2   = Request.form("EmailRecipientAddress2")
  
  BackupFailureEmailBody   = Request.form("BackupFailureEmailBody")
  NoChangeEmailBody        = Request.form("NoChangeEmailBody")
  NoChangeEmailBody2       = Request.form("NoChangeEmailBody2")
   
  GBMailPath               = Request.form("GBMailPath")
  DBbackupPath             = Request.form("DBbackupPath")
  DBbackupflagPath         = Request.form("DBbackupflagPath")
  DBbackupWaitTime         = Request.form("DBbackupWaitTime")
  DBBackupNumberOfRetry    = Request.form("DBBackupNumberOfRetry")
  PwdCapital			   = Request.form("PwdCapital")
  If PwdCapital = "" Then
     PwdCapital = 1
  End If
  PwdNumber				   = Request.form("PwdNumber")
  If PwdNumber = "" Then
     PwdNumber = 1
  End If
  PwdEnglish			   = Request.form("PwdEnglish")
  If PwdEnglish = "" Then
     PwdEnglish = 1
  End If

	
	if SystemIdleTimeout = "" then
					'first visit of page
					
					set rs1= Conn.Execute ("Exec Get_SystemSetting 'SystemIdleTimeout' ") 
					If Not Rs1.EoF Then
							SystemIdleTimeout = rs1("SettingValue")
					end if
					
					set rs2= Conn.Execute ("Exec Get_SystemSetting 'PasswordMinimumLength' ") 
					If Not Rs2.EoF Then
							PasswordMinimumLength = rs2("SettingValue")
					end if
					
					set rs3= Conn.Execute ("Exec Get_SystemSetting 'PasswordMaximumAge' ") 
					If Not Rs3.EoF Then
							PasswordMaximumAge = rs3("SettingValue")
					end if
					
					set rs4= Conn.Execute ("Exec Get_SystemSetting 'EmailSMTPServer' ") 
					If Not Rs4.EoF Then
							EmailSMTPServer = rs4("SettingValue")
					end if
					

							
						set rs6= Conn.Execute ("Exec Get_SystemSetting 'EmailSenderAddress' ") 
					If Not Rs6.EoF Then
							EmailSenderAddress = rs6("SettingValue")
					end if

						set rs7= Conn.Execute ("Exec Get_SystemSetting 'EmailRecipientAddress' ") 
					If Not Rs7.EoF Then
							EmailRecipientAddress = rs7("SettingValue")
					end if

					set rs8= Conn.Execute ("Exec Get_SystemSetting 'NoChangeEmailBody' ") 
					If Not Rs8.EoF Then
							NoChangeEmailBody = rs8("SettingValue")
					end if
	
					set rs9= Conn.Execute ("Exec Get_SystemSetting 'GBMailPath' ") 
					If Not Rs9.EoF Then
							GBMailPath = rs9("SettingValue")
					end if


					set rs10= Conn.Execute ("Exec Get_SystemSetting 'DBbackupPath' ") 
					If Not Rs10.EoF Then
							DBbackupPath = rs10("SettingValue")
					end if


					set rs11= Conn.Execute ("Exec Get_SystemSetting 'DBbackupflagPath' ") 
					If Not Rs11.EoF Then
							DBbackupflagPath = rs11("SettingValue")
					end if
					
					set rs12= Conn.Execute ("Exec Get_SystemSetting 'BackupFailureEmailBody' ") 
					If Not Rs12.EoF Then
							BackupFailureEmailBody = rs12("SettingValue")
					end if
	
					set rs13= Conn.Execute ("Exec Get_SystemSetting 'DBBackupNumberOfRetry' ") 
					If Not Rs13.EoF Then
							DBBackupNumberOfRetry = rs13("SettingValue")
					end if
	
						set rs14= Conn.Execute ("Exec Get_SystemSetting 'DBBackupWaitTime' ") 
					If Not Rs14.EoF Then
							DBBackupWaitTime = rs14("SettingValue")
					end if

	                    set rs15= Conn.Execute ("Exec Get_SystemSetting 'PasswordMaximumReuse' ") 
					If Not Rs15.EoF Then
							PasswordMaximumReuse = rs15("SettingValue")
					end if

                        set rs16= Conn.Execute ("Exec Get_SystemSetting 'DisplayWarning' ") 
					If Not Rs16.EoF Then
							DisplayWarning = rs16("SettingValue")
					end if

                        set rs17= Conn.Execute ("Exec Get_SystemSetting 'LockAccount' ") 
					If Not Rs17.EoF Then
							LockAccount = rs17("SettingValue")
					end if
					
					  set rs18= Conn.Execute ("Exec Get_SystemSetting 'PwdCapital' ") 
					If Not Rs18.EoF Then
							PwdCapital = rs18("SettingValue")
					end if

					    set rs19= Conn.Execute ("Exec Get_SystemSetting 'PwdEnglish' ") 
					If Not Rs19.EoF Then
							PwdEnglish = rs19("SettingValue")
					end if
			
					set rs20= Conn.Execute ("Exec Get_SystemSetting 'PwdNumber' ") 
					If Not Rs20.EoF Then
							PwdNumber = rs20("SettingValue")
					end if


					Session("SystemMessage") = ""
	Else
	
	 MemberID = Session("MemberID")
	
	' Update database
	
		If SystemIdleTimeout <> SystemIdleTimeout2 Then
				
					set rs1= Conn.Execute ("Exec Update_SystemSetting 'SystemIdleTimeout', '"&SystemIdleTimeout&"' ")

				  ' Write system log
				  Message = "Update System Idle Timeout from <b>" & SystemIdleTimeout2 & "</b> to <b>" & SystemIdleTimeout & "</b> at " & Request.ServerVariables("REMOTE_ADDR")
				  
                                 Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"")
                                 
		End If 

		If PasswordMinimumLength <> PasswordMinimumLength2 Then
					 
					set rs2= Conn.Execute ("Exec Update_SystemSetting 'PasswordMinimumLength', '"&PasswordMinimumLength&"' ") 
					
				' Write system log
				 Message = "Update Minimum Password Length from <b>" & PasswordMinimumLength2 & "</b> to <b>" & PasswordMinimumLength & "</b> at " & Request.ServerVariables("REMOTE_ADDR")
				  
                    Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
                    
       End If
                    
       If PasswordMaximumAge <> PasswordMaximumAge2 Then
                                 
					set rs3= Conn.Execute ("Exec Update_SystemSetting 'PasswordMaximumAge', '"&PasswordMaximumAge&"' ") 

				' Write system log
				Message = "Update Password expired from <b>" & PasswordMaximumAge2 & "</b> to <b>" & PasswordMaximumAge & "</b> at " & Request.ServerVariables("REMOTE_ADDR")
				  
                    Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 
                    
       End If

	     If EmailSMTPServer <> EmailSMTPServer2 Then

					set rs4= Conn.Execute ("Exec Update_SystemSetting 'EmailSMTPServer', '"&EmailSMTPServer&"' ") 

				' Write system log
				Message = "Update SMTP Server from <b>" & EmailSMTPServer2 & "</b> to <b>" & EmailSMTPServer & "</b> at " & Request.ServerVariables("REMOTE_ADDR")
				  
                    Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 

	    End If
	    
	  If EmailSenderAddress <> EmailSenderAddress2 Then

					set rs6= Conn.Execute ("Exec Update_SystemSetting 'EmailSenderAddress', '"&EmailSenderAddress&"' ")
					
				' Write system log
				Message = "Update Sender Address from <b>" & EmailSenderAddress2 & "</b> to <b>" & EmailSenderAddress & "</b> at " & Request.ServerVariables("REMOTE_ADDR")
				  
                    Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 

	    End If

		If EmailRecipientAddress <> EmailRecipientAddress2 Then
 
					set rs7= Conn.Execute ("Exec Update_SystemSetting 'EmailRecipientAddress', '"&EmailRecipientAddress&"' ") 
					
					' Write system log
				Message = "Update Recipient Address from <b>" & EmailRecipientAddress2 & "</b> to <b>" & EmailRecipientAddress & "</b> at " & Request.ServerVariables("REMOTE_ADDR")
				  
                    Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 

		 End If

		If NoChangeEmailBody <> NoChangeEmailBody2 Then
					
					set rs8= Conn.Execute ("Exec Update_SystemSetting 'NoChangeEmailBody', '"&NoChangeEmailBody&"' ") 
					
					' Write system log
				Message = "Update Email Body if Content Not Change from <b>" & NoChangeEmailBody2 & "</b> to <b>" & NoChangeEmailBody & "</b> at " & Request.ServerVariables("REMOTE_ADDR")
				  
                    Conn.Execute ("Exec AddLog '"&Message&"', "&MemberID&"") 

					
		End If			
					
					set rs9= Conn.Execute ("Exec Update_SystemSetting 'GBMailPath', '"&GBMailPath&"' ") 
					set rs10= Conn.Execute ("Exec Update_SystemSetting 'DBbackupPath', '"&DBbackupPath&"' ") 
					set rs11= Conn.Execute ("Exec Update_SystemSetting 'DBbackupflagPath', '"&DBbackupflagPath&"' ") 
					set rs12= Conn.Execute ("Exec Update_SystemSetting 'BackupFailureEmailBody', '"&BackupFailureEmailBody&"' ") 
					set rs13= Conn.Execute ("Exec Update_SystemSetting 'DBBackupNumberOfRetry', '"&DBBackupNumberOfRetry&"' ") 
					set rs14= Conn.Execute ("Exec Update_SystemSetting 'DBBackupWaitTime', '"&DBBackupWaitTime&"' ") 
					set rs15= Conn.Execute ("Exec Update_SystemSetting 'PasswordMaximumReuse', '"&PasswordMaximumReuse&"' ") 
					set rs16= Conn.Execute ("Exec Update_SystemSetting 'DisplayWarning', '"&DisplayWarning&"' ") 
					set rs17= Conn.Execute ("Exec Update_SystemSetting 'LockAccount', '"&LockAccount&"' ")
					set rs18= Conn.Execute ("Exec Update_SystemSetting 'PwdCapital', '"&PwdCapital&"' ") 
					set rs19= Conn.Execute ("Exec Update_SystemSetting 'PwdEnglish', '"&PwdEnglish&"' ") 
					set rs20= Conn.Execute ("Exec Update_SystemSetting 'PwdNumber', '"&PwdNumber&"' ") 
 
				
					Session("SystemMessage") = "System settings updated"

'response.write ("Exec Update_SystemSetting 'SystemIdleTimeout', '"&SystemIdleTimeout&"' ")
	
	
	
	
	
	
	End if	
	

%>




<form name="fm1" method="post" action="">
<table width="99%" border="0" class="normal">
  <tr> 
      <td colspan="2"  align="right">
<font color="red">*</font>	 Denotes a mandatory field</td>
    </tr>
  <tr> 
      <td colspan="2"  align="right">
　</td>
    </tr>
    
 <tr> 
      <td width="43%">System Idle Timeout</td> 
      <td width="55%">
      	     
<input name="SystemIdleTimeout" type=text value="<% = SystemIdleTimeout %>" size="5"> 
Seconds</td>
    </tr>
<input name="SystemIdleTimeout2" type=hidden value="<% = SystemIdleTimeout %>" size="5"> 
 <tr> 
      <td width="43%">　</td> 
      <td width="55%">
      	     
　</td>
    </tr>
     <tr> 
      <td width="43%"><b>Password Policy:</b></td> 
      <td width="55%">
　</td>
    </tr>

 <tr> 
      <td width="43%">Password will be expired after</td> 
      <td width="55%">
      	     
<input name="PasswordMaximumAge" type=text value="<% = PasswordMaximumAge %>" size="5"> 
Days</td>
    </tr>
<input name="PasswordMaximumAge2" type=hidden value="<% = PasswordMaximumAge %>" size="5"> 
     <tr> 
      <td width="43%">Cannot reuse last</td> 
      <td width="55%"><input name="PasswordMaximumReuse" type=text value="<% = PasswordMaximumReuse %>" size="5">
 old passwords 
<input name="PasswordMaximumReuse2" type=hidden value="<% = PasswordMaximumReuse %>" size="5"> 
      　</td>
    </tr>
     <tr> 
      <td width="43%">Minimum Password Length</td> 
      <td width="55%">

<input name="PasswordMinimumLength" type=text value="<% = PasswordMinimumLength %>" size="5">&nbsp; 
Characters</td>
    </tr>
<input name="PasswordMinimumLength2" type=hidden value="<% = PasswordMinimumLength %>" size="5">    
         
     <tr> 
      <td width="43%">Lock Account after</td> 
      <td width="55%"><input name="LockAccount" type=text value="<% = LockAccount%>" size="5">
 times fail attemp
<input name="LockAccount2" type=hidden value="<% = LockAccount %>" size="5"> 
      　</td>
    </tr>
    
         <tr> 
      <td width="43%">Password Combination</td> 
      <td width="55%">Capital Letter&nbsp;
      <input type = checkbox name="PwdCapital" value="0" <% If PwdCapital = "0" Then %>Checked<% End If %>>&nbsp;
      English Letter <input type = checkbox name="PwdEnglish" value="0" <% If PwdEnglish = "0" Then %>Checked<% End If %>> 
		Number 
<input type = checkbox name="PwdNumber" value="0" <% If PwdNumber = "0" Then %>Checked<% End If %>>　</td>
    </tr>
    
     <tr> 
      <td width="43%">Display warning</td> 
      <td width="55%"><input name="DisplayWarning" type=text value="<% = DisplayWarning %>" size="5">
 day before password expiry
<input name="DisplayWarning2" type=hidden value="<% = DisplayWarning %>" size="5"> 
      　</td>
    </tr>

<tr> 
      <td width="43%">　</td> 
      <td width="55%">

      　</td>
    </tr>
     <tr> 
      <td width="43%"><b>Mail Server setting:</b></td> 
      <td width="55%">
　</td>
    </tr>


	<tr> 
			<td width="43%">SMTP Server</td> 
			<td width="55%">
			 
			<input name="EmailSMTPServer" type=text value="<% = EmailSMTPServer %>" size="36"></td>
	</tr>
<input name="EmailSMTPServer2" type=hidden value="<% = EmailSMTPServer %>" size="36">	
	<tr> 
			<td width="43%">Sender Email Address</td> 
			<td width="55%">
			
			<input name="EmailSenderAddress" type=text value="<% = EmailSenderAddress %>" size="36"></td>
	</tr>
			
<input name="EmailSenderAddress2" type=hidden value="<% = EmailSenderAddress %>" size="36">	
	<tr> 
			<td width="43%">Recipient Email Address (use , as a seperator if multiple recipients)</td> 
			<td width="55%">
			
			<input name="EmailRecipientAddress" type=text value="<% = EmailRecipientAddress %>" size="36"></td>
	</tr>
<input name="EmailRecipientAddress2" type=hidden value="<% = EmailRecipientAddress %>" size="36">
	<tr> 
			<td width="43%">File path of mail executable</td> 
			<td width="55%">
			
			<input name="GBMailPath" type=text value="<% = GBMailPath %>" size="36"></td>
	</tr>
<input name="GBMailPath2" type=hidden value="<% = GBMailPath %>" size="36">	
     <tr> 
      <td width="43%">　</td> 
      <td width="55%">

      　</td>
    </tr>
     <tr> 
      <td width="43%"><b>Database setting:</b></td> 
      <td width="55%">
　</td>
    </tr>
    
	<tr> 
			<td width="43%">File path and filename of DB backup</td> 
			<td width="55%">
			
			<input name="DBbackupPath" type=text value="<% = DBbackupPath %>" size="36"></td>
	</tr>
<input name="DBbackupPath2" type=hidden value="<% = DBbackupPath %>" size="36">	
		<tr> 
			<td width="43%">File path and filename of DB backup flag file</td> 
			<td width="55%">
			
			<input name="DBbackupflagPath" type=text value="<% = DBbackupflagPath %>" size="36"></td>
	</tr>
			<tr> 
			<td width="43%">Numbers of retry for locating backup and flag files</td> 
			<td width="55%">
			
			<input name="DBbackupNumberOfRetry" type=text value="<% = DBbackupNumberOfRetry %>" size="36"></td>
	</tr>
		<tr> 
			<td width="43%">Time to be wait between each retry (in HH:MM:SS format)</td> 
			<td width="55%">
			
			<input name="DBBackupWaitTime" type=text value="<% = DBBackupWaitTime %>" size="36"></td>
	</tr>
	


     <tr> 
      <td width="43%">　</td> 
      <td width="55%">

      　</td>
    </tr> 
 
     <tr> 
      <td width="43%"><b>Email notification when DB backup files are not found:</b></td> 
      <td width="55%">
　</td>
    </tr>
	<tr> 
			<td width="43%">Email Body (not exceeding 1024 characters and no quotation marks)</td> 
			<td width="55%">
			
			<textarea name="BackupFailureEmailBody"  rows=3  cols="36"><% = BackupFailureEmailBody %></textarea>
	</tr>
	

 
     <tr> 
      <td width="43%">　</td> 
      <td width="55%">

      　</td>
    </tr> 
 
     <tr> 
      <td width="43%"><b>Email notification when DB content has not been changed:</b></td> 
      <td width="55%">
　</td>
    </tr>
	<tr> 
			<td width="43%">Email Body (not exceeding 1024 characters and no quotation marks)</td> 
			<td width="55%">
			
			<textarea name="NoChangeEmailBody"  rows=3  cols="36"><% = NoChangeEmailBody %></textarea>
	</tr>
<input name="NoChangeEmailBody2" type=hidden value="<% = NoChangeEmailBody %>" size="5000">	

     <tr> 
      <td width="43%">　</td> 
      <td width="55%">
　</td>
    </tr>

     <tr> 
      <td width="43%"><b>Stock Reconciliation Setup:</b></td> 
      <td width="55%">
　</td>
    </tr>
<% ArchiveFolder = "E:\Data\Recon\Archive\" %>
     <tr> 
      <td width="43%">Archive Folder:</td> 
      <td width="55%"><input name="ArchiveFoler" type=text value="<% = ArchiveFolder %>" size="36">
　</td>
    </tr>
	
<input name="NoChangeEmailBody2" type=hidden value="<% = NoChangeEmailBody %>" size="5000">	

     <tr> 
      <td width="43%">　</td> 
      <td width="55%">
　</td>
    </tr>
<tr> 
<td colspan="2" align="center"> 

<input type="button" value="    Submit  " onClick="doSubmit();" class="Normal">
<input type="hidden" value="" name="whatdo">
</td>
</tr>

                      </tr>
                          <tr align="center"> 
                    <td colspan="2" height="28" class="RedClr"><% = Session("SystemMessage") %>
                    </td>
                  </tr>
</table>
</form>

<%


'-----------------------------------------------------------------------------
'
'      End of the main Content 
'
'-----------------------------------------------------------------------------
%>
</div>
            
              </body>

              </html>
              
<%Session("SystemMessage") = ""%>