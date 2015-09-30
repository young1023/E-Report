<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" class="Common">
        <tr valign="top" align="center">
          <td class="HSEBlue" height="21"></td>
        </tr>
        <tr valign="top" align="center">
          <td>


	<table align=center border=0 cellpadding=5 cellspacing=0 width=160>
        <tr> 
          <td height=28   class="NaviBar">Statement (User Menu)</td>
        </tr>
       

<%
        if session("shell_power") > 0 then
%>  
        <tr> 
          <td height=8   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="ClientStatement.asp" style="TEXT-DECORATION: none">Client Statement</a></td>
       </tr>
       <tr> 
          <td height=8   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="TransactionHistory.asp" style="TEXT-DECORATION: none">Transaction History</a></td>
        </tr>
        
<tr> 
          <td height=8   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="DetailTrade.asp" style="TEXT-DECORATION: none">Detail Trade</a> </td>
        </tr>
        <tr> 
          <td height=8   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="ClientSummary.asp" style="TEXT-DECORATION: none">Client Summary</a> </td>
        </tr>
        <tr> 
          <td height=8   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="ClientInfo.asp" style="TEXT-DECORATION: none">Client Information</a></td>
        </tr>
        <tr> 
          <td height=8   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="MarginCall.asp" style="TEXT-DECORATION: none">Margin Call</a> </td>
        </tr>
   		<tr> 
          <td height=8   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="ChangePassword.asp" style="TEXT-DECORATION: none">Change Password</a> </td>
        </tr>
        <% end if %>
        <%
        if session("shell_power") > 2 then
        %>
 
 		<tr> 
          <td height=8   class="NaviBar">Administrating Menu </td>
        </tr>

      <tr>
          <td height=28   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="Upload.asp" style="TEXT-DECORATION: none">Upload Margin Value</a></td>
        </tr>

            <tr>
          <td height=28   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="Unlock.asp" style="TEXT-DECORATION: none">Unlock User</a></td>
        </tr>
           <tr>
          <td height=28   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="sa_group.asp" style="TEXT-DECORATION: none">User Group</a></td>
        </tr>
<% end if %>
<%  if session("shell_power") > 5 then %>   
       <tr>
          <td height=28   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="sa_member.asp?id=1" style="TEXT-DECORATION: none">Member</a> </td>
        </tr>
    
        <tr>
          <td height=28   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="Setup.asp" style="TEXT-DECORATION: none">Setup</a></td>
        </tr>
        <tr>
          <td height=28   class="NaviBar">&nbsp;&nbsp;&nbsp;&nbsp;<a href="MenuSetup.asp" style="TEXT-DECORATION: none">Menu Setup</a></td>
        </tr>
        <tr>
          <td height=28   class="NaviBar"><a href="Audit.asp" style="TEXT-DECORATION: none">&nbsp;&nbsp;&nbsp;Audit log</a> </td>
        </tr>
 
<% end if %>
       <tr>
          <td height=28   class="NaviBar"><a href="logout.asp" style="TEXT-DECORATION: none">Logout</a> </td>
        </tr>
       
</table>


  </td>
        </tr>
      </table>