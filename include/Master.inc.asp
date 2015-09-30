<%
'for each x in Session.Contents
'  Response.Write(x & "=" & Session.Contents(x) & "<br />")
'next

'response.write FormatDateTime(now(),2)
%>

<div id="Header">
<img alt="UOB Kay Hian" align="left" src="images/Logo.gif"/>
</div>


<div id="Title">
<% =Title %> 

</div>

<div id="BlueSeparator">
<% if Session("DBLastModifiedDate") <> "" then %>
	<font size=2>Last updated:<%=Session("DBLastModifiedDate")%></font>
<% end if%>
</div>
   
<div id="Navigation">
<!-- #include file ="cmenu.inc.asp" -->
</div>

<div id="VerticalBlueSeparator">
<p></p>
</div>

<div id="ThinBlueSeparator">
<p></p>
</div>

<div id="Curve">
<img src="images/Curve.gif" width="22" height="16" />
</div>
