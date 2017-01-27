<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



Title = "Conversion File Setup"
%>

<%


Message = ""


' Modify the Setup
'*******************
if trim(request("action_button")) = "editSetup" then

    TrackingName          = trim(request.form("TrackingName"))

	PartnerCode           = trim(request.form("PartnerCode"))

    PartnerReferenceCode  = trim(request.form("PartnerReferenceCode"))
      
    EstablishCode         = trim(request.form("EstablishCode"))

    Rate                  = trim(request.form("Rate"))

    DepotFolder           = trim(request.form("DepotFolder"))

    FileType              = trim(request.form("FileType"))

    FirstRow              = trim(request.form("FirstRow"))

    Delimiter             = trim(request.form("Delimiter"))

	
		strsql="Update AsiaMileSetup set TrackingName = '" & TrackingName & "', PartnerCode = '"

        strsql= strsql & PartnerCode 

        strsql= strsql & "' , PartnerReferenceCode = '" 

        strsql= strsql & PartnerReferenceCode

        strsql= strsql & "' , EstablishmentCode ='"

        strsql= strsql & EstablishCode 

        strsql= strsql & "' , ExchangeRate = " & Rate 

        strsql= strsql & ", DepotFolder = '" & DepotFolder 

        strsql= strsql & "', FileType = '" & Filetype 

        strsql= strsql & "' , FirstRow = " & FirstRow

        strsql= strsql & " , Delimiter ="& Delimiter 

        strsql= strsql & " where DepotNo = 100"

        

        'response.write strsql
		
	    conn.execute strsql 

        Message =  "Done."


	

	
end if




%>

<!--#include file="include/SQLconn.inc.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<SCRIPT language=JavaScript>
<!--

function doSubmit()
{
       

	if(document.fm1.TrackingName.value =="")
       {
		alert("Please enter the TrackingName!");
        document.fm1.TrackingName.focus();
        return false;
       }

	if(document.fm1.PartnerCode.value =="")
       {
		alert("Please enter the Partner  Code!");
        document.fm1.PartnerCode.focus();
        return false;
       }


	if(document.fm1.PartnerReferenceCode.value =="")
       {
		alert("Please enter the Partner Reference Code!");
        document.fm1.PartnerReferenceCode.focus();
        return false;
       }


    if(document.fm1.EstablishCode.value =="")
       {
		alert("Please enter Establishment Code!");
        document.fm1.EstablishCode.focus();
        return false;
       }

    if(document.fm1.Rate.value =="")
       {
		alert("Please enter Exchange Rate!");
        document.fm1.Rate.focus();
        return false;
       }

    if(document.fm1.DepotFolder.value =="")
       {
		alert("Please enter upload folder!");
        document.fm1.DepotFolder.focus();
        return false;
       }

    
	else
		{
		document.fm1.action_button.value="editSetup";
		document.fm1.submit();
		}
}




//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">
<%
'-----------------------------------------------------------------------------
'
'      Main Content of the page is inserted here
'
'-----------------------------------------------------------------------------


    SQL    = "Select * from AsiaMileSetup where DepotNo = 100"

    Set Rs = Conn.Execute(SQL)

    If Not Rs.Eof then

    PartnerCode          = trim(Rs("PartnerCode"))

    PartnerReferenceCode = trim(Rs("PartnerReferenceCode"))

    TrackingName         = trim(Rs("TrackingName"))

    EstablishmentCode    = trim(Rs("EstablishmentCode"))

    Rate                 = trim(Rs("ExchangeRate"))

    DepotFolder          = trim(Rs("DepotFolder"))

    FileType             = trim(Rs("Filetype"))

    FirstRow             = trim(Rs("FirstRow"))

    Delimiter            = trim(Rs("Delimiter"))

    End if

%>


<form name="fm1" method="post" action="">


<table width="80%" border="0" cellpadding="4" class="normal">



 <tr> 
    
 <tr> 
      <td width="27%">
Tracking Name</td> 
      <td width="69%">
      <Input name="TrackingName" type=text value="<% = TrackingName %>" size="30"></td>
    </tr>

 <tr> 

 <tr> 
      <td width="27%">
Partner Code</td> 
      <td width="69%">
      <Input name="PartnerCode" type=text value="<% = PartnerCode %>" size="30" MaxLength="4"></td>
    </tr>

 <tr> 

 <tr> 
      <td width="27%">
Partner Reference Code</td> 
      <td width="69%">
      <Input name="PartnerReferenceCode" type=text value="<% = PartnerReferenceCode %>" MaxLength="10" size="30"></td>
    </tr>

 <tr> 

 <tr> 
      <td width="27%">
Establishment Code</td> 
      <td width="69%">
      <Input name="EstablishCode" type=text value="<% = EstablishmentCode %>" size="30" MaxLength="10"></td>
    </tr>

 <tr> 

 <tr> 
      <td width="27%">
Asia Mile Exchange Rate</td> 
      <td width="69%">
      <Input name="Rate" type=text value="<% = Rate %>" size="30" MaxLength="4"></td>
    </tr>

 <tr> 



 <tr> 
      <td width="27%">
Upload Folder</td> 
      <td width="69%">
      <Input name="DepotFolder" type=text value="<% = DepotFolder %>" size="30">&nbsp;
       </td>
    </tr>

<tr> 
      <td width="27%">
File Extension</td> 
      <td width="69%">
      <Input name="FileType" type=text value="<% = FileType %>" size="30" MaxLength="4">&nbsp;


       </td>
</tr>

<tr> 
      <td width="27%">
First row of data</td> 
      <td width="69%">
      <Input name="FirstRow" type=text value="<% = FirstRow %>" size="30" MaxLength="2">&nbsp;


       </td>
    </tr>

<tr> 
      <td width="27%">
Delimiter</td> 
      <td width="69%">
      	<select size="1" name="Delimiter" class="common">
		
		  <option value="0" <% If Delimiter=0 Then %> Selected <% End If %>>Comma</option>
			
		  <option value="1" <% If Delimiter=1 Then %> Selected <% End If %>>|</option>
	
	      <option value="2" <% If Delimiter=2 Then %> Selected <% End If %>>Tab</option>

          <option value="3" <% If Delimiter=3 Then %> Selected <% End If %>>Fixed Width</option>

          <option value="4" <% If Delimiter=4 Then %> Selected <% End If %>>Semicolon</option>
	
	    </select>

       </td>
    </tr>
 
 <tr> 
      <td width="27%">
¡@</td> 
      <td width="69%">
      	<input type="button" value="    Submit    " onClick="javascript:doSubmit();">
         <input type="hidden" name="action_button" value="">   
¡@</td>
    </tr>

<tr> 
      <td  align =center><font color="red"><% = Message %></font></td> 
      <td >
¡@</td>
    </tr>

    </table>


</div>
            
              </body>

              </html>