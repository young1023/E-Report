<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



Title = "Conversion File Setup"
%>

<%

MemberID = Session("MemberID")

Message = ""


' Modify the Setup
'*******************
if trim(request("action_button")) = "editSetup" then

    MemberID = Request.Form("MemberID")

    TrackingName          = trim(request.form("TrackingName"))

	PartnerCode           = trim(request.form("PartnerCode"))

    PartnerReferenceCode  = trim(request.form("PartnerReferenceCode"))
      
    EstablishCode         = trim(request.form("EstablishCode"))

    Rate                  = trim(request.form("Rate"))

    DepotFolder           = trim(request.form("DepotFolder"))

    FileType              = trim(request.form("FileType"))

    FirstRow              = trim(request.form("FirstRow"))

    Delimiter             = trim(request.form("Delimiter"))

    GivenName             = trim(request.form("GivenName"))

    FamilyName            = trim(request.form("FamilyName"))

    Mile                  = trim(request.form("Mile"))

    ActivityDate          = trim(request.form("ActivityDate"))

    Membership            = trim(request.form("Membership"))


       
       delsql = "Delete from AsiaMileSetup where DepotNo = " & MemberID

       conn.execute delsql 

	
		strsql="Insert into AsiaMileSetup (DepotNo, TrackingName , PartnerCode , PartnerReferenceCode "

        strsql= strsql & ", EstablishmentCode , ExchangeRate , DepotFolder , FileType , FirstRow , Delimiter "

        strsql= strsql & ", GivenName, FamilyName, Mile, ActivityDate, Membership)"

        strsql= strsql & "Values (" & MemberID & ",'" & TrackingName & "', '" & PartnerCode & "','"

        strsql= strsql &  PartnerReferenceCode & "','" & EstablishCode & "', '"  

        strsql= strsql &  Rate & "','" & DepotFolder & "','"  

        strsql= strsql &  Filetype & "','"

        strsql= strsql &  FirstRow & "','"

        strsql= strsql &  Delimiter & "', "

        strsql= strsql &  GivenName & ", " & FamilyName & "," & Mile & "," & ActivityDate 

        strsql= strsql & "," & Membership & ")"
 
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

function doSelect()
{
  	document.fm1.action_button.value="ViewSetup";
		document.fm1.submit();
	
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

    If trim(request("action_button")) = "ViewSetup" then

    MemberID = Request.Form("MemberID")
    
    SQL    = "Select * from AsiaMileSetup where DepotNo = " & MemberID

    Else

    SQL    = "Select * from AsiaMileSetup where DepotNo = " & MemberID

    End if

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

    GivenName             = trim(Rs("GivenName"))

    FamilyName            = trim(Rs("FamilyName"))

    Mile                  = trim(Rs("Mile"))

    ActivityDate          = trim(Rs("ActivityDate"))

    Membership            = trim(Rs("Membership"))

    End if

%>


<form name="fm1" method="post" action="">


<table width="80%" border="0" cellpadding="4" class="normal">


      <% 
           
           Lsql = " Select * from Member where MemberID < > 7 order by LoginName"
           Set LRs = conn.execute(Lsql)
  
%>


	<tr>

      <td height="18">Member</td>
      <td valign="bottom" height="18">
				    	
<select name="MemberID" class="common"  size="1" onchange="doSelect()">
          <% 
                             If Not LRs.EoF Then
                        LRs.MoveFirst
							do while not LRs.eof
                              if trim(MemberID) = trim(LRs("MemberID")) then
                                 response.write "<option value="&LRs("MemberID")&" selected>"&trim(LRs("LoginName"))&"</option>"
                                 else
                                 response.write "<option value="&LRs("MemberID")&">"&trim(LRs("LoginName"))&"</option>"
                               end if
                               LRs.movenext
							loop
						
						End if
					%>
        </select>
				
				
				</td>


    </tr>

    
 <tr> 
      <td width="27%">
Tracking Name</td> 
      <td width="69%">
      <Input name="TrackingName" type=text value="<% = TrackingName %>" size="30"></td>
    </tr>


 <tr> 
      <td width="27%">
Partner Code</td> 
      <td width="69%">
      <Input name="PartnerCode" type=text value="<% = PartnerCode %>" size="30" MaxLength="4"></td>
    </tr>



 <tr> 
      <td width="27%">
Partner Reference Code</td> 
      <td width="69%">
      <Input name="PartnerReferenceCode" type=text value="<% = PartnerReferenceCode %>" MaxLength="10" size="30"></td>
    </tr>



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
Given Name Position</td> 
      <td width="69%">
      <Input name="GivenName" type=text value="<% = GivenName %>" size="30" MaxLength="2">&nbsp;


       </td>
    </tr>

<tr> 
      <td width="27%">
Family Name Position</td> 
      <td width="69%">
      <Input name="FamilyName" type=text value="<% = FamilyName %>" size="30" MaxLength="2">&nbsp;


       </td>
    </tr>

<tr> 
      <td width="27%">
Mile</td> 
      <td width="69%">
      <Input name="Mile" type=text value="<% = Mile %>" size="30" MaxLength="2">&nbsp;


       </td>
    </tr>

<tr> 
      <td width="27%">
Activity Date</td> 
      <td width="69%">
      <Input name="ActivityDate" type=text value="<% = ActivityDate %>" size="30" MaxLength="2">&nbsp;


       </td>
    </tr>

<tr> 
      <td width="27%">
Membership</td> 
      <td width="69%">
      <Input name="Membership" type=text value="<% = Membership %>" size="30" MaxLength="2">&nbsp;


       </td>
    </tr>

 <tr> 
      <td width="27%">
�@</td> 
      <td width="69%">
      	<input type="button" value="    Submit    " onClick="javascript:doSubmit();">
         <input type="hidden" name="action_button" value="">   
�@</td>
    </tr>

<tr> 
      <td  align =center><font color="red"><% = Message %></font></td> 
      <td >
�@</td>
    </tr>

    </table>


</div>
            
              </body>

              </html>