
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



page_id=request("page_id")

if session("shell_power")="" then
  response.redirect "default.asp"
end if
%>
<% 
Title = "User Group"
%>



<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<SCRIPT language=JavaScript>
<!--
function dosubmit(){
 document.ok.action="UserGroup.asp?sid=<%=SessionID%>";
document.ok.submit();
}
//-->
</SCRIPT>

</head>

<body leftmargin="0" topmargin="0">

<!--#include file="include/Master.inc.asp" -->

<div id="Content">
<%
'-----------------------------------------------------------------------------
'
'      Main Content of the page is inserted here
'
'-----------------------------------------------------------------------------
%>

<DIV align=center>

  <TABLE border=0 cellPadding=0 cellSpacing=0 height=100% width=99%>
    <TBODY> 
    <TR>
      <TD vAlign=top>
        <table width="100%" border="0" cellpadding="1" cellspacing="1" height="100%">
          <tr> 
            <td bgcolor="#000000">
              <table width="100%" border="0" cellpadding=0 cellspacing="0" bgcolor="#FFFFFF" height="100%">
                <tr>
                  <td valign="top" align="center" bgcolor="#E6EBEF">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E6EBEF" class="normal">
                        <tr> 
                          <td valign="top" align="center">
                            <form name=ok method=post>
<table width="99%" border="0" cellspacing="1" bgcolor="#FFFFFF" class="normal">
                              <tr> 
                                <td height="28"> 
                                  <%
		pageid=trim(request.form("pageid"))
		if pageid="" then
		  pageid=1
		end if
        findnum=replace(trim(request.form("findnum")),"%","％")
        findnum=replace(findnum,"'","''")
        
        SortingID = Request("SortingID")
		FilterID = request("FilterID")


            fsql="select * from UserGroup where 1 = 1 "
            
            if findnum <> "" then
            
            fsql = fsql & "and  Name like '%"&findnum&"%' "
            
		    end if
		    
		    If SortingID <> "" Then
		    
		    fsql = fsql & "order by "&SortingID
		     
		    End If
		    
		     If FilterID = 1 Then
        
				fsql  =  fsql  &  " desc" 
	        
		        End If

        set frs=createobject("adodb.recordset")
        

		frs.cursortype=1
		frs.locktype=1
        frs.open fsql,conn

       if frs.RecordCount=0 then
%>
      <font color=red>No Record <a href="javascript:history.go(-1);">[ Return ]</a></font>
          
<%
       else
          findrecord=frs.recordcount
          response.write "Total <font color=red>"&findrecord&"</font> Records ; Total <font color=blue>"

         frs.PageSize = 500
         call countpage(frs.PageCount,pageid)
	     response.write "&nbsp;&nbsp;<input type='text' name='findnum' size='10' value='"&findnum&"' class='common'>"
		 response.write "&nbsp;&nbsp;<input type='button' value='   Search by name  ' onClick='findenum();' class='common'>"
	   end if
	  %>
                                </td>
                              </tr>
                              <tr> 
                                <td valign="top" height="28"> 
                                  <div align="center">
                                  	
                                  	
                                  <table width="99%" border="0" cellspacing="1" class="normal">
                                    <tr> 
                                      <td height="23" width="10%">
										
										<a href="sa_group.asp?SortingID=GroupID&sid=<%=SessionID%>" style="text-decoration: none">
									<img border="0" src="images/up.gif"></a>Group ID												
									<a href="sa_group.asp?SortingID=GroupID&FilterID=1&sid=<%=SessionID%>" style="text-decoration: none">
									<img border="0" src="images/down.gif"></a>

									</td>
                                      <td height="23" width="25%">
										
										<a href="sa_group.asp?SortingID=Name&sid=<%=SessionID%>" style="text-decoration: none">
									<img border="0" src="images/up.gif"></a>Name												
									<a href="sa_group.asp?SortingID=Name&FilterID=1&sid=<%=SessionID%>" style="text-decoration: none">
									<img border="0" src="images/down.gif"></a>
                                    </td>
                                    	<td>
										
										<a href="sa_group.asp?SortingID=Description&sid=<%=SessionID%>" style="text-decoration: none">
									<img border="0" src="images/up.gif"></a>Description												
									<a href="sa_group.asp?SortingID=Description&FilterID=1&sid=<%=SessionID%>" style="text-decoration: none">
									<img border="0" src="images/down.gif"></a>
                                    </td>
                 
                                       <td width="5%">Delete</td>
                                    </tr>
                                    <%
 i=0
 if frs.recordcount>0 then
  frs.AbsolutePage = pageid
  do while (frs.PageSize-i)
   if frs.eof then exit do
   i=i+1

   response.write "<tr bgcolor="&mycolor&">"
   response.write "<td>" & frs("GroupID") & "</td>"
   response.write "<td onmouseover=javascript:style.background='#cccccc' onmouseout=javascript:style.background='"&mycolor&"'>"
%>

<a href="UserGroup.asp?sid=<%=SessionID%>&id=<%=frs("GroupID")%>"><% = frs("Name") %></a>
           </td>
       
     <td width="50%"><% = frs("Description") %></td>

              
 <%
   response.write "<td align=center>"
     response.write "<input type=checkbox name=mid value="&frs("GroupID")&">"
   response.write "</td>"
   response.write "</tr>"
   flage=not flage
   frs.movenext
  loop
 end if
  %>
                                  </table>
                                	</div>
                                </td>
                              </tr>
                              <tr> 
                                <td align="right" height="28"> 
                                  <script language=JavaScript>
<!--
function delcheck(){
k=0;
document.ok.action="delid.asp?sid=<%=SessionID%>"
	if (document.ok.mid!=null)
	{
		for(i=0;i<document.ok.mid.length;i++)
		{
			if(document.ok.mid[i].checked)
			  {
			  k=1;
			  i=1;
			  break;
			  }
		}
		if(i==0)
		{
			if (!document.ok.mid.checked)
               k=0;
			else
               k=1;
		}
	}

if (k==0)
  alert("You must  select one record at least !");	
else if (k==1)
 {
  var msg = "Are you sure ?";
  if (confirm(msg)==true)
   {
    document.ok.whatdo.value="delete group";
    document.ok.submit();
   }
 }

}

function gtpage(what)
{
document.ok.pageid.value=what;
document.ok.action="sa_group.asp?sid=<%=SessionID%>"
document.ok.submit();
}

function findenum()
{
document.ok.pageid.value=1;
document.ok.action="sa_group.asp?sid=<%=SessionID%>"
document.ok.submit();
}
//-->
</script>
                                  <%
	 if frs.recordcount>0 then
             call countpage(frs.PageCount,pageid)
			 response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			 if Clng(pageid)<>1 then
                 response.write " <a href=javascript:gtpage('1') style='cursor:hand' >First</a> "
                 response.write " <a href=javascript:gtpage('"&(pageid-1)&"') style='cursor:hand' >Previous</a> "
			 else
                 response.write " First "
                 response.write " Previous "
			 end if
	         if Clng(pageid)<>Clng(frs.PageCount) then
                 response.write " <a href=javascript:gtpage('"&(pageid+1)&"') style='cursor:hand' >Next</a> "
                 response.write " <a href=javascript:gtpage('"&frs.PageCount&"') style='cursor:hand' >Last</a> "
             else
                 response.write " Next "
                 response.write " Last "
			 end if
	         response.write "&nbsp;&nbsp;"
	 end if
%>
                                </td>
                              </tr>
                              <tr> 
                                <td height="28" align="center"> 
                                  <%
  response.write "<input type='button' value='   New Group   ' onClick='dosubmit();' class='common'>&nbsp;"


  if frs.recordcount>0 then
  response.write "<input type='button' value='   Delete   ' onClick='delcheck();' class='common'>"
  response.write "<input type=hidden value='' name=whatdo>"
  response.write "<input type=hidden value="&pageid&" name=pageid>"
  end if
  response.write "<input type=hidden value='' name='doc_type'>"

			  frs.close
			  set frs=nothing
			  conn.close
			  set conn=nothing
%>
                                </td>
                              </tr>
                              <tr> 
                                <td valign="top"></td>
                              </tr>
                            </table>
                          </form>



</DIV>
<%

  ' function
  Sub countpage(PageCount,pageid)
  response.write pagecount&"</font> Pages "
	   if PageCount>=1 and PageCount<=10 then
		 for i=1 to PageCount
		   if (pageid-i =0) then
              response.write "<font color=green> "&i&"</font> "
		   else
             response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		   end if
		 next
	   elseif PageCount>11 then
	      if pageid<=5 then
		     for i=1 to 10
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       else
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
		     next
		  else
		    for i=(pageid-4) to (pageid+5)
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       elseif i=<pagecount then
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
			next
		  end if
	   end if
  end sub
%>



























<%
'-----------------------------------------------------------------------------
'
'      End of the main Content 
'
'-----------------------------------------------------------------------------
%>
</td>
              </tr>
                </table>
                </td>
                </tr>
              </table>
              <center>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td colspan="2" height="1"></td>
                </tr>
                <tr class="HSEBlue">
                  <td colspan="2" height="1"></td>
                </tr>
              </table>
              </center>
            
              </body>

              </html>