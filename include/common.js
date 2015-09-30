<SCRIPT language=JavaScript>
<!--

function disableSelection(target){
if (typeof target.onselectstart!="undefined") //IE route
	target.onselectstart=function(){return false}
else if (typeof target.style.MozUserSelect!="undefined") //Firefox route
	target.style.MozUserSelect="none"
else //All other route (ie: Opera)
	target.onmousedown=function(){return false}
target.style.cursor = "default"
}



function PopupHelpWindow() {
		newwindow=window.open( "HelpContent.asp?Title=<%=Title%>&sid=<%=SessionID%>", "myWindow", 
									"left=300, top=200, status = 0, scrollbars=1, toolbar=0, height = 400, width = 600, resizable = 1'"  )
		 if (window.focus) {
           newwindow.focus();
       }
 			
}



function PopupWindow() {
		newwindow=window.open( "SearchClientNumber.asp?sid=<%=SessionID%>", "myWindow", 
									"status = 1, height = 300, width = 800, resizable = 1'"  )
		 if (window.focus) {
           newwindow.focus();
       }
 			
}
function PopupSearchAE() {
		newwindow=window.open( "SearchAE.asp?sid=<%=SessionID%>", "myWindow", 
									"status = 1, height = 300, width = 800, resizable = 1'"  )
		 if (window.focus) {
           newwindow.focus();
       }
 			
}





sub OrderVariable(iorder)
'response.write iorder & "," & Search_Order 
  'User click the same field
	if iorder = Search_Order then 
		'reverse the direction
		if Search_Direction = "ASC" then
			response.write "'" & iorder & "','DESC'"
		else
			response.write "'" & iorder & "','ASC'"
		end if
	else
		'User click a different field, default is ascending order
		response.write "'" & iorder & "','ASC'"
	end if 

End sub 

sub OrderImage(iorder)
  'User click the same field
	if iorder = Search_Order then 
		'reverse the direction
		if Search_Direction = "ASC" then
			response.write "<img border=0 src='images/up.jpg'>" 
		else
			response.write "<img border=0 src='images/down.jpg'>" 
		end if
	else
		'User click a different field, default is ascending order
		' do nothing
	end if 
		
end sub



//-->
</SCRIPT>



