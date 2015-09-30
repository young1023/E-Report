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



//-->
</SCRIPT>


<script language="JavaScript">
function disableCtrlKeyCombination(e)
{
        //list all CTRL + key combinations you want to disable
        var forbiddenKeya = 'a';
        var forbiddenKeyc = 'c';
        var forbiddenKeyx = 'x';


        var key;
        var isCtrl;

        if(window.event)
        {
                key = window.event.keyCode;     //IE
                if(window.event.ctrlKey)
                        isCtrl = true;
                else
                        isCtrl = false;
        }
        else
        {
                key = e.which;     //firefox
                if(e.ctrlKey)
                        isCtrl = true;
                else
                        isCtrl = false;
        }

        //if ctrl is pressed check if other key is in forbidenKeys array
        if(isCtrl)
        {
            
                {
                        //case-insensitive comparation
                        if(forbiddenKeya.toLowerCase() == String.fromCharCode(key).toLowerCase())
                        {
                                return false;
                        }
                        if(forbiddenKeyc.toLowerCase() == String.fromCharCode(key).toLowerCase())
                        {
                                return false;
                        }

						if(forbiddenKeyx.toLowerCase() == String.fromCharCode(key).toLowerCase())
                        {
                                return false;
                        }

                }
        }
        return true;
}
</script>

<script language="JavaScript">
<!--
// disable right click
var message="Sorry, The right click function is disable."; // Message for the alert box

function click(e) {
if (document.all) {
if (event.button == 2) {
alert(message);
return false;
}
}
if (document.layers) {
if (e.which == 3) {
alert(message);
return false;
}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;
// --> 
</script>
