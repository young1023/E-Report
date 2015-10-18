<!--
var enabled = 0; today = new Date();
var day; var date;
if(today.getDay()==0) day = "Sunday"
if(today.getDay()==1) day = "Monday"
if(today.getDay()==2) day = "Tuesday"
if(today.getDay()==3) day = "Wednesday."
if(today.getDay()==4) day = "Thursday"
if(today.getDay()==5) day = "Friday"
if(today.getDay()==6) day = "Saturday."
date =(today.getMonth() + 1 ) + "/" + today.getDate() + " " + (today.getYear()) + " " + day;
document.write('<font color="#ffffff">'+date+'</font>');
// -->
