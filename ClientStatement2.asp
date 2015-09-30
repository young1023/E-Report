
<% 
'*********************************************************************************
'NAME       : ClientStatement2.asp           
'DESCRIPTION: Client Statement Friendly Print Page
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 090401 Gary Yeung   Prototype
'MODIFIED   : 090403 Roger Wong   Record and page control
'							090712 Roger Wong		Add Shared Group
'********************************************************************************

' Section Order	Section Code									        	
' --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'      1		         CN		Contract Note / Trade Summary 		(Cash & Custodian)        	
'      4		         TM		Contract Note / Trade Summary 		(Margin)		           	
'      2		         CC		Sub-total of Each Trade					           	
'      2		         CL		Total Amount of All Trade				           	
'      4		         CM		Cash Movement							
'      6		         SM		Securities Movement						
'      4		         BM		Cash Opening Balance						
'      5		         CB		Cash Closing Balance						
'      7		         UT		UnSettled and Pending Trade					
'      8		         TL		Total Value of UnSettled and Pending Trade					
'      9		         MP		Stock Portfolio 				(Margin)			
'      10		         ML		Total Market Value of Stock Portfolio	(Margin)			
'      11		         SP		Stock Portfolio 				(Cash & Custodian)	
'      12		         SL		Total Market Value of Stock Portfolio	(Cash & Custodian)	
'      13 		       MS		Summary 				(Margin)				   	
'      14		         NS		Summary 				(Cash & Custodian)		   	
%>

<%
On Error resume Next
%>

<!--#include file="include/SessionHandler.inc.asp" -->
<%

Dim MSG_BUSY
MSG_BUSY = "The selection criteria are too broad for the system to process. Please use more specific selection criteria then resubmit."

if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if
%>

<%
Title = "Client Statement"

' **************
' Authorisation
' **************
if session("shell_power")="" then
  response.redirect "logout.asp?r=-1"
end if



'**************
'Initialisation
'**************


Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adCmdText = &H0001

Const PAGE_SIZE = 5  ' The size of our pages.

Dim iPageCurrent ' The page we're currently on
Dim iPageCount   ' Number of pages of records
Dim iRecordCount ' Count of the records returned
Dim I            ' Standard looping variable


Dim cnnSearch  ' ADO connection
Dim rstSearch  ' ADO recordset




strURL = Request.ServerVariables("URL") ' Retreive the URL of this page from Server Variables
%>






<html>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css">
</head>
<body leftmargin="0" topmargin="0">


<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=-1"
end if
%>

         

<%




'**************
'Argument handler
'**************

Dim Search_GroupID
Dim Search_ClientFrom
Dim Search_ClientTo
Dim Search_AEFrom
Dim Search_AETo
Dim Search_AEGroup
Dim Search_Statement
Dim Search_Monthly_Month
Dim Search_Monthly_Year
Dim Search_Daily_Day
Dim Search_Daily_Month
Dim Search_Daily_Year
Dim Search_SharedSelection 
Dim Search_SharedGroup
Dim Search_SharedGroupMember
Dim Search_NetValue

Search_AEGroup	    = Request("GroupID")
Search_ClientFrom   = Request("ClientFrom")
Search_ClientTo     = Request("ClientTo")
Search_AEFrom       = Request("AEFrom")
Search_AETo         = Request("AETo")
Search_Statement    = Request("Statement")
Search_Monthly_Month= Request("SMMonth")
Search_Monthly_Year = Request("SMYear")
Search_Daily_Day    = Request("SDay")
Search_Daily_Month  = Request("SMonth")
Search_Daily_Year   = Request("SYear")
Search_NetValue = Request("NetValue")	

Search_SharedSelection      =  Request("ShareSelection")	
Search_SharedGroup  = Request("SharedGroup")
Search_SharedGroupMember  = Request("SharedGroupMember")


'AECode search permission
Select Case Session("shell_power")
	case "1"
		'AE shall access their own clients only
		Search_AEFrom = Session("id")
		Search_AETo = Session("id")
		Search_AEGroup = Session("GroupID")
		
  case "5"
		' Branch Manager shall access all AE's clients belongs to 
		Search_AEGroup = Session("GroupID")
		
		'Other having full access
end select

'Shared group handler


' If User enter From value only, change the "To" value to "From"
if Search_ClientTo = "" then
   Search_ClientTo = Search_ClientFrom
end if

	iPageCurrent = Clng(Request("page"))


session("GroupID")               =  Search_AEGroup	               
session("ClientFrom")            =  Search_ClientFrom              
session("ClientTo")              =  Search_ClientTo                
session("AEFrom")                =  Search_AEFrom                  
session("AETo")                  =  Search_AETo                    


set RsGroupID = server.createobject("adodb.recordset")
RsGroupID.open ("Exec Retrieve_AvailableGroupID ") ,  StrCnn,3,1


'*****************************************************************
' Start of report body
'*****************************************************************
	
	
	set Rs1 = server.createobject("adodb.recordset")
	

 'Rs return 2 value
 '1) Total number of matched client
 '2) all records for targeted client

	
response.write ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '"&session("shell_power")&"', '"&Search_SharedGroup&"','"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&iPageCurrent&"', '1' ")
 
	Select Case  Search_SharedSelection
	case "share2"
		'shared group member

	'		response.write ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&iPageCurrent&"', '1' ")
			Rs1.open ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_SharedGroupMember&"', '"&Search_SharedGroupMember&"', '', '"&session("shell_power")&"', '', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_NetValue&"','"&iPageCurrent&"', '1' ") ,  StrCnn,3,1

	case "share3"

		'	response.write ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"','"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&iPageCurrent&"', '1' ")
			
			Rs1.open ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"',  '', '', '', '"&session("shell_power")&"', '"&Search_SharedGroup&"','"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_NetValue&"','"&iPageCurrent&"', '1' ") ,  StrCnn,3,1
	case else
		'normal
			'response.write ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_NetValue&"','"&iPageCurrent&"', '1' ")
			Rs1.open ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '', '', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&Search_NetValue&"','"&iPageCurrent&"', '1' ") ,  StrCnn,3,1		
	end select	 
 'Rs1.open ("Exec Retrieve_ClientStatement '"&Search_ClientFrom&"', '"&Search_ClientTo&"', '"&Search_AEFrom&"', '"&Search_AETo&"', '"&Search_AEGroup&"', '"&Search_Statement&"', '"&Search_Monthly_Month&"', '"&Search_Monthly_Year&"', '"&Search_Daily_Day&"', '"&Search_Daily_Month&"', '"&Search_Daily_Year&"', '"&iPageCurrent&"', '1' ") ,  StrCnn,3,1
 

 
	'assign total number of pages
	iPageCount = rs1(0)


  if iPageCount <= 0 then

		If Err.Number <> 0 then
			
			'SQL connection error handler
			response.write  "<table><tr><td class='RedClr'>" & MSG_BUSY & "<br></td></tr></table>"
			
		else
			'no record found
			response.write ("No record found")
				
		End If
	else
		'record found
		
		'move to next recordset
  	Set rs1 = rs1.NextRecordset() 
 
%>

</span>


<div id="reportbody1" >
<script type="text/javascript">
var somediv=document.getElementById("reportbody1")
disableSelection(somediv) //disable text selection within DIV with id="mydiv"
</script>	


<div id="PrintArea">
<table width="99%" border="0" class="normal"  cellspacing="1" cellpadding="2">
<tr bgcolor="#FFFFCC">
		<td  width="30%">　</td> 
      <% if Search_statement="Monthly" then %>
      	<td align="center" width="40%">月結單<br><u>Monthly Statement</u></td>
      <%else %>
				<td align="center" width="40%">
					<% if cint(rs1("Quantity")) > 0 then %>
							綜合日結單/買賣合約<br><u>Daily Combined Statement/Contract Note</u></td>
					<% else %>
							綜合日結單<br><u>Daily Combined Statement</u></td>
					
					<% end if%>
			<% end if %>
      <td align="right" valign="bottom" width="30%">　</td>
</tr>


<tr bgcolor="#FFFFCC">
			<td  width="30%"></td> 
			<td  width="40%"></td>      
			<td align="right" valign="bottom" width="30%" id="noprint">

						<% If PrintAllowed = 1 then %>  
							<a href="javascript:window.print()">Friendly Print</a>
						<% end if %>
			</td>
		
</tr>
</table>


<br>

<% 

''''''''''''''''''''''''''
' Header (CH)
''''''''''''''''''''''''''

Dim AccountType

if (  Not rs1.EOF)  then
			if rs1("sectioncode") = "CH" then 
						
						'Accounttype: If or not Margin client 
						AccountType = rs1("AcctType")
						%>
						<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
					
											<tr bgcolor="#ADF3B6"> 
											<td width="30%" height="33">Client Name<br>客戶名稱</td>
											<td width="70%" height="33">Client Address<br>客戶聯絡地址</td> 
											</tr>
											<tr bgcolor="#FFFFCC"> 
											<td width="30%"><%=rs1("ClntName") %></td>
											<td width="70%"><%=rs1("Address") %></td>
								</tr>
						</table>
			
						<br>
						
						<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
						
								<tr bgcolor="#ADF3B6"> 
											<td width="14%" height="33">Print Date<br>列印日期</td>
											<td width="16%" height="33">Account No.<br>客戶編號</td>
											<td width="12%" height="33">AE Code<br>經紀編號</td>
											<td width="26%" height="33">Debit Interest Rate (HKD)<br>借貸利率</td>
											
												<% if AccountType = "MRGN" then
														response.write "<td width='10%' height='33'>Margin Limit<br>保證金限額</td>"
													 end if
												%>
											
											
													<% if AccountType = "MRGN" then
														response.write "<td width='17%' height='33'>Margin Authority Expiry Date<br>信貸限額到期日</td> "
													 end if
												%>
												
								</tr>
						
								
								<tr bgcolor="#FFFFCC"> 
											<td width="14%"><%=rs1("StatementDate") %></td>
											<td width="16%">  <img border=0 src='images/tel.gif' onClick="PopupClientContact('<%=rs1("clnt") %>')"></img><%=rs1("clnt") %></td>
											<td width="12%"><%=rs1("AE") %></td>
											<td width="26%"><%= rs1("DebitInterestRate") %>%</td>
											
												<% if AccountType = "MRGN" then %>
													 	<td width='10%'> <%=formatnumber(rs1("MarginLimit")) %>
			    									</td>
														<td width='17%'>

														<%=rs1("ExpiryDate") %>
													  </td> 
												<% end if	%>

								</tr>
								
					</table>
					<%
					rs1.movenext
			end if
end if

%>


<%
'''''''''
' Loop for next records
'''''''''

do while (Not rs1.EOF)
			
			Select Case rs1("sectioncode")
					case "IN" exit do
					case "CM" exit do
					case "SM" exit do
					case "SP" exit do
					case "CB" exit do
					'case "MG" exit do
					case "CN" exit do
			
			end select
			rs1.movenext
loop

'''''''''
%>

<% 

''''''''''''''''''''''''''
' Cash statement (CM)
''''''''''''''''''''''''''

Dim TotalBalance
Dim LastCurrency
LastCurrency = ""
if (  Not rs1.EOF)  then
			
			' Cash statement section is including Interest section
			if (rs1("sectioncode") = "CM" or rs1("sectioncode") = "IN" ) then %>
			
						
						<br>
						<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
									
									<tr bgcolor="#FFFFCC"> 
												<td width="95%" colspan="12" ><span lang="zh-tw">現金帳戶記錄 Cash Statement</td>
									</tr>
									<tr bgcolor="#ADF3B6"> 
												<td width="5%">Date<br>交易日期</td>
												<td width="6%">Sdate<br>交收日期</td>
												<td width="6%">Ref No<br>參考編號</td>
												<td width="10%">Stock Code<br>股票代號</td>
												<td width="6%">Bought/In<br>買 / 入</td>
												<td width="4%">Sold/Out<br>沽 / 出</td>
												<td width="21%">Description<br>摘要</td>
												<td width="9%">Price<br>單價</td>
												<td width="7%">CCY<br>貨幣</td>
												<td width="7%">Debit<br>支帳/結欠</td>
												<td width="7%">Credit<br>存帳/結存</td>
												<td width="10%">Balance<br>結餘</td> 
									</tr>


									
									
									<% 
									''''''''''''''''''''''''''
									' Cash statement (CM)

									'TotalBalance = cdbl(rs1("Openingbalance"))
									if rs1("sectioncode") = "CM" then
											do while (  Not rs1.EOF)
											
														if rs1("sectioncode") = "CM" then
														
																' if currency change, display the balance B/F statement
																if rs1("ccy") <> LastCurrency then
																TotalBalance = cdbl(rs1("Openingbalance"))
		
																	%>
																	<tr bgcolor="#FFFFCC">
																				<td colspan="5" align="right">承上結餘Balance B/F </td>
																				<td colspan="3" align="right"> </td>
																				<td><%=rs1("CCY")%> </td>
																				<td colspan="2" align="right"> </td>
																			  <td><%=formatnumber(rs1("Openingbalance"),2,-2,-1)%> </td>
								
																	</tr>									
										            <% end if 
																	TotalBalance = TotalBalance + cdbl(rs1("Amount"))
										            
										            %>

																	
																	
																	
																	<tr bgcolor="#FFFFCC"> 
																	<td width="5%" height="19"><%=rs1("TradeDate") %>　</td>
																	<td width="6%" height="19"><%=rs1("SettleDate") %>　</td>
																	<td width="6%" height="19"><%=rs1("RefNo") %>　</td>
																	<td width="10%" height="19"><%=rs1("Instrument") %>　</td>
																	<td width="6%" height="19"><% if rs1("Quantity") > "0" then response.write formatnumber(rs1("Quantity"),0,-2,-1) %>　</td>
																	<td width="4%" height="19"><% if rs1("Quantity") < "0" then response.write formatnumber(rs1("Quantity"),0,-2,-1) %>　</td>
																	<td width="21%" height="19"><%=rs1("comment") %>　</td> 
																	<td width="9%" height="19"><% if rs1("Price") <> "0" then response.write formatnumber(rs1("Price"),3,-2,-1) %>　</td> 
																	<td width="7%" height="19"><%=rs1("CCY") %>　</td> 
																	<td width="7%" height="19"><% if rs1("Amount") < "0" then response.write formatnumber(rs1("Amount"),2,-2,-1) %>　</td> 
																	<td width="7%" height="19"><% if rs1("Amount") > "0" then response.write  formatnumber(rs1("Amount"),2,-2,-1) %>　</td> 
																	<td width="10%" height="19"><%=formatnumber(TotalBalance, 2,-2,-1)%></td> 
																	</tr>
																	<%
		
																	LastCurrency = rs1("ccy") 
																	rs1.movenext
																	
																	if not rs1.eof then

																			'display  Balance C/F if currency changed
																			if ( (rs1("ccy") <> LastCurrency and rs1("sectioncode") = "CM") or (rs1("sectioncode") <> "CM") ) then
																				%>
																						<tr bgcolor="#FFFFCC">
																									<td colspan="5" align="right">結餘轉下Balance C/F </td>
																									<td colspan="3" align="right"> </td>
																									<td><%=LastCurrency%> </td>
																									<td colspan="2" align="right"> </td>
																								  <td><%=formatnumber(TotalBalance,2,-2,-1)%> </td>
													
																						</tr>								
													            <% 
													            end if 

																			if (rs1("sectioncode") <> "CM")  then 
																					exit do
																			end if
																	end if
														end if
																				
												loop
									end if		
									%>
									
								
									<%

									''''''''''''''''''''''''''
									' Daily Interest (IN)

									if (  Not rs1.EOF)  then
											if rs1("sectioncode") = "IN" then
													 do while (  Not rs1.EOF)
													
																if rs1("sectioncode") = "IN"  then
																%>
																
																		<tr bgcolor="#FFFFCC">
																					<td width="5%"><%=rs1("ConfirmationDate")%></td>
																					<td colspan="7" align="right">Daily Interest Accrued </td>
																					<td width="5%"><%=rs1("CCY")%></td>
																					<td width="21%" colspan="2" align="right"> </td>
																				  <td width="10%"><%=formatnumber(rs1("MTDDebitInterest"),2,-2,-1)%> </td>
									
																		</tr>		
																<%
																end if	
																rs1.movenext
																
																if not rs1.eof then
																		if rs1("sectioncode") <> "IN" then 
																				exit do
																		end if
																end if
													loop

													'rs1.movenext
												end if
									end if 
								%>
						</table>

						<%
			
			end if 
end if

%>


<%
'''''''''
' Loop for next records
'''''''''

do while (Not rs1.EOF)
			
			Select Case rs1("sectioncode")
					case "IN" exit do
					case "CM" exit do
					case "SM" exit do
					case "SP" exit do
					case "CB" exit do
					'case "MG" exit do
					case "CN" exit do
			
			end select
			rs1.movenext
loop

'''''''''
%>



<% 
	''''''''''''''''''''''''''
	' Stock Journal (SM)
	''''''''''''''''''''''''''


if (  Not rs1.EOF)  then
			if rs1("sectioncode") = "SM" then %>
						<br>			
						<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
								
									<tr bgcolor="#FFFFCC"> 
												<td width="95%" colspan="9" >股票帳戶記錄 Stock Journal for </td>
									</tr>
									
									
									<tr bgcolor="#ADF3B6"> 
												<td width="5%">Date<br>交易日期</td>
												<td width="6%">Sdate<br>交收日期</td>
												<td width="6%">Ref No<br>參考編號</td>
												<td width="10%">Market<br>市場</td>
												<td width="10%">Stock Code<br>股票代號</td>
												<td width="6%">Bought/In<br>買/入</td>
												<td width="4%">Sold/Out<br>沽/出</td>
												<td width="15%">Description<br>摘要</td>
												<td width="20%">Stock Name<br>股票名稱</td> 
									</tr>
									
									<% do while (  Not rs1.EOF)
									
												if rs1("sectioncode") = "SM"  then
												%>
												
															<tr bgcolor="#FFFFCC"> 
																		<td width="5%"><%= rs1("tradedate") %></td>
																		<td width="6%"><%= rs1("settledate") %>　</td>
																		<td width="6%"><%= rs1("RefNo") %>　</td>
																		<td width="10%"><%= rs1("Market") %>　</td>
																		<td width="10%"><%= rs1("Instrument") %>　</td>
																		<td width="6%"><%if clng(rs1("Quantity")) > 0 then response.write formatnumber(rs1("Quantity"),0) end if%>　</td>
																		<td width="4%"><%if clng(rs1("Quantity")) < 0 then response.write formatnumber(abs(cDbl(rs1("Quantity"))),0) end if%>　</td>
																		<td width="15%"><%= rs1("Comment") %>　</td> 
																		<td width="20%"><%= rs1("InstrumentDesc") %>　</td> 
															</tr>
															<%
												end if	
												rs1.movenext
												
												if not rs1.eof then
														if rs1("sectioncode") <> "SM" then 
																exit do
														end if
												end if
									loop
									
									
									
									%>
						</table>
			<%
			end if 

end if
%>

<%
'''''''''
' Loop for next records
'''''''''

do while (Not rs1.EOF)
			
			Select Case rs1("sectioncode")
					case "IN" exit do
					case "CM" exit do
					case "SM" exit do
					case "SP" exit do
					case "CB" exit do
					'case "MG" exit do
					case "CN" exit do
			
			end select
			rs1.movenext
loop

'''''''''
%>


<% 

	''''''''''''''''''''''''''
	'  Portfolio Statement (SP)
	''''''''''''''''''''''''''

dim TotalPortfolioValue
dim TotalMarginValue

TotalPortfolioValue=0
TotalMarginValue=0

if (  Not rs1.EOF)  then
			if rs1("sectioncode") = "SP" then %>
			
						<br>
						<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
								
									<tr bgcolor="#FFFFCC"> 
												<td width="95%" colspan="13">投資組合記錄 Portfolio Statement </td>
									</tr>
									
									<tr bgcolor="#ADF3B6">
												<td width="10%">Market<br>市場</td>
												<td width="10%">Stock Code<br>股票代號</td>
												<td width="12%">Stock Name<br>股票名稱</td>
												<td width="10%">Holding B/F<br>承上</td>
												<td width="15%">Movement<br>In/Out<br>出/入</td>
												<td width="10%">Holding C/F<br>轉下</td>
												<td width="10%">Closing Price<br>收市價</td>
												<td width="10%">CCY<br>貨幣</td> 
												<td width="17%">Value<br>市場價值</td> 
												<td width="17%">#Exch Rate<br>匯率</td> 
												<td width="17%">Value (HKD)<br>市場價值(港元)</td> 
																															
													<%if AccountType = "MRGN" then %>
															
														<td width="17%"> %</td>
														<td width="17%">Margin Value<br>按倉價值(港元)</td> 
													<%	end if %>
									</tr>
									
									<% 
									do while (  Not rs1.EOF)
									
												if rs1("sectioncode") = "SP" then
												%>

															<tr bgcolor="#FFFFCC"> 
																		<td width="10%"><%= rs1("Market") %></td>
																		<td width="10%"><%= rs1("Instrument") %></td>
																		<td width="12%"><% response.write rs1("InstrumentDesc") & "<br>" & rs1("InstrumentCDesc") %>　</td>
																		<td width="10%">
																			<%
																			
																			response.write formatnumber(rs1("OpeningBalance"),0,-2,-1) 
																			
																				if (clng(rs1("orfee1")) <> 0  ) then
																					response.write "<br>*" & formatnumber(clng(rs1("orfee1")),0,-2,-1)
																				end if																			
																			%>
																			</td>
																		<td width="10%">
																			<%
																				response.write formatnumber(rs1("Netbalance"),0,-2,-1) 

																			%>　
																		</td>
																		<td width="10%">
																			<%
																				response.write formatnumber(rs1("EndingBalance"),0,-2,-1) 
																				'stock on hold positive
																				'response.write clng(rs1("orfee1"))
																				if (clng(rs1("orfee1")) <> 0  ) then
																					response.write "<br>*" & formatnumber(clng(rs1("orfee2")),0,-2,-1)
																				end if
																				
														
																			%>　
																		</td>

																		<td width="10%"><%= rs1("Price") %>　</td>
																		<td width="10%"><%= rs1("CCY") %>　</td> 
																		<td width="17%"><%= formatnumber(cDbl(rs1("StockPortValue")) / cDbl(rs1("MarginLimit")),2,-2,-1) %>　</td> 
																		<td width="17%"><%= formatnumber(rs1("MarginLimit"),4) %>　</td> 
																		<td width="17%"><%= formatnumber( int(cDbl(rs1("StockPortValue"))*100)/100  ,2,-2,-1) %>　</td> 


																		<%if AccountType = "MRGN" then %>
																				
																			<td width="17%"><%=rs1("MarginPercent")%></td>
																			<td width="17%"><%=formatnumber(cDbl(rs1("MarginValue")),2,-2,-1)%></td> 
																		<%	end if %>
																		
																		<%
																				TotalPortfolioValue  = TotalPortfolioValue + cDbl(rs1("StockPortValue"))
																				TotalMarginValue = TotalMarginValue + cDbl(rs1("MarginValue"))
																		%>
															</tr>
															

															<%
												end if	
												rs1.movenext
												if not rs1.eof then
														if rs1("sectioncode") <> "SP" then 
																exit do
														end if
												end if
									loop
							
									
									%>
									<tr bgcolor="#FFFFCC">
												<td colspan="10" align="right">
													<b>股份組合市值( Portfolio Value (HKD): </b></td>
												<td><b><%=formatnumber(TotalPortfolioValue,2,-2,-1) %></b></td>

													<%
													if AccountType = "MRGN" then 
														response.write "<td></td><td><b>"
														response.write formatnumber(TotalMarginValue,2,-2,-1) 
														response.write "</b></td>"
													end if %>
													
												</td>
									</tr>
									
						</table>
			<%
			
			end if 
end if

%>

<%
'''''''''
' Loop for next records
'''''''''

do while (Not rs1.EOF)
			
			Select Case rs1("sectioncode")
					case "IN" exit do
					case "CM" exit do
					case "SM" exit do
					case "SP" exit do
					case "CB" exit do
					'case "MG" exit do
					case "CN" exit do
			
			end select
			rs1.movenext
loop

'''''''''
%>


<%
	''''''''''''''''''''''''''
	'  Cash Balance (CB)
	''''''''''''''''''''''''''

dim TotalCashValue
dim TotalNetValue
dim TotalMarginNetValue

TotalCashValue=0
TotalNetValue=0
TotalMarginNetValue=0

'Response.write rs1("sectioncode")

if (  Not rs1.EOF)  then
			if rs1("sectioncode") = "CB" then %>

			
						<br>
						<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
				
								
						
									<tr bgcolor="#FFFFCC"> 
												<td width="115%" colspan="8">現金結餘 Cash Balance</td>
									</tr>
									
									<tr bgcolor="#ADF3B6">
												<td width="14%">CCY<br>貨幣</td>
												<td width="16%">Available Balance<br>可用結餘</td>
												<td width="12%">Unavailable Balance<br>未可用結餘</td>
												<td width="26%">Ledger Balance<br>帳面結餘</td>
												<td width="10%">## Interest Accrued<br>## 應計利息</td>
												<td width="10%">#Exchange Rate<br># 匯率</td>
												<td width="10%">Balance (HKD)<br>結餘</td>
												<% if AccountType = "MRGN" then %>
														<td width="17%"></td> 
												<% end if %>
									</tr>
									
									
									<%  
									do while ( Not rs1.EOF)
									
												if rs1("sectioncode") = "CB" then
												%>
												
												<tr bgcolor="#FFFFCC"> 
															<td width="14%"><%= rs1("CCY") %></td>
															<td width="16%"><%= formatnumber(rs1("OpeningBalance"),2,-2,-1) %>　</td>
															<td width="12%"><%= formatnumber(rs1("Netbalance"),2,-2,-1) %>　</td>
															<td width="26%"><%= formatnumber(rs1("EndingBalance"),2,-2,-1) %>　</td>
															<td width="10%"><%= formatnumber(rs1("MTDDebitInterest"),2,-2,-1)  %>　</td>
															<td width="10%"><%= formatnumber(rs1("MarginLimit"),4,-2,-1) %>　</td>
															<td width="10%"><%= formatnumber(rs1("StockPortValue"),2,-2,-1) %>　</td>
															<% if AccountType = "MRGN" then %>
																	<td width="17%"></td> 
															<% end if %>
															
												</tr>
												<%
												TotalCashValue=TotalCashValue+cdbl(rs1("StockPortValue"))
												end if	
												rs1.movenext
												
												
												if not rs1.eof then
														if rs1("sectioncode") <> "CB" then 
																exit do
														end if
												end if
									loop
									%>
									
									<%
										TotalNetValue = TotalCashValue  + TotalPortfolioValue
										TotalMarginNetValue = TotalBalance + TotalMarginValue
									%>
									<tr bgcolor="#FFFFCC">
												<td colspan="6" align="right"><b>現金值 Cash Value: </b></td>
												<td><%=formatnumber(TotalCashValue,2,-2,-1) %></td>
															<% if AccountType = "MRGN" then %>
																	<td width="17%"></td> 
															<% end if %>
								  </tr>
								  <tr bgcolor="#FFFFCC">
												<td colspan="6" align="right"><b>總值 Net Value: </b></td>
												<td><b><%=formatnumber(TotalNetValue,2,-2,-1) %></b></td>
												
													<%
													if AccountType = "MRGN" then %>
														<td><b> <%= formatnumber(TotalMarginNetValue,2,-2,-1) %>
															</b>
														</td>
													<% end if %>												
									</tr>
									
									<% if AccountType = "MRGN" then %>
									<tr bgcolor="#FFFFCC">
												<td colspan="7" align="right"></td>
												
													<%
													
														if TotalMarginNetValue >= 0 then
																	response.write "<td colspan=1><b>Margin Available</b></td>"
														else
																	response.write "<td colspan=1><b>Margin Call</b></td>"
														end if
													%>
													
								  </tr>
									<% end if 
									
									'foot note
									%>
									<tr bgcolor="#FFFFCC">
											<td colspan="4">
												# Exchange Rate used are for indication only 匯率只作指引用途 <br>
												## Interest Accrued  will be credited/debited to your account on the last day of each month
												<br> 應計利息將會於每月的最後一天從閣下之戶口存入\提取
											<td>
											<td><td>
											<% if AccountType = "MRGN" then %>
													<td width="17%"></td> 
											<% end if %>									</tr>	
						</table>
						
						<%
						' Foot Note
						
						
						
			end if
					
			 
'end if
else
%>
<br>
			<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">			 <tr bgcolor="#FFFFCC">
					<td align="right"><b>總值 Net Value: </b></td>
					<td width="20%"><b><%=formatnumber(TotalPortfolioValue,2,-2,-1) %></b></td>
												
																						
			  </tr>
			</table>
<%
			End If	

%>

<%
'''''''''
' Loop for next records
'''''''''

do while (Not rs1.EOF)
			
			Select Case rs1("sectioncode")
					case "IN" exit do
					case "CM" exit do
					case "SM" exit do
					case "SP" exit do
					case "CB" exit do
					'case "MG" exit do
					case "CN" exit do
			
			end select
			rs1.movenext
loop

'''''''''
%>

<% 
	''''''''''''''''''''''''''
	'  Trade Detail (CN)
	''''''''''''''''''''''''''

if (  Not rs1.EOF)  then
	if rs1("sectioncode") = "CN" then 
	
	
	'temp instrument datastore
	dim j,k

  dim Market
  dim RefNo
  dim TradeDate
  dim SettleDate
  dim Instrument
  dim BuySell
  dim CCY  
  dim NetBalance
  dim Amount 
  dim InstrumentDesc	
  dim InstrumentCDesc	
	dim orfee(10)
	dim feename(10)
	dim quantity()
	dim price()
 
	%>


<br>

<table  width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">

						
<tr bgcolor="#FFFFCC"> 
      <td width="166%" height="18" colspan="12">交易詳情 <span lang="en-us">Trade Details</span></td>
</tr>

<tr bgcolor="#ADF3B6">
   <td width="8%" height="18">Market<br>市場</td>
   <td width="5%" height="18">Trade No<br>交易編號</td>
   <td width="5%" height="18">T/D <br>交易</td>
   <td width="5%" height="18">S/D<br>交收</td>
   <td width="18%" height="18">Stock<br>股票</td>
   <td width="10%" height="18">B/S<br>買沽</td>
   <td width="5%" height="18">Quanlity<br>數量</td>
   <td width="5%" height="18">Price<br>成交價</td>
   <td width="5%" height="18">CCY<br>貨幣</td> 
   <td width="10%" height="18">Gross Amount<br>總額</td> 
   <td width="17%" height="18">Transaction Cost<br>交易費用</td> 
   <td width="10%" height="18">Net Amount<br>淨額</td> 
</tr>
		<% do while (  Not rs1.EOF)

				
				if rs1("sectioncode") = "CN" and rs1("comment") = "1" then
					j=0


					'destroy array
					erase price
					erase quantity		

				  Market          = rs1("Market")
				  RefNo           = rs1("RefNo")
				  TradeDate       = rs1("TradeDate")
				  SettleDate      = rs1("SettleDate")
				  Instrument      = rs1("Instrument")
				  BuySell         = rs1("BuySell")
				  CCY             = rs1("CCY")
				  NetBalance      = rs1("NetBalance")
				  InstrumentDesc	= rs1("InstrumentDesc")
				  InstrumentCDesc	= rs1("InstrumentCDesc")
					Amount          = rs1("Amount")
					
					feename(0) = rs1("feename1")
					feename(1) = rs1("feename2")
					feename(2) = rs1("feename3")
					feename(3) = rs1("feename4")
					feename(4) = rs1("feename5")
					feename(5) = rs1("feename6")
					feename(6) = rs1("feename7")
					feename(7) = rs1("feename8")
					feename(8) = rs1("feename9")
					feename(9) = rs1("feename10")

					orfee(0) = rs1("ORFee1")
					orfee(1) = rs1("ORFee2")
					orfee(2) = rs1("ORFee3")
					orfee(3) = rs1("ORFee4")
					orfee(4) = rs1("ORFee5")
					orfee(5) = rs1("ORFee6")
					orfee(6) = rs1("ORFee7")
					orfee(7) = rs1("ORFee8")
					orfee(8) = rs1("ORFee9")
					orfee(9) = rs1("ORFee10")
							
					'store orfee and quantity into array
					
					'loop to next new ref no
					do while (Not rs1.EOF) 'find all quantity and price of trading

						
						if  rs1.eof or rs1("RefNo") <> RefNo  then 
						%>

		
						<%  'write table content %>
						<tr bgcolor="#FFFFCC"> 
						   <td><%= Market %>　</td>
						   <td><%= RefNo %>　</td>
						   <td><%= TradeDate %>　</td>
						   <td><%= SettleDate %>　</td>
						   <td><%= Instrument %>&nbsp;<%= InstrumentDesc %>&nbsp;<%= InstrumentCDesc %>　</td>
						   <td><%= BuySell %>　   </td>
						   <td><br>
						   <%
						   For Each item In quantity
								if item <> "" then
									Response.Write(formatnumber(item,0) & "<br>")
								end if
								Next
							%>　  
						   </td>
						   <td><br>
						   <%
						   For Each item In price
								if item <> "" then
									Response.Write(formatnumber(item,4) & "<br>")
								end if
								Next
							%>　      
						   </td>
						   
						   <td><%= CCY %>　       </td> 
						   <td><%= formatnumber(NetBalance,2,-2,-1) %>　</td> 
						   <td>
						   <% 
						   For k=0 to 10 
						   	if feename(k) <> ""  and cDbl(orfee(k)) <> 0  then 
						   		Response.Write (feename(k) & ": " & formatnumber(cDbl(orfee(k)),2,-2,-1) & "<br>")
						   	end if
						   next %>
			
						   </td> 
							 <td width="17%" height="18"><%= formatnumber(cDbl(Amount),2,-2,-1) %>　</td> 
						</tr>							
					<%
						'move to next new ref no
						exit do
					else 
				
							' if not the first record, add price and quantity into array
							ReDim Preserve price(j+1)
							ReDim Preserve quantity(j+1)	
							price(j) = rs1("Price")
							quantity(j) = rs1("Quantity")
							j=j+1
							rs1.movenext
							
							
					end if
					

				loop



				
			else
							rs1.movenext
			end if	' end of if rs1("sectioncode") = "CN" and rs1("comment") = "1" then
			

			
			if not rs1.eof then
				if rs1("sectioncode") = "SP" then
					exit do
				end if
			end if
			
		loop
	          %>
		<%  'write table content 
				'write buffered record
		%>
		
						<tr bgcolor="#FFFFCC"> 
						   <td><%= Market %>　</td>
						   <td><%= RefNo %>　</td>
						   <td><%= TradeDate %>　</td>
						   <td><%= SettleDate %>　</td>
						   <td><%= Instrument %>&nbsp;<%= InstrumentDesc %>&nbsp;<%= InstrumentCDesc %>　</td>
						   <td><%= BuySell %>　   </td>
						   <td><br>
						   <%
						   For Each item In quantity
								if item <> "" then
									Response.Write(formatnumber(item,0) & "<br>")
								end if
								Next
							%>　  
						   </td>
						   <td><br>
						   <%
						   For Each item In price
								if item <> "" then
									Response.Write(formatnumber(item) & "<br>")
								end if
								Next
							%>　      
						   </td>
						   
						   <td><%= CCY %>　       </td> 
						   <td><%= formatnumber(NetBalance,2,-2,-1) %>　</td> 
						   <td>
						   <% 
						   For k=0 to 10 
						   	if feename(k) <> ""  and cDbl(orfee(k)) <> 0  then 
						   		Response.Write (feename(k) & ": " & formatnumber(cDbl(orfee(k)),2,-2,-1) & "<br>")
						   	end if
						   next %>
			
						   </td> 
							 <td width="17%" height="18"><%= formatnumber(cDbl(Amount),2,-2,-1) %>　</td> 
						</tr>				

	<%
		


%>


 </table>
 
 	<%end if 
	end if
	
		%>
		<br>
		<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
		<tr bgcolor="#FFFFCC"> 
					<td width="166%" height="18" colspan="12">
						1. For Margin Accounts, the margin limit is subject to change from time to time
						at our sole discretion. Your current margin limit is stated in the first page of 
						this statement
						 <br>
						 2. Stock Borrowing Services for Hong Kong Market is now available in UOB KayHian 
						(Hong Kong) Limi ted, for more details, please contact your Account Manager or customer
						service hotline at 2826 4868. 
						<br>
						3. Effective from 11 Aug 2009, we require all clients to sign the W-8BEN certificate to 
						certify their foreign status of US before trading US Securities. W-8BEN certificate is a 
						Certificate of Foreign Status of US. With W-8BEN certificate, the client wil l be exempted 
						from US Capital Gains Tax. For further information, kindly contact your responsible salespersons.
			</td>
		</tr>		
	</table>
	<br>
		<table width="99%" border="0" class="normal" style="border-width: 0" bgcolor="#808080" cellspacing="1" cellpadding="2">
		<tr bgcolor="#FFFFCC"> 
					<td width="166%" height="18" colspan="12" align="center">結單完結<br>End of Statement</td>
		</tr>
		
	</table>
	</td>
	</tr>
</table>

</div>      



<% if (session("shell_power") = 1 or session("shell_power") = 5) then %>  
		</span>
<% end if %>

<%
	
end if   'having client number if statement
%>



<%
'*****************************************************************
' End of report body
'*****************************************************************
%>
</span>


</td></tr></table>
</body>
</html>

<%
'*****************************************************************
' Termination
'*****************************************************************

 'Rs1.Close
 set Rs1 = Nothing
 'Rs2.close
 Set Rs2 = Nothing
 Conn.Close
 Set Conn = Nothing
%>