<HTML>
<Body>

<%

TheTimes = 4000
	Response.write "Use normal String Concatenation Method (&) <BR>"
	Response.Write time() 
	Response.write "<BR>"

	for i = 1 to TheTimes
		str = str & "1" & "1"  & "1"  & "1" & "1" & "1" & "1"  & "1"  & "1" & "1" 
	next

'	response.write Str 
	response.write "The Length is " & len(str)
	Response.write "<BR>"

	Response.write "Use Fast String Concatenation Method <BR>"
	Response.Write time()
	Response.write "<BR>"
dim tmp

Set tmp = new StringBuilder

	For i = 1 to TheTimes
		tmp.Append ("1")
		tmp.append ("1")
		tmp.append ("1")
		tmp.append ("1")
		tmp.append ("1")
		tmp.Append ("1")
		tmp.append ("1")
		tmp.append ("1")
		tmp.append ("1")
		tmp.append ("1")

	next 

	str = tmp.ToString() 
'	response.write str 
	response.write "The Length is " & len(str)
	Response.write "<BR>"
	Response.Write time()



Class StringBuilder
	Dim arr
	Dim growthRate
	Dim itemCount 

	Private Sub Class_Initialize()
		growthRate = 50
		itemCount = 0
		ReDim arr(growthRate)
	End Sub

	Public Sub Append(ByVal strValue)
		If itemCount > UBound(arr) Then
			ReDim Preserve arr(UBound(arr) + growthRate)
		End If

		arr(itemCount) = strValue
		itemCount = itemCount + 1
	End Sub

	Public Function ToString() 
		ToString = Join(arr, "")
	End Function
End Class


%>

</Body>
</HTML>