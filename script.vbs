loaded = true
sub hello
msgbox("hello!")
end sub

Sub LoadListItems1
	Dim colOptions
	Dim objOption
	
	Set colOptions = ToolList.options
	
		Set objOption = document.createElement("OPTION")
		colOptions.add(objOption)
		objOption.innerText = "11. Online Test 1"

		Set objOption = document.createElement("OPTION")
		colOptions.add(objOption)
		objOption.innerText = "12. Online Test 2"
End Sub

Sub MoreTools
	IF tools99.selectedIndex = 10 Then  '' Blank
		result = MsgBox ("Not Yet Implemented" & vbCrLf  & vbCrLf & "Do you want to continue?", vbYesNo, "Title")
	
		Select Case result
		Case vbYes
			''Set objFile2 = fso.CreateTextFile(strTools)
			''strLine = 	"TRUNCATE TABLE plu;" & _
			''			"TRUNCATE TABLE plulocation;" & _
			''			"TRUNCATE TABLE nutrifacts"
			''objFile2.WriteLine strLine
			''objFile2.Close
			''objShell.Run "cmd /C " & strMySQL & " -u gap -pgap -sN -h " & Host_IP.Value & " " & Host_DB.Value & " < " & strTools, 0, True
			''MsgBox("Complete")
		Case vbNo
		End Select
	End If
	
	IF tools99.selectedIndex = 11 Then  '' Blank
		result = MsgBox ("Not Yet Implemented" & vbCrLf  & vbCrLf & "Do you want to continue?", vbYesNo, "Title")
	
		Select Case result
		Case vbYes
			''Set objFile2 = fso.CreateTextFile(strTools)
			''strLine = 	"TRUNCATE TABLE plu;" & _
			''			"TRUNCATE TABLE plulocation;" & _
			''			"TRUNCATE TABLE nutrifacts"
			''objFile2.WriteLine strLine
			''objFile2.Close
			''objShell.Run "cmd /C " & strMySQL & " -u gap -pgap -sN -h " & Host_IP.Value & " " & Host_DB.Value & " < " & strTools, 0, True
			''MsgBox("Complete")
		Case vbNo
		End Select
	End If

End Sub
