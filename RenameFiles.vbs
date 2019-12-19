' Initial Version Created by Wolfix Cai on 2016/06/15
' 2016/06/15	Add Filter for File starts with JCI-0764 


Dim re, s, objMatch, colMatches
Set re = New RegExp
re.Global = True
re.IgnoreCase = True
	
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
	
'Input Folder
myFolder = InputBox("Enter Folder which contains the files:")

' Collect All Files
For Each myFile In objFSO.GetFolder(myFolder).Files
	' Filter out Files starts with: JCI-0764
	re.Pattern = "^JCI-0764"
	If re.Test(myFile.Name) Then
		' Read File
		Set objTextFile = objFSO.OpenTextFile(myFile, ForReading)
		strText = objTextFile.ReadAll
		objTextFile.Close

		' Check Content
		'Define the Rule
		re.Pattern = "CSG\+19461010"

		If re.Test(strText) Then
			'Found
			'Rename the file
			' Old: JCI-0764.O01770000000000X00A.AVIEXP.000415409
			' New: BMW.SY78.JIT.000415409.AVIEXP.dat
			myIDArray = split(myFile.Name, ".")
			myID = myIDArray(ubound(myIDArray))
			
			myFile.Name = "BMW.SY78.JIT." & myID & ".AVIEXP.dat"
		Else
			' Check Another Pattern
			re.Pattern = "CSG\+15281010"
			If re.Test(strText) Then
				myIDArray = split(myFile.Name, ".")
				myID = myIDArray(ubound(myIDArray))
				
				myFile.Name = "BMW.SY88.JIT." & myID & ".AVIEXP.dat"
			Else
				msgbox myFile & " does not have valid content!"
			End if
		End If

		set strText = Nothing
	End If
Next

'Msgbox "Completed!"
