

'Usage:
'   CScript (scriptname).vbs <dbname>

dim sFilename
dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

If (WScript.Arguments.Count <> 1) then
	MsgBox "You must specify the filename of the database", vbExclamation, "Error"
	Wscript.Quit()
End if

sFilename = fso.GetAbsolutePathName(WScript.Arguments(0))


Remove_DBO_Prefix sFilename


Function Remove_DBO_Prefix(sDbFilename)
	
	Dim oApplication
	Set oApplication = CreateObject("Access.Application")
	oApplication.OpenCurrentDatabase sDbFilename
	
	Dim dbs
	Set dbs = oApplication.CurrentData
	
	Dim obj
	WScript.Echo "Looping through tables..."
	For Each obj in dbs.AllTables
		If Left(obj.Name, 4) = "dbo_" then
			WScript.Echo "  table name=(" & obj.Name & "), left part=(" & Left(obj.Name, 4) & "),  new = (" & Mid(obj.Name, 5) & ")"
			'WScript.Echo "Doing the rename..."
			' This can't possibly work... it uses "acTable", which doesn't exist...?
			'DoCmd.Rename Mid(obj.Name, 5), acTable, obj.Name
		End If
	Next
	WScript.Echo "[DONE]"
	
	oApplication.CloseCurrentDatabase
	
	REM Set oApplic
	REM 'As Object
	REM oApplication.OpenCurrentDatabase sDbFilename
	
REM '	= Application.CurrentData
	
	REM ''''Set oApplication
	
	REM '
	
	
	REM 'Search for open AccessObject objects in AllTables collection.
	REM Dim obj As AccessObject
	REM For Each obj In oApplication.AllTables
		REM 'If found, remove prefix
		REM If Left(obj.Name, 4) = "dbo_" Then
			REM DoCmd.Rename Mid(obj.Name, 5), acTable, obj.Name
		REM End If
	REM Next obj
End Function