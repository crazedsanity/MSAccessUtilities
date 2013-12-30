' Usage:
'	CScript (thisfile).vbs <DRIVE:\path\to\source\files\>

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")


'' TODO: get these constants from vbscript!
const acForm = 2
const acModule = 5
const acMacro = 4
const acReport = 3


dim sADPFilename
If (WScript.Arguments.Count = 0) then
    MsgBox "Please specify the filename!", vbExclamation, "Error"
    Wscript.Quit()
End if
sADPFilename = WScript.Arguments(0)
sADPFilename = fso.GetAbsolutePathName(sADPFilename)

sBasePath = fso.GetAbsolutePathName(sADPFilename)
sBasePath = fso.GetParentFolderName(sBasePath)
Wscript.Echo "BASE: " & sBasePath
sBasePath = sBasePath & "\Source\"

WScript.Echo "FILE::: "& sADPFilename & ", PATH::: "& sBasePath


Dim oApplication
Set oApplication = CreateObject("Access.Application")

WScript.Echo "Opening database: " & sADPFilename
oApplication.OpenCurrentDatabase sADPFilename


dim sRealExt
dim sTypeExt
dim sBaseName
dim sJustName

Dim myFile
Dim folder
Set folder = fso.GetFolder(sBasePath)

Dim myFolder
Dim sDisplay
Dim iType
Dim sScriptPath
For Each myFolder in folder.SubFolders
	WScript.Echo "Processing folder ("& myFolder.Name &")..."
	For Each myFile In myFolder.Files
		sBaseName = fso.GetBaseName(myFile.Name)
		sJustName = fso.GetBaseName(sBaseName)
		sRealExt = fso.GetExtensionName(myFile)
		sTypeExt = fso.GetExtensionName(sBaseName)
		sDisplay = ""
		
		
		Select Case sTypeExt
			Case "frm"
				sDisplay = "Form"
				iType = acForm
			Case "mcr"
				sDisplay = "Macro"
				iType = acMacro
			Case "bas"
				sDisplay = "Module"
				iType = acModule
			Case "report"
				sDisplay = "Report"
				iType = acReport
			Case else 
				Wscript.Echo "Invalid type (" & sTypeExt &")"
				WScript.Quit()
		End Select
		
		sScriptPath = sBasePath & myFolder.Name & "\" & myFile.Name
		WScript.Echo "  "& sDisplay & " -- '" & sJustName 
		' & "' -> "& sScriptPath
		
		oApplication.LoadFromText iType, sJustName, sScriptPath
	Next
Next

oApplication.CloseCurrentDatabase