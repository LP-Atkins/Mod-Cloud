'Valheim-Cloud Server Client
'Re-written in vbs because windows defender kept falsly flagging the .exe version as a trojan.

'Instructions
'1.) Create a shared cloud storage location (I use onedrive).
'2.) Sync the library to your cloud storage so that it appears as a local folder.
'3.) Run the app once to create the mods folder.
'4.) Add all changes to your game files from the root directory of the game files.
'5.) Press the launch button to launch the game.

'Note: If the game directory is not found you will need to input location of the vanilla files.


'Original Release by LongParnsip
'-------------------------------------------------------------------------------
'Version        Date                Author                  Description
'
'v1.00          24/09/2021          LongParsnip             Original Release (exe release)
'v1.01			24/09/2021			LongParsnip				VBS release
'

'User Settings.
Dim ModdedName:		ModdedName = "Cloud"		'The name of the modified valheim folder, This string will be appened to the valheim folder when it is copied.
Call Main


	
	Sub Main()
		
		'Find the valheim directory.
		Dim sVanillaPath:	sVanillaPath = FindInstallDir(":\Program Files (x86)\Steam\steamapps\common\Valheim")
		Dim sModdedPath:	sModdedPath = sVanillaPath & " " & ModdedName
		If sVanillaPath = "NOT FOUND" Then
			Msgbox "Couldn't locate the valheim directory", vbOkOnly, "Valheim Missing"
			Exit Sub
		End If
		
		'Find the mods folder.
		Dim sModsFolder: sModsFolder = GetScriptPath & "\Mods"
		If Not FolderExists(sModsFolder) Then
			Msgbox "Couldn't locate the 'mods' folder, make sure it is in the same directory as this script.", vbOkOnly, "Mods Missing"
			Exit Sub
		End If

		'Create Files
		Msgbox RoboCopy(sVanillaPath, sModdedPath, "/MIR")	'Game Files
		Msgbox RoboCopy(sModsFolder, sModdedPath, "/s")   	'Copy mods to the correct locations.
		

		'Launch Application
		Call OpenFile(sModdedPath & "\Valheim.exe")
		

	End Sub









'---------------------------------------------------------------------------------------------
'----- Subprocedures
'---------------------------------------------------------------------------------------------

	'Opens a file at the specified path.
	'Call openFile("C:\MyFile.txt")
	Public Sub OpenFile(strFilePath)
	Dim objShellApplication: Set objShellApplication = CreateObject("Shell.Application")
		objShellApplication.Open strFilePath
	End Sub


'---------------------------------------------------------------------------------------------
'----- Functions
'---------------------------------------------------------------------------------------------

	'Finds the installed folder directory:
	Public Function FindInstallDir(strInstallDir)
	Dim i
	Dim bFound

		'64 Bit Windows.
		For i = 3 To 26
			If FolderExists(ConvertToLetter(i) & strInstallDir) Then
				FindInstallDir = ConvertToLetter(i) & strInstallDir
				bFound = True
				Exit For
			End If
		Next

	If FindInstallDir = "" Then FindInstallDir = "NOT FOUND"
	End Function


	'Returns the path that the script is executing from.
	Public Function GetScriptPath()
		GetScriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
	End Function


	'Checks if a file exists.
	Public Function FileExists(strPath)
	Dim objFSO:	Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FileExists(strPath) Then FileExists = True
	End Function
	
	
	'Checks if a folder exists.
	Public Function FolderExists(strPath)
	Dim objFSO:	Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(strPath) Then FolderExists = True
	End Function


	'Converts a number to a letter
	'https://support.microsoft.com/en-au/help/833402/how-to-convert-excel-column-numbers-into-alphabetical-characters
	'Call ConvertToLetter(1)
	Public Function ConvertToLetter(iCol)
	   Dim iAlpha
	   Dim iRemainder
	   iAlpha = Int(iCol / 27)
	   iRemainder = iCol - (iAlpha * 26)
	   If iAlpha > 0 Then
		  ConvertToLetter = Chr(iAlpha + 64)
	   End If
	   If iRemainder > 0 Then
		  ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
	   End If
	End Function
	
	
	'Robocopy with StdOut
	Public Function RoboCopy(src, dest, switches)
	Dim objShell:   Set objShell = CreateObject("Wscript.Shell")
	Dim objExec
	Dim strTemp

		Set objExec = objShell.Exec("robocopy " & Chr(34) & src & Chr(34) & " " & Chr(34) & dest & Chr(34) & " " & switches)
		Do While Not objExec.StdOut.AtEndOfStream
			strTemp = strTemp & objExec.StdOut.ReadLine()
		Loop
		
		
		RoboCopy = strTemp
	End Function