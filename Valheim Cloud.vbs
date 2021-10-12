'Valheim-Cloud Server Client
'Re-written in vbs because windows defender kept falsly flagging the .exe version as a trojan.
'
'Operational description:
'-----------------------------
'1.) Downloads the latest mod file and extracts it.
'2.) Makes a copy of the original game files.
'3.) Patches the game folder with the mods directory.
'4.) Launches the game via a steam command.
'5.) Waits until the game is closed and then restores the vanilla files.
'

'Instructions ***IMPORTANT*** - This script was written for useage with Microsoft Onedrive only.
'Note1: Check the 'How to use.pdf' file in the repository for more detailed instructions.
'Note2: The default embeddedURL will point you to my public server, this server will continue to run indefinitely.
'-----------------------------
'1.) Create a zip file of all your mods from the root directory of valheim.
'2.) Upload to onedrive and generate an embedded link.
'3.) Copy the embedded 'src' link and replace the user setting variable embeddedURL with the new string.
'4.) Save the .vbs file and distribute to clients.
'5.) Launch the game via this vbs file.

'Note: If the game directory is not found you will need to input location of the vanilla files.


'Original Release by LongParnsip
'Current Version: 1.03
'-------------------------------------------------------------------------------
'Version        Date                Author                  Description
'
'v1.00          24/09/2021          LongParsnip             Original Release (exe release).
'v1.01			24/09/2021			LongParsnip				VBS release.
'v1.02			27/09/2021			LongParsnip				Added SteamUI support.
'v1.03			12/10/2021			LongParsnip				Re-written to download from a direct link which makes end user usage a 'one click' operation.


'User Settings:
'-----------------------------
Dim embeddedURL:	embeddedURL = "https://onedrive.live.com/embed?cid=6002A8DD978B973A&resid=6002A8DD978B973A%2142416&authkey=AOw7p-kegsQxPrw"		'This is the embedded link to the onedrive file.
Dim BackupName:		BackupName = "(unmodded)"																										'The name of the modified backed up valheim folder, which will be restored when the game closes.
Call Main


Sub Main()

	
	'Find the valheim directory by reading the registry and steam library file.
	Dim sVanillaPath: sVanillaPath = LocateFolder
	If sVanillaPath = "NOT FOUND" Then
		Msgbox "Couldn't locate the valheim directory", vbOkOnly, "Valheim Missing"
		Exit Sub
	End If
	If FolderExists(sVanillaPath & "\BepInEx") Then	'Mods detected warning.
		If Msgbox("It looks like your game is modified, it is recommend to restore it to vanilla. Would you like to cancel the launch?", vbYesNo, "Mods Detected") = vbYes Then Exit Sub
	End If
	
	
	
	'Create a MSXML object so that data can be scraped from the zip file URL.
	Dim objMSXML:	Set objMSXML = CreateObject("MSXML2.ServerXMLHTTP")
	Dim downloadURL:	downloadURL = Replace(embeddedURL, "embed", "download")
		objMSXML.open "GET", embeddedURL, false
		objMSXML.send



	'Get the last modified data.
	Dim strTemp, arrTemp, DateModified
		strTemp = Mid(objMSXML.responseText,Instr(1, objMSXML.responseText, """modifiedDate""") + 15, 30)
		arrTemp = Split(strTemp,","): DateModified = arrTemp(0)



	'Download File, if required.
	CreateFolder(sVanillaPath & "\mods")	'Create mods directory if it doesn't exist.
	If Not FileExists(sVanillaPath & "\mods\mods-" & DateModified & ".zip") Then
		Msgbox "Downloading Update - Press OK to start the download, this may take a while.", vbOkOnly, "Update Available"
		Dim objStream:		Set objStream = CreateObject("Adodb.Stream")
			objMSXML.Open "GET", downloadURL, false
			objMSXML.Send
			objStream.type = 1
			objStream.open
			objStream.write objMSXML.responseBody
			objStream.savetofile sVanillaPath & "\mods\mods-" & DateModified & ".zip", 2
			objStream.close
	End If
	
	
	
	
	
	
	'Create Game Files
	'-----------------------------------------------------------------------------------------------------

	'Remove old extracted versions of the mods if they exist.
	Dim CurrentVersion
	If FileExists(sVanillaPath & "\mods\unzipped\version.txt") Then
		CurrentVersion = GetExtractedVersion(sVanillaPath & "\mods\unzipped\version.txt")
		If CurrentVersion <> DateModified Then
			Call DeleteFolder(sVanillaPath & "\mods\unzipped")
		End If
	Else
		If FolderExists(sVanillaPath & "\mods\unzipped\") Then Call DeleteFolder (sVanillaPath & "\mods\unzipped\")		'Delete unzipped folder just incase it logically exists when it shouldn't.
	End If
	
	Call CreateFolder(sVanillaPath & "\mods\unzipped")
	If CurrentVersion <> DateModified Then Call UnzipFiles(sVanillaPath & "\mods\mods-" & DateModified & ".zip", sVanillaPath & "\mods\unzipped")		'Unzip the files if they haven't been already.
	Call DebugPrint(sVanillaPath & "\mods\unzipped\version.txt", DateModified)
	Call RoboCopy(sVanillaPath, sVanillaPath & BackupName, "/MIR")				'Game Files.
	Call RoboCopy(sVanillaPath & "\mods\unzipped", sVanillaPath, "/s")   		'Copy mods to the correct locations.
	Call CleanupOldFiles(sVanillaPath & BackupName & "\mods\", DateModified)	'Removes old versions of the mods zip file.
	
	
	
	'Launch Application
	Call OpenFile("steam://run/892970")
	
	
	'Wait for VH to be closed.
	'This restores the vanilla files once the game is closed (as the only way to get the SteamUI to work is to launch via steam."
	Dim vhPID
	WScript.Sleep 15000
	vhPID = getPID("valheim.exe")
	Do While vhPID <> 0
		WScript.Sleep 5000
		vhPID = getPID("valheim.exe")
	Loop
	Call RoboCopy(sVanillaPath & BackupName, sVanillaPath, "/MIR")	'Copy back original files.
	

End Sub






'---------------------------------------------------------------------------------------------
'----- Subprocedures
'---------------------------------------------------------------------------------------------

	Public Sub DebugPrint(outputLocation, whatToPrint)
	Dim objFile: Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(outputLocation,2,True)

		objFile.WriteLine(whatToPrint)
		objFile.Close

	End Sub


	
	Public Sub UnzipFiles(zipPath, extractPath)
	Dim objShellApp:	Set objShellApp = CreateObject("Shell.Application")
	Dim zipFiles:		Set zipFiles = objShellApp.NameSpace(zipPath).Items
	
		objShellApp.NameSpace(extractPath).copyHere zipFiles, 16
		
	End Sub
	
	
	
	'Note: We are cleaning up the backup directory, because robocopy will purge the vanilla directory later when the game exits.
	Public Sub CleanupOldFiles(zipDirectory, CurrentVersion)
	Dim arrFiles: arrFiles = ListFiles(zipDirectory)
	Dim i, o
	
		For i = LBound(arrFiles) To UBound(arrFiles)
			If arrFiles(i) <> "" Then
				If Instr(1, arrFiles(i), CurrentVersion) = 0 Then
					Call DeleteFile(zipDirectory & arrFiles(i))
				End If
			End If
		Next
	
	End Sub
	
	
	
	Public Sub DeleteFile(filePath)
	Dim objFSO:	Set objFSO = CreateObject("Scripting.FileSystemObject")
		objFSO.DeleteFile filePath, True
	End Sub
	
	
	
	Public Sub DeleteFolder(folderPath)
	Dim objFSO:	Set objFSO = CreateObject("Scripting.FileSystemObject")
		objFSO.DeleteFolder folderPath, True
	End Sub
	
	
	
	'Opens a file at the specified path.
	'Call openFile("C:\MyFile.txt")
	Public Sub OpenFile(strFilePath)
	Dim objShellApplication: Set objShellApplication = CreateObject("Shell.Application")
		objShellApplication.Open strFilePath
	End Sub
	

'---------------------------------------------------------------------------------------------
'----- Functions
'---------------------------------------------------------------------------------------------

	'Returns the path that the script is executing from.
	Public Function GetScriptPath()
		GetScriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
	End Function



	Public Function LocateFolder
		Dim steamPath
		
			steamPath = GetSteamInstallPath
			If FileExists(steamPath & "\steamapps\appmanifest_892970.acf") Then
				LocateFolder = steamPath & "\steamapps\common\valheim"
				Exit Function
			End If
			
			
			Dim objFSO:	Set objFSO = CreateObject("Scripting.FileSystemObject")
			Dim objFile: Set objFile = objFSO.OpenTextFile(steamPath & "\steamapps\libraryfolders.vdf",1)
			Dim strLibs
			Dim strLine
				Do Until objFile.AtEndOfStream
					strLine = objFile.ReadLine
					If InStr(1, strLine, "path") > 0 Then strLibs = strLibs & Replace(Replace(Trim(Replace(Replace(Trim(strLine),Chr(34),""),"path","")),"\\","\"),vbTab,"") & "|"
				Loop
		
			Dim arrLibs: arrLibs = Split(strLibs, "|")
			Dim i
			For i = LBound(arrLibs) To UBound(arrLibs)
				If FileExists(arrLibs(i) & "\steamapps\appmanifest_892970.acf") Then
					LocateFolder = arrLibs(i) & "\steamapps\common\valheim"
					Exit For
				End If
			Next
		
		objFile.Close
		If LocateFolder = "" Then LocateFolder = "NOT FOUND"
	End Function



	Public Function GetSteamInstallPath()
		Dim objShell:	Set objShell = CreateObject("Wscript.shell")
			GetSteamInstallPath = Replace(objShell.RegRead("HKEY_CURRENT_USER\SOFTWARE\Valve\Steam\SteamPath"), "/", "\")
	End Function



	'This sub waits for the game client to start and then attaches the steam UI to the PID
	Public Function getPID(exeName)
	Dim objWMI:		Set objWMI = GetObject("winmgmts:!\\.\root\cimv2")
	Dim wmiQuery, process
	Dim PID
	Dim i
	
		'Wait for the process to start and obtain it's Process Id (PID).
		For i = 1 To 20
			
			wmiQuery = "SELECT ProcessId From Win32_Process WHERE Name = '" & exeName & "'"
			For Each process In objWMI.ExecQuery(wmiQuery)
				PID = process.ProcessId
				If PID <> 0 Then
					Exit For
				Else
					WScript.Sleep 1000
				End If
			Next
			
		Next
		
		
	getPID = PID
	End Function
	
	
	
	'Checks if a file exists.
	Public Function FileExists(strPath)
	Dim objFSO:	Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FileExists(strPath) Then FileExists = True
	End Function
	
	
	
	'Checks if a folder exists.
	Public Function FolderExists(strPath)
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(strPath) Then FolderExists = True
	End Function
	
	
	
	'Checks for a folder and creates it if it doesn't exist
	Public Sub CreateFolder(strPath)
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")

		If Not objFSO.FolderExists(strPath) Then objFSO.CreateFolder (strPath)

	End Sub
	
	
	
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
	
	
	
	'This Function Returns a list of folders.
	'Array = ListFiles(objShell.CurrentDirectory & "\Folder\")
	Function ListFiles(sFolder)

		Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
		Dim folder, files, dirLen, arrTemp(), i

		Set folder = objFSO.GetFolder(sFolder)
		Set files = folder.Files

		dirLen = Len(sFolder)
		
		i = 0
		Redim arrTemp(0)
		
		For each x In files
			arrTemp(i) = Right(x,Len(x) - dirLen)
			i = i + 1
			Redim Preserve arrTemp(i)
		Next
	  
		ListFiles = arrTemp
	  
	End Function
	
	
	
	Public Function GetExtractedVersion(txtFilePath)
	Dim objFile: Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(txtFilePath,1)
		GetExtractedVersion = objFile.ReadLine()
	End Function