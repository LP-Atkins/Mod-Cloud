Attribute VB_Name = "mdlVarious"
'This module contains all of the various subs & functions that are called by the forms.
'


'Finds the installed folder directory:
Public Function FindInstallDir(strInstallDir As String) As String
Dim i As Integer
Dim bFound As Integer


    '64 Bit Windows.
    For i = 3 To 26
        If FolderExists(ConvertToLetter(i) & strInstallDir) Then
            FindInstallDir = ConvertToLetter(i) & strInstallDir
            bFound = True
            Exit For
        End If
    Next i

If FindInstallDir = "" Then FindInstallDir = "NOT FOUND"
End Function





'Uses the FileSystemObject to copy a specified folder to a destination.
'Example Usage:
'Call CopyFolder("C:\FolderA", "C:\FolderB", True)
Public Sub CopyFolder(ByVal source As String, _
                    ByVal destination As String, _
                    ByVal overwrite As Boolean)
Dim objFSO:    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Call objFSO.CopyFolder(source, destination, True)

End Sub



'Uses the FileSystemObject to delete a specified folder.
'Example Usage:
'Call DeleteFolder("C:\FolderA", True)
Public Sub DeleteFolder(ByVal folderspec As String, _
                        ByVal force As Boolean)
Dim objFSO:    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Call objFSO.DeleteFolder(folderspec, force)
                        
End Sub



'Executes a robocopy command & string using the shell object.
'Example Usage: (will copy the contents of test1 to the folder of test2)
'Call Robocopy("""c:\test"" ""c:\test2"" /MIR")     '/MIR will purge the destination, so remove it if not required.
Public Sub RoboCopy(strCmd)
Dim objShell:   Set objShell = CreateObject("Wscript.Shell")

    objShell.Run "robocopy " & strCmd

End Sub



'Returns True or False if a input folder exists or not
'Call FolderExists("C:\Folder")
Public Function FolderExists(FolderPath As String) As Boolean
On Error Resume Next
    
    If FolderPath = "" Then Exit Function

    If Dir(FolderPath, vbDirectory) <> "" Then
        FolderExists = True
        If Err.Number <> 0 Then FolderExists = False
    Else
        FolderExists = False
    End If

End Function




'Checks for a folder and creates it if it doesn't exist
'Call CreateFolder("C:\", "New Folder")
Public Sub CreateFolder(FolderPath As String, FolderName As String)

    Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject") 'Global file system object

    If Dir(FolderPath & FolderName & "\", vbDirectory) = vbNullString Then
        objFSO.CreateFolder (FolderPath & FolderName)
    End If

End Sub




'Returns True or False if a input file exists or not
'Call FileExists("C:\File.xtn")
Public Function FileExists(FilePath As String) As Boolean

    If Dir(FilePath) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If

End Function



'Opens a file at the specified path.
'Call openFile("C:\MyFile.txt")
Public Sub OpenFile(strFilePath)

    Dim objShellApplication: Set objShellApplication = CreateObject("Shell.Application")
    objShellApplication.Open strFilePath

End Sub




'Converts a number to a letter
'https://support.microsoft.com/en-au/help/833402/how-to-convert-excel-column-numbers-into-alphabetical-characters
'Call ConvertToLetter(1)
Public Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function
