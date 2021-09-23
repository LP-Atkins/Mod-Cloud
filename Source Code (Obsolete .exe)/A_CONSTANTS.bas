Attribute VB_Name = "A_CONSTANTS"
'This sub contains all the constants used.
'This also contains the subs for reading and writing config files.



'Constants:     Note some of these are set by the Sub SetConstants()
'--------------------------------------------------------------------
Public CONFIG_FOLDER As String
Public CONFIG_FOLDER_PATH As String
Public Const CONFIG_FILE As String = "config.ini"
Public CONFIG_FILE_PATH As String
'



'Config Constants.
'--------------------------------------------------------------------
Public MuteMusic As Boolean
Public cVanilla_Path As String
Public cModded_Path As String
Public cModded_Version As Double
'


'This sets all the constants that must be set at runtime, should be the first sub that Main() calls.
Sub SetConstants()
    
    CONFIG_FOLDER = App.EXEName
    CONFIG_FOLDER_PATH = Environ("TEMP") & "\" & CONFIG_FOLDER
    CONFIG_FILE_PATH = CONFIG_FOLDER_PATH & "\" & CONFIG_FILE

End Sub



'All config settings go here.
Sub ParseConfigLine(ByRef strLine As String)


    'Determine the delimiter.
    Dim Delimiter As String
    If InStr(1, strLine, ":") > 0 Then Delimiter = ":"
    If InStr(1, strLine, "=") > 0 Then Delimiter = "="

    'Split the variable name and the set to value.
    Dim arrSplit As Variant
    arrSplit = Split(strLine, Delimiter)
    Dim i As Integer
    For i = LBound(arrSplit) To UBound(arrSplit)    'Remove white space.
        arrSplit(i) = Trim(arrSplit(i))
    Next i


    'Select case based on the variable name.
    'Each program variable must appear in this select statement.
    Select Case UCase(arrSplit(0))
        Case UCase("MuteMusic"):          MuteMusic = CBool(arrSplit(1))
        Case UCase("cVanilla_Path"):   cVanilla_Path = CStr(arrSplit(1))
        Case UCase("cModded_Path"):    cModded_Path = CStr(arrSplit(1))
        Case UCase("cModded_Version"):    cModded_Version = CStr(arrSplit(1))
    End Select
    
    

End Sub







Sub ReadConfigFile()

    'Check that the config folder exists, and make it if it doesn't.
    If Not FolderExists(Environ("Temp") & "\") Then Call CreateFolder(Environ("Temp") & "\", App.EXEName)


    'Read & Parse Config file.
    If FileExists(CONFIG_FILE_PATH) Then
    
            
        Dim strLine As String
        Open CONFIG_FILE_PATH For Input As #1
        Do While Not EOF(1)
            Line Input #1, strLine
            Call ParseConfigLine(strLine)
        Loop
        Close #1
        
    End If


End Sub




Sub WriteConfigFile()

    'Check that the config folder exists, and make it if it doesn't.
    If Not FolderExists(Environ("Temp") & "\" & App.EXEName) Then Call CreateFolder(Environ("Temp") & "\", App.EXEName)

    Open CONFIG_FILE_PATH For Output As #1
    
        Print #1, "This config file stores all the user settings. Syntax is Varname = Value."
        Print #1, "One line per variable only!"
        Print #1, ""
        Print #1, "cVanilla_Path = " & cVanilla_Path
        Print #1, "cModded_Path = " & cModded_Path
        Print #1, "cModded_Version = " & cModded_Version
        
    Close #1

End Sub






