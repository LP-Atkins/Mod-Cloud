Attribute VB_Name = "a_Main"
'Valheim-Cloud Server Client
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
'v1.00          23/09/2020          LongParsnip             Original Release
'




'Semi Constants
'------------------------------------------------
Public PATH_App As String
'




Public Sub Main()
On Error Resume Next

    Call InitSemiConstants
    
    If App.PrevInstance Then
        AppActivate App.Title
    Else
        Call SetConstants
        frmMain.Show
    End If
    
End Sub



'Initialises all the program semi constants.
Private Sub InitSemiConstants()
    
    PATH_App = App.Path
    
End Sub
