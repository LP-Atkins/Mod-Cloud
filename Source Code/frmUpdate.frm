VERSION 5.00
Begin VB.Form frmUpdate 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3960
      Top             =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while files are being created and patched."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()

    Timer1.Interval = 0
    If Not FolderExists(cModded_Path) Then Call CopyFolder(cVanilla_Path, cModded_Path, True)
    Unload Me
    
    Call RoboCopy(Chr(34) & App.Path & "\Mods" & Chr(34) & " " & Chr(34) & cModded_Path & Chr(34) & " /s")   'Uses Robocopy to copy mods to the correct locations.
    
    'Enable buttons.
    frmMain.btnLaunch.Visible = True
    frmMain.btnReload.Visible = True
    
End Sub
