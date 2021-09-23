VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valheim-Cloud Server Client"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9630
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnReload 
      BackColor       =   &H008080FF&
      Caption         =   "Reload Files"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton btnLaunch 
      BackColor       =   &H00FFFF80&
      Caption         =   "Launch"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   1080
   End
   Begin VB.TextBox txtVanilla 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   9255
   End
   Begin VB.Label Label2 
      Caption         =   "Please ensure that the following path is completely vanilla."
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Vanilla Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   9240
      Width           =   1095
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1931
      _cy             =   1296
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************
'***    Form Load / Unload Events (Timers & Set Defaults)
'*************************************************************************************************
Private bAllowLoad
Private bFormReady As Boolean


    Private Sub Form_Initialize()
        
        
        'Sets the values of all the semiconstants that are in use.
        Call SetDefaultValues
        
        
        'Check for mods directory.
        If Not FolderExists(App.Path & "\Mods") Then
            Call CreateFolder(App.Path, "\Mods")
            MsgBox "Mods folder has been created, place your mods into the folder based on their position to the root game folder and launch the app again", vbOKOnly, "Setup"
        Else
            bAllowLoad = True
        End If
        
        
    End Sub


    Private Sub Form_Load()

        'Unload form if being prevented from loading.
        If Not bAllowLoad Then
            Unload Me
            Exit Sub
        End If
  
        txtVanilla.text = cVanilla_Path
        If cVanilla_Path <> "NOT FOUND" Then Timer1.Interval = 250
        bFormReady = True
    
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Call WriteConfigFile
    End Sub
    
    
    
    Private Sub Timer1_Timer()
    
        If Timer1.Interval <> 0 Then
        'Create Modded VH directory if it doesn't already exist.
            Timer1.Interval = 0
            frmUpdate.Show vbModal
        End If
    
    End Sub
    

    Private Sub SetDefaultValues()
        
        Call ReadConfigFile
        
        'Vanilla Path
        If Not FolderExists(cVanilla_Path) Then cVanilla_Path = FindInstallDir(":\Program Files (x86)\Steam\steamapps\common\Valheim")
        
        'Modded Path
        If Not FolderExists(cModded_Path) Then cModded_Path = cVanilla_Path & " " & App.EXEName
        
    End Sub








    

'********************************
'***    GUI Element Events
'*************************************************************************************************



    Private Sub btnLaunch_Click()
        
        'Invalid Game Path.
        If Not FileExists(cModded_Path & "\Valheim.exe") Then
            MsgBox "Game path is still invalid, sort that out first.", vbOKOnly, "Bad Game Path"
            Exit Sub
        End If
        
        
        'Launch the game.
        Call OpenFile(cModded_Path & "\Valheim.exe")
        Unload Me
        
    End Sub
    
    
    Private Sub btnReload_Click()

        'Delete game folder if it exists.
        If FolderExists(cVanilla_Path & " " & App.EXEName) Then
            Call DeleteFolder(cVanilla_Path & " " & App.EXEName, True)
        End If
        
        frmUpdate.Show vbModal

    End Sub




    Private Sub txtVanilla_Change()
        If Not bFormReady Then Exit Sub
        If FileExists(txtVanilla.text & "\Valheim.exe") Then
            cVanilla_Path = txtVanilla.text
            cModded_Path = txtVanilla.text & " " & App.EXEName
            frmUpdate.Show vbModal
        End If
    End Sub
