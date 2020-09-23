VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4230
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   495
         Left            =   2865
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Select Sound"
         Height          =   495
         Left            =   2865
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "Set Alarm Time"
         Height          =   495
         Left            =   885
         TabIndex        =   2
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000D&
         Caption         =   "Exit"
         Height          =   495
         Left            =   4845
         TabIndex        =   1
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1095
         Left            =   630
         TabIndex        =   5
         Top             =   1545
         Width           =   5655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1095
         Left            =   630
         TabIndex        =   4
         Top             =   510
         Width           =   5655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1095
         Left            =   630
         TabIndex        =   3
         Top             =   2625
         Width           =   5655
      End
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   4335
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   "C:\WINDOWS\Media\ringin.wav"
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   10
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdStop_Click()
MediaPlayer1.Stop
End Sub

Private Sub Command1_Click()
frmAlarm.Show
Me.Hide
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
frmAS.Show
Me.Hide
End Sub

Private Sub Form_Load()
strTimey = (Time)
Label1.Caption = strTimey
End Sub

Private Sub tmrTime_Timer()
strTimey = (Time)
Label1.Caption = strTimey
If strTimey = strAlarm Then
MediaPlayer1.Play
End If
tmrTime.Interval = 1000
tmrTime.Enabled = True
End Sub

