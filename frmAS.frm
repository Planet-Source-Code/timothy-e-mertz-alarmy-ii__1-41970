VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAS 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2055
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   1875
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3585
      Begin VB.CommandButton Command1 
         Caption         =   "Preview Sound"
         Height          =   495
         Left            =   2310
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog dlgSWF 
         Left            =   0
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtPC 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Text            =   "1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSWF 
         Caption         =   "Select Sound"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdRTN 
         Caption         =   "Return"
         Height          =   495
         Left            =   1215
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Times"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         Caption         =   "Repeat:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRTN_Click()
strPC = txtPC.Text
frmMain.MediaPlayer1.PlayCount = strPC
frmMain.Show
Unload Me
End Sub

Private Sub cmdSWF_Click()
Dim sFile As String

    With dlgSWF
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "*.wav; *.mp3; *.wma(*.wav; *.mp3; *.wma)|*.wav; *.mp3; *.wma"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        frmMain.MediaPlayer1.FileName = sFile
    End With
    

End Sub

Private Sub Command1_Click()
frmMain.MediaPlayer1.PlayCount = txtPC.Text
frmMain.MediaPlayer1.Play
End Sub
