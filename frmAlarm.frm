VERSION 5.00
Begin VB.Form frmAlarm 
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAlarm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK!"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtAlarm 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   1980
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   97
      Width           =   7080
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Input in this order  :    Hour: Minute: Second  AM/PM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   840
         TabIndex        =   5
         Top             =   2400
         Width           =   5460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Input Alarm Time:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Left            =   600
         TabIndex        =   3
         Top             =   1800
         Width           =   2220
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Height          =   3135
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
strAlarm = txtAlarm.Text
Unload Me
frmMain.Show
End Sub
