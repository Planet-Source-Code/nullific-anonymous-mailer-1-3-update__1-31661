VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6225
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblLoad 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label lblProg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anonymous Mailer 1.3"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   6735
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nullify"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   24
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   6735
   End
   Begin VB.Image imgNull 
      Height          =   6500
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6750
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Pause(Duration As Double)
Dim Current As Long
Current = Timer
Do Until Timer - Current >= Duration
DoEvents
Loop
End Sub

Private Sub Form_Load()
Me.Show
Pause 5
frmMain.Show
Pause 0.5
Me.Hide
End Sub
