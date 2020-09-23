VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Nullify's Anonymous Mailer 1.3"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5294.118
   ScaleMode       =   0  'User
   ScaleWidth      =   5633.539
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar prgProg 
      Height          =   255
      Left            =   2880
      TabIndex        =   32
      Top             =   4200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "About"
      Height          =   255
      Left            =   2280
      TabIndex        =   28
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdDIS 
      Caption         =   "Disconnect"
      Height          =   255
      Left            =   960
      TabIndex        =   26
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdCon 
      Caption         =   "Connect"
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Info 
      Caption         =   "Your Info"
      Height          =   1455
      Left            =   2880
      TabIndex        =   18
      Top             =   2760
      Width           =   2415
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Caption         =   "dd/mm/yy"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblList 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "373 Servers"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblBR 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblGBR 
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes Recieved:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblHN 
         Alignment       =   2  'Center
         Caption         =   "HN"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblGETHN 
         BackStyle       =   0  'Transparent
         Caption         =   "HostName:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblIP 
         Alignment       =   2  'Center
         Caption         =   "0.0.0.0"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblGETIP 
         BackStyle       =   0  'Transparent
         Caption         =   "IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame STAT 
      Caption         =   "Status"
      Height          =   1455
      Left            =   2880
      TabIndex        =   15
      Top             =   0
      Width           =   2415
      Begin VB.TextBox txtStat 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Options 
      Caption         =   "Options"
      Height          =   1335
      Left            =   2880
      TabIndex        =   9
      Top             =   1440
      Width           =   2415
      Begin VB.CheckBox chkIP 
         Caption         =   "127.0.0.1"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Recommended"
         Top             =   840
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkME 
         Caption         =   "My IP"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox cmbServer 
         Height          =   315
         ItemData        =   "frmMain.frx":030A
         Left            =   120
         List            =   "frmMain.frx":076D
         TabIndex        =   10
         Text            =   "SMTP Server"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblMyip 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0.0.0"
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblHELO 
         Alignment       =   2  'Center
         Caption         =   "HELO IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   0
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Time 
      Left            =   0
      Top             =   1680
   End
   Begin VB.Frame Main 
      Caption         =   "Main"
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtBody 
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtSub 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblBody 
         Alignment       =   2  'Center
         Caption         =   "Body:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label lblSubject 
         Caption         =   "Subject:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblTo 
         Caption         =   "To:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblFrom 
         Caption         =   "From:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      Caption         =   "Nullify's Anonymous Mailer 1.3"
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   4560
      Width           =   5415
   End
   Begin VB.Label lblDisC 
      Caption         =   "Disclaimer: The author of this program is not responsible for your actions."
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   4080
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmed by Nullify2002
'There are currently 373 SMTP Servers to choose from.
'I am not sure how i got the addresses of these servers.
'I apologize for some of the code's quality it has been months since i last
'programmed in Visual BASIC.
'Known problems,
'You MUST leave the Pause's, if you dont the commands are sent to fast and may not work
'or the status will not change.
'Some people are reporting that when they click connect it says its connected but when they
'try to send the mail it says they arnt connected to a server. Ionno if this is fixed yet so
'please tell me if you have this problem and if you do then remove a sub and tell me what happens
'If you find any other bugs then e-mail me
'Delphi10000@aol.com
'I added like 2 more things so im just gonna call his version "1.3"
'"Bytes Recieved" bug fixed thnx to nspyrou@yahoo.com
'Check for updates
'Hope you like it, vote, Bai
Private Connected As Boolean 'This defines "connected" it will be either True or False, see WS Connect.
Private IP As String 'Defines "IP"

Public Sub Pause(Duration As Double) 'I needed this Sub to make intervals between the sending of commands
Dim Current As Long 'Duration ican be change (i mean the amout of time)
Current = Timer
Do Until Timer - Current >= Duration 'Loops event until the current time matches the Duration defined
DoEvents
Loop
End Sub

Private Sub cmdCon_Click()
If cmbServer.Text = "SMTP Server" Then 'if the text of cmbServer is still "SMTP Server" then it will show a message box and Exit the sub
MsgBox "I think you should select a server before connecting to one.", vbCritical, "Anonymous Mailer"
Exit Sub 'Exits sub to prevent it from continuing to connect with no server
End If
If cmbServer.Text = "" Then 'Same as the top one, but this time if the text of cmbServer is nothing ("") it will exit the sub
MsgBox "I think you should select a server before connecting to one.", vbCritical, "Anonymous Mailer"
Exit Sub
End If
'If the text of cmbServer is somthing other then "SMTP Server" or nothing it will do the following:
WS.Close 'this closes Winsock incase it was already open.
WS.Connect cmbServer, 25 'connected to the server(cmbServer's text) on the defined port, in this case 25
Pause 2 'Waits 2 seconds so winsock can connect before the next event
lblProgress.Caption = "Connected to: " & WS.RemoteHost 'changes the caption of lblProgrss and shows the IP of the remote host
If Connected = False Then
MsgBox "Connection problem try again.", vbCritical, "Anonymous Mailer"
End Sub 'Ends Sub

Private Sub cmdDIS_Click() 'when cmdDIS is clicked
WS.Close 'Closes winsock the disconnected it from the server
lblProgress.Caption = "Disconnected" 'changes the caption of lblProgress
End Sub

Private Sub cmdHelp_Click() 'when the user clicks cmd send...
frmAbout.Show 'it will show frmAbout
End Sub 'ends sub

Private Sub cmdSend_Click() 'when cmdSend is clicked...
'this is the longest sub
Dim Tfrom As String, Tto As String, Tsub As String, Tbody As String 'these are all variables...
Tfrom = txtFrom.Text 'sets Tfrom
Tto = txtTo.Text 'sets Tto
Tsub = txtSub.Text 'sets Tsub
Tbody = txtBody.Text 'sets Tbody
If chkIP.Value = Checked Then 'if chkIP is checked then "IP" is set to "127.0.0.1"
IP = "127.0.0.1" 'sets "IP"
End If 'if chkIP isnt checked it will continue...
If chkME.Value = Checked Then 'if chkME is checked IP is set to your IP
IP = WS.LocalIP 'Sets IP to your IP, "WS.LocalIP" tells your IP
End If
If Connected = False Then 'If you havnt connected it will display a Message Box...
MsgBox "You Havnt Connected to a Server Yet!", vbCritical, "Anonymous Mailer" 'MSGBox shows the message box, then you put a message and "vbCritical" is what shows that picture, "Anonymous Mailer" is its title
Exit Sub 'if there is no connection the sub will exit
End If 'continues, and o yeah if you are have the connection problem then remover the message box in lines 69-72 and then tell me what happens.
If chkME.Value = Checked And chkIP.Value = Checked Then 'if both chkME and chkIP are checked it will show that message box
MsgBox "You Can Only Choose One HELO Address!", vbCritical, "Anonymous Mailer"
Exit Sub 'Exits the sub if both are checked
End If
If txtFrom.Text = "" Then 'if there is no text it txtFrom the following Message Box will be displayed
MsgBox "Did you want to fake your e-mail address?", vbCritical, "Anonymous Mailer"
Exit Sub 'exits sub if txtfrom is blank
End If
If txtTo.Text = "" Then ''if there is no text it txtTo the following Message Box will be displayed
MsgBox "So, you didnt want anyone to actually RECEIVE this e-mail?", vbCritical, "Anonymous Mailer"
Exit Sub 'Exits sub if txtTo is blank
End If
If txtBody.Text = "" Then 'if there is no text it txtBody the following Message Box will be displayed
MsgBox "Umm, did you want to say anything?", vbCritical, "Anonymous Mailer"
Exit Sub 'exits sub if txtBody is blank
End If
'Now it starts to send commands to the server, see "SMTP Commands.txt" for a list of the commands and about SMTP Servers
WS.SendData "HELO " & IP & Chr(13) & Chr(10) 'Send the command "HELO " and whatever IP is set to, to the server
lblProgress.Caption = "Saying HELO..." 'changes the caption of lblProgress
prgProg.Value = 25 'now sets the progress bar to 25%
Pause 1 'pauses for a second
WS.SendData "MAIL FROM: " & "<" & Tfrom & ">" & Chr(13) & Chr(10) 'sends the e-mail address
lblProgress.Caption = "Sending 1 of 3..." 'changes the caption of lblProgress
prgProg.Value = prgProg.Value + 25 'adds another 25% to the progress bar (you could just put prgProg.Value=50 though)
Pause 1 'Pauses for a second
WS.SendData "RCPT TO: " & "<" & Tto & ">" & Chr(13) & Chr(10) 'tells the server who will be receiving the e-mail
lblProgress.Caption = "Sending 2 of 3..." 'changes the caption of lblProgress
prgProg.Value = prgProg.Value + 25 'adds another 25% to prgProg
Pause 1 'pauses for a second
WS.SendData "DATA " & Chr(13) & Chr(10) 'starts sending Data
lblProgress.Caption = "Sending 3 of 3..." 'changes the caption of lblProgress
WS.SendData "SUBJECT: " & Tsub & Chr(13) & Chr(10) 'specifies the subject of the e-mail
WS.SendData "Importance: high" & Chr(13) & Chr(10) 'Sets Importance(this will be customizible in newer version)
WS.SendData "MIME-Version: 1.0" & Chr(13) & Chr(10) 'Gives MIME Version
WS.SendData "X-Mailer: Nullify's Anonymous Mailer 1.3" 'Gives mail clients name
WS.SendData "Content-Type: Text/HTML; charset=us=ascii" & Chr(13) & Chr(10) 'Gives content type
WS.SendData "Contect-Transfer-Encoding: 7bit" & Chr(13) & Chr(10) 'gives encoding (this will be customizible in newer version)
WS.SendData Tbody & Chr(13) & Chr(10) 'send the body of the e-mail
prgProg.Value = prgProg.Value + 25 'adds yet another 25% to prgProgress making it 100%
Pause 1 'pauses for a second
lblProgress.Caption = "Ending..." 'changes the caption of lblProgress
WS.SendData "." & Chr(13) & Chr(10) 'ends the sending of Data
Pause 1 'Pauses for a second
lblProgress.Caption = "Exiting..." 'changes the caption of lblProgress
Pause 0.5 'Pauses for half a second
WS.SendData "QUIT" & vbCrLf 'Exits server
Pause 1 'Pauses for a second
lblProgress.Caption = "DONE" 'changes the caption of lblProgress
prgProg.Value = 0 'resets prgProg
End Sub 'finally ends the sub

Private Sub Form_Load() 'when the application loads it will do the following...
lblIP.Caption = WS.LocalIP 'change the caption of lblIP to your IP
lblMyip.Caption = WS.LocalIP 'change the caption of lblMyIp to your IP also
lblHN.Caption = WS.LocalHostName 'changes the caption of lblHN to your HostName using WS.LocalHostName
lblDate.Caption = Date 'changes the caption of lblDate to the current date in dd/mm/yy fromat
End Sub ' ends sub

Private Sub txtStat_Change()
txtStat.SelStart = Len(txtStat.Text) 'Scrolls to the last line of text in txtStat
End Sub 'ends sub

Private Sub WS_Connect() 'when Winsock connects
txtStat.Text = "Conneted To: " & WS.RemoteHost & vbCrLf 'changes the text of txtstat to "Connected To: " and displays the IP of the server.
Connected = True 'It is connected to somthing therefor Connected will be true.
End Sub 'ends sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long) 'bytes recieved counter fixed thnx to nspyrou@yahoo.com
'If WS.BytesReceived <> 0 Then ...
    If WS.BytesReceived <> 0 Then
'add the total bytes received to the caption of lblBR
        lblBR.Caption = lblBR.Caption + bytesTotal
    End If

Dim Data As String 'defines data
WS.GetData Data 'gets data
LogText Data 'logs data on txtStat
End Sub 'ends sub

Sub LogText(Text As String) 'Now when "LogText" is typed it will display the text in txtStat w/o using "txtStat.text"
txtStat.Text = txtStat.Text & Text & Chr(13) & Chr(10) 'the text of txtStat will be the same when new Text is added, "Text" is the new text and it will be added to txtStat
End Sub 'Ends sub
