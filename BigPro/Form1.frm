VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "OTE-Link"
   ClientHeight    =   5235
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7920
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1515
      TabIndex        =   5
      Top             =   3675
      Width           =   4065
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3120
      Left            =   1485
      TabIndex        =   4
      Top             =   495
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   5503
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0442
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   -15
      TabIndex        =   3
      Top             =   4650
      Width           =   2070
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   240
      Left            =   2100
      TabIndex        =   2
      Top             =   4680
      Width           =   795
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   -510
      Top             =   5235
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00EB8669&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4950
      Width           =   7920
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   7575
      Picture         =   "Form1.frx":04C6
      Stretch         =   -1  'True
      Top             =   30
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "Dissconnected"
      Height          =   255
      Left            =   6015
      TabIndex        =   1
      Top             =   60
      Width           =   1530
   End
   Begin VB.Menu file 
      Caption         =   "Server/Client"
      Begin VB.Menu st 
         Caption         =   "Start Server/Client"
      End
      Begin VB.Menu v 
         Caption         =   "Stop Server/Client"
      End
   End
   Begin VB.Menu opt 
      Caption         =   "Options"
      Begin VB.Menu cone 
         Caption         =   "Connection"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sto As Boolean
Dim na As String, uname As String
Private Sub Command1_Click()
On Error Resume Next
MSComm1.PortOpen = True
ExCommand "CONNECT|" & Text2.Text
End Sub

Private Sub cone_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Load Form2
Text2.Text = Form2.Text1.Text
na = Form2.Text2.Text
End Sub

Private Sub st_Click()
On Error Resume Next
MSComm1.PortOpen = True
MainLoop
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

Case 13
MSComm1.Output = "MSG|" & na & ":" & Text3.Text & vbCr
Text3.Text = ""
End Select
End Sub

Private Sub v_Click()
On Error Resume Next
MSComm1.PortOpen = False
Command1.Enabled = True
Text1.Text = ""
Label1.Caption = "Dissconnected"
End Sub
Private Function MainLoop()

Do

DoEvents
temp = MSComm1.Input


ExCommand (temp)

Loop

End Function

Private Function ExCommand(com)
DoEvents
Dim strr As String
com = com & "|"
sp = Split(com, "|")
strr = sp(0)

strr = Replace(strr, vbCrLf, "")
strr = Replace(strr, vbCr, "")
strr = Replace(strr, vbLf, "")

Select Case strr


Case "RING"

Text1.Text = "Recieve call"
DoEvents

MSComm1.Output = "ATA" & vbCr
Command1.Enabled = False
Case "CONNECTED"

Text1.Text = "Connected at " & sp(1)
Label1.Caption = Text1.Text

MSComm1.Output = "NAME " & na
Case "CONNECT"

MSComm1.Output = "ATDT" & sp(1) & vbCr
Text1.Text = "Connecting at " & sp(1)
Label1.Caption = "Connecting"
Command1.Enabled = False

Case "NAME"
uname = sp(1)

Case "MSG"

RichTextBox1.Text = RichTextBox1.Text & sp(1)
RichTextBox1.SelStart = Len(RichTextBox1.Text)

End Select
End Function
