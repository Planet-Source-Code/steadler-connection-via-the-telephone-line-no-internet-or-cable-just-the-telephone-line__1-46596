VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2700
      TabIndex        =   5
      Top             =   1530
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   435
      Left            =   1515
      TabIndex        =   4
      Top             =   1530
      Width           =   1065
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1515
      TabIndex        =   2
      Top             =   1005
      Width           =   2310
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1515
      TabIndex        =   0
      Top             =   525
      Width           =   2310
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   315
      Left            =   195
      TabIndex        =   3
      Top             =   990
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Phone number"
      Height          =   315
      Left            =   330
      TabIndex        =   1
      Top             =   525
      Width           =   1155
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
savesettings
Form1.Text2.Text = Text1.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Function savesettings()
Open App.Path & "\set.ini" For Output As 1
Print #1, Text1.Text
Print #1, Text2.Text

Close #1

End Function

Private Function loadsettings()
On Error Resume Next
Open App.Path & "\set.ini" For Input As 1
Line Input #1, a
Text1.Text = a
Line Input #1, a
Text2.Text = a
Close #1
End Function

Private Sub Form_Load()
loadsettings
End Sub
