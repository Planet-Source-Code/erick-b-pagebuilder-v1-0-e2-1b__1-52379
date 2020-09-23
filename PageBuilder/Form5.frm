VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form5 
   Caption         =   "XML Editor"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11295
   LinkTopic       =   "Form5"
   ScaleHeight     =   8895
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser XMB 
      Height          =   4215
      Left            =   0
      TabIndex        =   37
      Top             =   4680
      Width           =   11295
      ExtentX         =   19923
      ExtentY         =   7435
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin RichTextLib.RichTextBox XMLT 
      Height          =   4335
      Left            =   4800
      TabIndex        =   35
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7646
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"Form5.frx":0000
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Index           =   9
      Left            =   3120
      TabIndex        =   34
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   33
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   32
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   31
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   30
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   29
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   28
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   27
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   26
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   25
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   9
      Left            =   1560
      TabIndex        =   24
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   8
      Left            =   1560
      TabIndex        =   23
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   7
      Left            =   1560
      TabIndex        =   22
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   21
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   20
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   19
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   18
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   17
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   16
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   15
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   4080
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   0
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Click Here to preview the XML code"
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   4440
      Width           =   11295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Data"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "End Tag"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Start Tag"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "General Tag"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
XMLT.Text = XMLT.Text + Command1(Index).Caption + vbCrLf
End Sub

Private Sub Command2_Click(Index As Integer)
XMLT.Text = XMLT.Text + Command2(Index).Caption + vbCrLf
End Sub

Private Sub Form_Load()
XMLT.Text = "<?xml version=""" + "1.0" + """?>" + vbCrLf + "<!-- File Name: asdf.xml -->" + vbCrLf
End Sub

Private Sub Label5_Click()
Open "asdf.xml" For Output As #1
Print #1, XMLT.Text
Close #1
XMB.Navigate App.Path + "\asdf.xml"
End Sub

Private Sub Text1_Change(Index As Integer)
Command1(Index).Caption = "<" + Text1(Index).Text + ">"
Command2(Index).Caption = "</" + Text1(Index).Text + ">"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
XMLT.Text = XMLT.Text + Text2.Text
Text2.Text = ""
End If
End Sub
