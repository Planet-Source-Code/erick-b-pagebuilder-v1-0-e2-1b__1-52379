VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form4 
   Caption         =   "JavaScript Editor"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form4"
   ScaleHeight     =   8895
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Enter"
      Height          =   495
      Left            =   9360
      TabIndex        =   13
      Top             =   3720
      Width           =   975
   End
   Begin RichTextLib.RichTextBox TXJ 
      Height          =   4575
      Left            =   0
      TabIndex        =   12
      Top             =   4320
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8070
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form4.frx":0000
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   3960
      Width           =   9255
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Index           =   4
      ItemData        =   "Form4.frx":0082
      Left            =   9840
      List            =   "Form4.frx":0084
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Index           =   3
      ItemData        =   "Form4.frx":0086
      Left            =   7920
      List            =   "Form4.frx":0088
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Index           =   2
      ItemData        =   "Form4.frx":008A
      Left            =   4440
      List            =   "Form4.frx":008C
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Index           =   1
      ItemData        =   "Form4.frx":008E
      Left            =   3000
      List            =   "Form4.frx":0090
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Index           =   0
      ItemData        =   "Form4.frx":0092
      Left            =   0
      List            =   "Form4.frx":0094
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   4
      Left            =   9840
      TabIndex        =   11
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   10
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   9
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   8
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Custom Line"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3720
      Width           =   9255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Dim B As Integer
Dim C As String

Private Sub Command1_Click()
TXJ.Text = TXJ.Text + Text1.Text + ";" + vbCrLf
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Form_Load()
B = 0
A = Dir("JSLIB\*.txt", vbNormal)
Do While A <> ""
Open "JSLIB\" + A For Input As #1
Do While Not EOF(1)
Line Input #1, C
List1(B).AddItem C
DoEvents
Loop
Close #1
Label2(B).Caption = Mid(A, 1, Len(A) - 4)
A = Dir()
B = B + 1
Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Text1.Text = Form1.Text1.Text + "<SCRIPT LANGUAGE=JavaScript>" + vbCrLf + TXJ.Text + "</SCRIPT>" + vbCrLf
End Sub

Private Sub List1_Click(Index As Integer)
A = List1(Index).List(List1(Index).ListIndex)
If (Not InStr(A, ".")) And Index <> 4 Then
TXJ.Text = TXJ.Text + A + ";" + vbCrLf
Else
TXJ.Text = TXJ.Text + A
End If
End Sub
