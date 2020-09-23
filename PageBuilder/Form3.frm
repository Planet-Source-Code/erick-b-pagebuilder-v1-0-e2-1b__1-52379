VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form3 
   Caption         =   "PageBuilder v1.0 - CSS Editor"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   LinkTopic       =   "Form3"
   ScaleHeight     =   7440
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List6 
      Height          =   7080
      ItemData        =   "Form3.frx":0000
      Left            =   8520
      List            =   "Form3.frx":0002
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List5 
      Height          =   3375
      ItemData        =   "Form3.frx":0004
      Left            =   6960
      List            =   "Form3.frx":0006
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List4 
      Height          =   3375
      ItemData        =   "Form3.frx":0008
      Left            =   5400
      List            =   "Form3.frx":000A
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List3 
      Height          =   3375
      ItemData        =   "Form3.frx":000C
      Left            =   3960
      List            =   "Form3.frx":000E
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   3375
      ItemData        =   "Form3.frx":0010
      Left            =   2160
      List            =   "Form3.frx":0012
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "Form3.frx":0014
      Left            =   0
      List            =   "Form3.frx":0016
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form3.frx":0018
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "HTML Tags"
      Height          =   255
      Left            =   8520
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Miscellaneous"
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "List-Margin-Text"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Font"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Border"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Background"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Dim B As String

Private Sub Form_Load()
Open "CSSLIB\HTML.txt" For Input As #1
List6.Clear
Do While Not EOF(1)
Line Input #1, A
List6.AddItem A
Loop
Close #1
Open "CSSLIB\CSS.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, A
If InStr(A, "background") Then
B = "bg"
ElseIf InStr(A, "border") Then
B = "bd"
ElseIf InStr(A, "font") Then
B = "ft"
ElseIf InStr(A, "list") Or InStr(A, "margin") Or InStr(A, "text") Then
B = "lt"
ElseIf (InStr(A, ":") < 1) Then B = "ms"
End If
If B = "ms" Then List5.AddItem A
If B = "lt" Then List4.AddItem A
If B = "ft" Then List3.AddItem A
If B = "bd" Then List2.AddItem A
If B = "bg" Then List1.AddItem A
Loop
Close #1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UBound(Split(Text1.Text, "{")) > UBound(Split(Text1.Text, "}")) Then Text1.Text = Text1.Text + "}"
If Text1.Text <> "" Then
Form1.Text1.Text = Form1.Text1.Text + "<STYLE TYPE=text/css>" + vbCrLf + Text1.Text + vbCrLf + "</STYLE>"
End If
End Sub

Private Sub List1_Click()
AddIt List1.List(List1.ListIndex)
End Sub

Private Sub List2_Click()
AddIt List2.List(List2.ListIndex)
End Sub

Private Sub List3_Click()
AddIt List3.List(List3.ListIndex)
End Sub

Private Sub List4_Click()
AddIt List4.List(List4.ListIndex)
End Sub

Private Sub List5_Click()
AddIt List5.List(List5.ListIndex)
End Sub

Private Sub List6_Click()
If Text1.Text <> "" Then Text1.Text = Text1.Text + "}" + vbCrLf
Text1.Text = Text1.Text + List6.List(List6.ListIndex) + vbCrLf + "{" + vbCrLf
End Sub


Public Sub AddIt(Z As String)
If InStr(Z, ":") Then

If InStr(Z, ">") Then
MsgBox 1
Dim A As String
A = InputBox("Enter " + Mid(Z, 3, Len(Z) - 3) + " value", Form1.Caption, Mid(Z, 2, Len(Z) - 1))
Z = ":" + A
End If
Text1.Text = Text1.Text + Z + ";" + vbCrLf
Else
Text1.Text = Text1.Text + Z
End If
End Sub



