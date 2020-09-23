VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Page Builder v1.0 - Engine Build 2.1b"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command41 
      Caption         =   "End Data"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6240
      TabIndex        =   64
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Table Data"
      Height          =   255
      Left            =   6240
      TabIndex        =   63
      Top             =   1800
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3615
      Left            =   0
      TabIndex        =   62
      Top             =   5400
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9120
      Top             =   1920
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Page Title"
      Height          =   615
      Left            =   6600
      TabIndex        =   59
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command38 
      Caption         =   "End Row"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   58
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Table Row"
      Height          =   255
      Left            =   4080
      TabIndex        =   57
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command36 
      Caption         =   "End Header"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2040
      TabIndex        =   56
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Table Header"
      Height          =   255
      Left            =   2040
      TabIndex        =   55
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command34 
      Caption         =   "End Table"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   54
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Table"
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   6720
      TabIndex        =   52
      Text            =   "0"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4920
      TabIndex        =   50
      Text            =   "0"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3120
      TabIndex        =   48
      Text            =   "0"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   46
      Text            =   "0"
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Center Justify"
      Height          =   255
      Left            =   5520
      TabIndex        =   44
      Top             =   1080
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Right Justify"
      Height          =   255
      Left            =   3960
      TabIndex        =   43
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Left Justify"
      Height          =   255
      Left            =   2400
      TabIndex        =   42
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Add Item"
      Height          =   615
      Left            =   6360
      TabIndex        =   38
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command31 
      Caption         =   "End UL"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5040
      TabIndex        =   37
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Unordered List"
      Height          =   255
      Left            =   5040
      TabIndex        =   36
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command29 
      Caption         =   "End OL"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   35
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Ordered List"
      Height          =   255
      Left            =   3840
      TabIndex        =   34
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command27 
      Caption         =   "End Underline"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   33
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Underlined Text"
      Height          =   255
      Left            =   2400
      TabIndex        =   32
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command25 
      Caption         =   "End Italics"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   31
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Italic Text"
      Height          =   255
      Left            =   1320
      TabIndex        =   30
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command23 
      Caption         =   "End Bold"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Bold Text"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      Caption         =   "End PRE"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   27
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Preformat"
      Height          =   255
      Left            =   5640
      TabIndex        =   26
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Paragraph"
      Height          =   255
      Left            =   7560
      TabIndex        =   25
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Break"
      Height          =   255
      Left            =   7560
      TabIndex        =   23
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command18 
      Caption         =   "End H3"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "End H2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "End H1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Anchor Tag"
      Height          =   255
      Left            =   7560
      TabIndex        =   19
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command14 
      Caption         =   "End Center"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "End Font"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "End Body"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Header 3"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Header 2"
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Header 1"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Image"
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Center"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Font"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Body"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   975
   End
   Begin VB.VScrollBar VS3 
      Height          =   2775
      Left            =   8520
      Max             =   255
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.VScrollBar VS2 
      Height          =   2775
      Left            =   8040
      Max             =   255
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.VScrollBar VS1 
      Height          =   2775
      Left            =   7560
      Max             =   255
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Comment"
      Height          =   255
      Left            =   7560
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preview"
      Height          =   255
      Left            =   8040
      TabIndex        =   2
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   6975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "XML Editor"
      Height          =   255
      Left            =   5880
      TabIndex        =   67
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "JavaScript Editor"
      Height          =   255
      Left            =   3360
      TabIndex        =   66
      Top             =   9000
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "CSS Editor"
      Height          =   255
      Left            =   960
      TabIndex        =   65
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7080
      TabIndex        =   61
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   60
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      Height          =   255
      Left            =   5520
      TabIndex        =   51
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cell Spacing"
      Height          =   255
      Left            =   3720
      TabIndex        =   49
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cell Padding"
      Height          =   255
      Left            =   1800
      TabIndex        =   47
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cell Border"
      Height          =   255
      Left            =   240
      TabIndex        =   45
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Table Element Alignment"
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Common Table Formatting Tags"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   720
      Width           =   7335
   End
   Begin VB.Line Line12 
      X1              =   7440
      X2              =   7440
      Y1              =   960
      Y2              =   2520
   End
   Begin VB.Line Line11 
      X1              =   120
      X2              =   7440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line10 
      X1              =   120
      X2              =   120
      Y1              =   2520
      Y2              =   960
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   7440
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Common Text Formatting and List Tags"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   2640
      Width           =   7335
   End
   Begin VB.Line Line8 
      X1              =   7440
      X2              =   7440
      Y1              =   2880
      Y2              =   3720
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   7440
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   7440
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   120
      Y1              =   2880
      Y2              =   3720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Common Page Formatting Tags"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3840
      Width           =   7335
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   7440
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   7440
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      X1              =   7440
      X2              =   7440
      Y1              =   4920
      Y2              =   4080
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   4080
      Y2              =   4920
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Color"
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Dim Q As String
Dim Z As String
Dim W As String
Dim X As String
Dim Y As String


Private Sub Command1_Click()
Text1.Text = Text1.Text + Text2.Text + Q
Text2.Text = ""
End Sub

Private Sub Command10_Click()
AddData "<H3>"
Command10.Enabled = 0
Command18.Enabled = 1
End Sub

Private Sub Command11_Click()
AddData "<P>"
End Sub

Private Sub Command12_Click()
AddData "</BODY>"
Command12.Enabled = 0
Command4.Enabled = 1
End Sub

Private Sub Command13_Click()
AddData "</FONT>"
Command5.Enabled = 1
Command13.Enabled = 0
End Sub

Private Sub Command14_Click()
AddData "</CENTER>"
Command6.Enabled = 1
Command14.Enabled = 0
End Sub

Private Sub Command15_Click()
AddData "<A HREF=Add URL Here>Add Text Here</A>"
End Sub

Private Sub Command16_Click()
AddData "</H1>"
Command8.Enabled = 1
Command16.Enabled = 0
End Sub

Private Sub Command17_Click()
AddData "</H2>"
Command9.Enabled = 1
Command17.Enabled = 0
End Sub

Private Sub Command18_Click()
AddData "</H3>"
Command18.Enabled = 0
Command10.Enabled = 1
End Sub

Private Sub Command19_Click()
AddData "<BR>"
End Sub

Private Sub Command2_Click()
Open "asdf.html" For Output As #1
Print #1, Text1.Text + "</HTML>"
Close #1
Form2.WB.Navigate App.Path + "\asdf.html"
Form2.Show
End Sub


Private Sub Command20_Click()
AddData "<PRE>"
Command20.Enabled = 0
Command21.Enabled = 1
End Sub

Private Sub Command21_Click()
AddData "</PRE>"
Command20.Enabled = 1
Command21.Enabled = 0
End Sub

Private Sub Command22_Click()
AddData "<B>"
Command22.Enabled = 0
Command23.Enabled = 1
End Sub

Private Sub Command23_Click()
AddData "</B>"
Command22.Enabled = 1
Command23.Enabled = 0
End Sub

Private Sub Command24_Click()
AddData "<I>"
Command24.Enabled = 0
Command25.Enabled = 1
End Sub

Private Sub Command25_Click()
AddData "</I>"
Command24.Enabled = 1
Command25.Enabled = 0
End Sub

Private Sub Command26_Click()
AddData "<U>"
Command26.Enabled = 0
Command27.Enabled = 1
End Sub

Private Sub Command27_Click()
AddData "</U>"
Command26.Enabled = 1
Command27.Enabled = 0
End Sub

Private Sub Command28_Click()
AddData "<OL>"
Command28.Enabled = 0
Command29.Enabled = 1
End Sub

Private Sub Command29_Click()
AddData "</OL>"
Command28.Enabled = 1
Command29.Enabled = 0
End Sub

Private Sub Command3_Click()
A = InputBox("Enter desired comment, excluding any HTML tags.", "PageBuilder")
Text1.Text = Text1.Text + "<! " + A + ">" + Q
End Sub

Private Sub Command30_Click()
AddData "<UL>"
Command30.Enabled = 0
Command31.Enabled = 1
End Sub

Private Sub Command31_Click()
AddData "</UL>"
Command30.Enabled = 1
Command31.Enabled = 0
End Sub

Private Sub Command32_Click()
A = InputBox("List item to add.  Do not include HTML Tags")
AddData "<LI>" + A
End Sub

Private Sub Command33_Click()
W = "<TABLE"
If Option1.Value = True Then W = W + " ALIGN=LEFT"
If Option2.Value = True Then W = W + " ALIGN=RIGHT"
If Option3.Value = True Then W = W + " ALIGN=CENTER"
If Val(Text3.Text) <> 0 Then W = W + " CELLBORDER=" + Text3.Text
If Val(Text4.Text) <> 0 Then W = W + " CELLPADDING=" + Text4.Text
If Val(Text5.Text) <> 0 Then W = W + " CELLSPACING=" + Text5.Text
If Val(Text6.Text) <> 0 Then W = W + " WIDTH=" + Text6.Text
W = W + ">"
AddData W
Command33.Enabled = 0
Command34.Enabled = 1
End Sub

Private Sub Command34_Click()
AddData "</TABLE>"
Command33.Enabled = 1
Command34.Enabled = 0
End Sub

Private Sub Command35_Click()
AddData "<TH>"
Command35.Enabled = 0
Command36.Enabled = 1
End Sub

Private Sub Command36_Click()
AddData "</TH>"
Command35.Enabled = 1
Command36.Enabled = 0
End Sub

Private Sub Command37_Click()
AddData "<TR>"
Command37.Enabled = 0
Command38.Enabled = 1
End Sub

Private Sub Command38_Click()
AddData "</TR>"
Command37.Enabled = 1
Command38.Enabled = 0
End Sub

Private Sub Command39_Click()
A = InputBox("Please input the title of your page", "PageBuilder")
AddData "<HEAD><TITLE>" + A + "</TITLE></HEAD>"
Form2.Caption = A
End Sub

Private Sub Command4_Click()
W = Hex(VS1.Value)
X = Hex(VS2.Value)
Y = Hex(VS3.Value)
If Len(W) < 2 Then W = "0" + W
If Len(X) < 2 Then X = "0" + X
If Len(Y) < 2 Then Y = "0" + Y
AddData "<BODY BGCOLOR=#" + W + X + Y + ">"
Command4.Enabled = 0
Command12.Enabled = 1
End Sub

Private Sub Command40_Click()
AddData "<TD>"
Command40.Enabled = 0
Command41.Enabled = 1
End Sub

Private Sub Command41_Click()
AddData "</TD>"
Command40.Enabled = 1
Command41.Enabled = 0
End Sub

Private Sub Command5_Click()
W = Hex(VS1.Value)
X = Hex(VS2.Value)
Y = Hex(VS3.Value)
If Len(W) < 2 Then W = "0" + W
If Len(X) < 2 Then X = "0" + X
If Len(Y) < 2 Then Y = "0" + Y
AddData "<FONT COLOR=#" + W + X + Y + ">"
Command5.Enabled = 0
Command13.Enabled = 1
End Sub

Private Sub Command6_Click()
AddData "<CENTER>"
Command6.Enabled = 0
Command14.Enabled = 1
End Sub

Private Sub Command7_Click()
AddData "<IMG SRC=Source>"
End Sub

Private Sub Command8_Click()
AddData "<H1>"
Command8.Enabled = 0
Command16.Enabled = 1
End Sub

Private Sub Command9_Click()
AddData "<H2>"
Command9.Enabled = 0
Command17.Enabled = 1
End Sub

Private Sub Form_Load()
Label1.BackColor = RGB(0, 0, 0)
Q = vbCrLf
Text1.Text = "<HTML>" + vbCrLf
Image1.Picture = LoadPicture("pblogo.JPG")
End Sub

Private Sub Label11_Click()
Form3.Show
End Sub

Private Sub Label14_Click()
Form4.Show
End Sub

Private Sub Label15_Click()
Form5.Show
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.Text = Text1.Text + Text2.Text + vbCrLf
Text2.Text = ""
End If
End Sub

Private Sub Timer1_Timer()
Label12.Caption = Time$
Label13.Caption = Len(Text1.Text)
Label13.Caption = Label13.Caption + " Bytes"
End Sub

Private Sub VS1_Change()
Label1.BackColor = RGB(VS1.Value, VS2.Value, VS3.Value)
End Sub

Private Sub VS2_Change()
Label1.BackColor = RGB(VS1.Value, VS2.Value, VS3.Value)
End Sub

Private Sub VS3_Change()
Label1.BackColor = RGB(VS1.Value, VS2.Value, VS3.Value)
End Sub


Public Sub AddData(Z As String)
Text1.Text = Text1.Text + Z + Q
End Sub
