VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  '없음
   Caption         =   "자리연습"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   Icon            =   "타자 연습 V1.0 6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   3840
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "한글연습"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "영어연습"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   9120
      MousePointer    =   2  '십자형
      TabIndex        =   10
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   9000
      MousePointer    =   2  '십자형
      TabIndex        =   9
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   9360
      MousePointer    =   2  '십자형
      TabIndex        =   7
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "맞은갯수 :"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "한글자리 연습을 할때에도 영어로 하고 입력해 주세요^-^"
      Height          =   975
      Left            =   840
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   8640
      MousePointer    =   2  '십자형
      TabIndex        =   1
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   48
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   5760
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "타자 연습 V1.0 6.frx":030A
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 한글자리(26), a, 영어자리(27)


Private Sub Form_Load()
Randomize
Label1.Visible = False
            If mumain.Option12 = True Then
               mumain.Agent1.Characters("Genie").Speak Label5.Caption
             End If

Label1.Visible = False
한글자리(1) = "ㄱ"
한글자리(2) = "ㄴ"
한글자리(3) = "ㄷ"
한글자리(4) = "ㄹ"
한글자리(5) = "ㅁ"
한글자리(6) = "ㅂ"
한글자리(7) = "ㅅ"
한글자리(8) = "ㅇ"
한글자리(9) = "ㅈ"
한글자리(10) = "ㅊ"
한글자리(11) = "ㅌ"
한글자리(12) = "ㅋ"
한글자리(13) = "ㅍ"
한글자리(26) = "ㅎ"
한글자리(15) = "ㅛ"
한글자리(16) = "ㅕ"
한글자리(17) = "ㅑ"
한글자리(18) = "ㅐ"
한글자리(19) = "ㅔ"
한글자리(20) = "ㅗ"
한글자리(21) = "ㅓ"
한글자리(22) = "ㅏ"
한글자리(23) = "ㅣ"
한글자리(24) = "ㅠ"
한글자리(25) = "ㅜ"
한글자리(26) = "ㅡ"
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Command3.Enabled = False
Command4.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
mumain.Agent1.Characters("Genie").Stop

End Sub

Private Sub Label7_Click()
mumain.Show
Unload Me
mumain.Agent1.Characters("Genie").Stop
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Label8_Click()
Label1.Visible = True
Label4.Caption = 0
Text1.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Command3.Enabled = False Then
If Label1.Caption = "ㄱ" Then
If KeyCode = vbKeyR Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㄴ" Then
If KeyCode = vbKeyS Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㄷ" Then
If KeyCode = vbKeyE Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㄹ" Then
If KeyCode = vbKeyF Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅁ" Then
If KeyCode = vbKeyA Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅂ" Then
If KeyCode = vbKeyQ Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅅ" Then
If KeyCode = vbKeyT Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅇ" Then
If KeyCode = vbKeyD Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅈ" Then
If KeyCode = vbKeyW Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅊ" Then
If KeyCode = vbKeyC Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅌ" Then
If KeyCode = vbKeyX Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅋ" Then
If KeyCode = vbKeyZ Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅍ" Then
If KeyCode = vbKeyV Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅎ" Then
If KeyCode = vbKeyG Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "ㅛ" Then
If KeyCode = vbKeyY Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "ㅕ" Then
If KeyCode = vbKeyU Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅑ" Then
If KeyCode = vbKeyI Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅐ" Then
If KeyCode = vbKeyO Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅔ" Then
If KeyCode = vbKeyP Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅗ" Then
If KeyCode = vbKeyH Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅓ" Then
If KeyCode = vbKeyJ Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅏ" Then
If KeyCode = vbKeyK Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅣ" Then
If KeyCode = vbKeyL Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅠ" Then
If KeyCode = vbKeyB Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅜ" Then
If KeyCode = vbKeyN Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "ㅡ" Then
If KeyCode = vbKeyM Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If




End If

If Command4.Enabled = False Then
If Label1.Caption = "A" Then
If KeyCode = vbKeyA Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "B" Then
If KeyCode = vbKeyB Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "C" Then
If KeyCode = vbKeyC Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If


If Label1.Caption = "D" Then
If KeyCode = vbKeyD Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "E" Then
If KeyCode = vbKeyE Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "F" Then
If KeyCode = vbKeyF Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "G" Then
If KeyCode = vbKeyG Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "H" Then
If KeyCode = vbKeyH Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "I" Then
If KeyCode = vbKeyI Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "J" Then
If KeyCode = vbKeyJ Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "K" Then
If KeyCode = vbKeyK Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "L" Then
If KeyCode = vbKeyL Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "M" Then
If KeyCode = vbKeyM Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "N" Then
If KeyCode = vbKeyN Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "O" Then
If KeyCode = vbKeyO Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "P" Then
If KeyCode = vbKeyP Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "Q" Then
If KeyCode = vbKeyQ Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "R" Then
If KeyCode = vbKeyR Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "S" Then
If KeyCode = vbKeyS Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "T" Then
If KeyCode = vbKeyT Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "U" Then
If KeyCode = vbKeyU Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "V" Then
If KeyCode = vbKeyV Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "W" Then
If KeyCode = vbKeyW Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "X" Then
If KeyCode = vbKeyX Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "Y" Then
If KeyCode = vbKeyY Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "Z" Then
If KeyCode = vbKeyZ Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
End If
If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If

If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If

If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 한글자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = 영어자리(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Label1.Caption = "" Then
If Command3.Enabled = False Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
End If
If Command3.Enabled = False Then
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
End If
If Command4.Enabled = False Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
End If
If Command4.Enabled = False Then
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
End If
End If
End Sub

Private Sub Label2_Click()
Label1.Visible = False
한글자리(1) = "ㄱ"
한글자리(2) = "ㄴ"
한글자리(3) = "ㄷ"
한글자리(4) = "ㄹ"
한글자리(5) = "ㅁ"
한글자리(6) = "ㅂ"
한글자리(7) = "ㅅ"
한글자리(8) = "ㅇ"
한글자리(9) = "ㅈ"
한글자리(10) = "ㅊ"
한글자리(11) = "ㅌ"
한글자리(12) = "ㅋ"
한글자리(13) = "ㅍ"
한글자리(26) = "ㅎ"
한글자리(15) = "ㅛ"
한글자리(16) = "ㅕ"
한글자리(17) = "ㅑ"
한글자리(18) = "ㅐ"
한글자리(19) = "ㅔ"
한글자리(20) = "ㅗ"
한글자리(21) = "ㅓ"
한글자리(22) = "ㅏ"
한글자리(23) = "ㅣ"
한글자리(24) = "ㅠ"
한글자리(25) = "ㅜ"
한글자리(26) = "ㅡ"
a = Int(Rnd(1) * 26)
Label1.Caption = 한글자리(a)
Command3.Enabled = False
Command4.Enabled = True
End Sub

Private Sub Label6_Click()
Text1.IMEMode = vbIMEModeAlpha  '영문
Label1.Visible = False
영어자리(1) = "A"
영어자리(2) = "B"
영어자리(3) = "C"
영어자리(4) = "D"
영어자리(5) = "E"
영어자리(6) = "F"
영어자리(7) = "G"
영어자리(8) = "H"
영어자리(9) = "I"
영어자리(10) = "J"
영어자리(11) = "K"
영어자리(12) = "L"
영어자리(13) = "M"
영어자리(14) = "N"
영어자리(15) = "O"
영어자리(16) = "P"
영어자리(18) = "Q"
영어자리(17) = "R"
영어자리(19) = "S"
영어자리(20) = "T"
영어자리(21) = "U"
영어자리(22) = "V"
영어자리(23) = "W"
영어자리(24) = "U"
영어자리(25) = "X"
영어자리(26) = "Y"
영어자리(27) = "Z"
a = Int(Rnd(1) * 27)
Label1.Caption = 영어자리(a)
Command3.Enabled = True
Command4.Enabled = False
End Sub
