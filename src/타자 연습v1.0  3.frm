VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  '없음
   Caption         =   "짧은글"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "타자 연습v1.0  3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   360
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   430
      Left            =   3480
      TabIndex        =   23
      Top             =   2910
      Width           =   5840
   End
   Begin MCI.MMControl mid 
      Height          =   375
      Left            =   4920
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   9120
      MousePointer    =   2  '십자형
      TabIndex        =   39
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   9000
      MousePointer    =   2  '십자형
      TabIndex        =   38
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   960
      TabIndex        =   36
      Top             =   4200
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "확인 :"
      Height          =   375
      Left            =   960
      TabIndex        =   35
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '투명
      Caption         =   "최대타수 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   34
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   33
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '투명
      Caption         =   "타수 :"
      Height          =   495
      Left            =   960
      TabIndex        =   32
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Caption         =   "0"
      Height          =   495
      Left            =   1680
      TabIndex        =   31
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Caption         =   "확인 :"
      Height          =   495
      Left            =   960
      TabIndex        =   30
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   29
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '투명
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   28
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   27
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "맞춘갯수 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   26
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '투명
      Caption         =   "0 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   25
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3480
      TabIndex        =   24
      Top             =   2280
      Width           =   5850
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   3600
      TabIndex        =   22
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   3840
      TabIndex        =   21
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   4080
      TabIndex        =   20
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   4320
      TabIndex        =   19
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   4560
      TabIndex        =   18
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   4800
      TabIndex        =   17
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   5040
      TabIndex        =   16
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   5280
      TabIndex        =   15
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   5520
      TabIndex        =   14
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   5760
      TabIndex        =   13
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   6000
      TabIndex        =   12
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   6240
      TabIndex        =   11
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   6480
      TabIndex        =   10
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   6720
      TabIndex        =   9
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   6960
      TabIndex        =   8
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   7200
      TabIndex        =   7
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   16
      Left            =   7440
      TabIndex        =   6
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   17
      Left            =   7680
      TabIndex        =   5
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   18
      Left            =   7920
      TabIndex        =   4
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   19
      Left            =   8160
      TabIndex        =   3
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   20
      Left            =   8400
      TabIndex        =   2
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   21
      Left            =   8640
      TabIndex        =   1
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   22
      Left            =   8880
      TabIndex        =   0
      Top             =   3720
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "타자 연습v1.0  3.frx":030A
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 문장(100), v, k, c, d
Dim cnt As Byte


Private Sub Form_Load()
Open App.Path & "\txt\타자문장.txt" For Input As #1
For i = 1 To 100
    Input #1, 문장(i)
Next
Close #1

a = Int(Rnd(1) * 100)
Label6.Caption = 문장(a)

Label6.Visible = False
k = Len(문장(a))
v = Text1.Text
For i = 0 To k
Label1(i).Visible = False
Timer3.Enabled = False
Next
Timer2.Enabled = False
Label13.Caption = 0
mid.FileName = App.Path + "\THEME ME.wav"
mid.Command = "open"
End Sub

Private Sub Form_Unload(Cancel As Integer)
mumain.Agent1.Characters("Genie").Stop

End Sub

Private Sub Label15_Click()
mumain.Show
Unload Me
mumain.Agent1.Characters("Genie").Stop
End Sub

Private Sub Label16_Click()

Text1.IMEMode = vbIMEModeHangul '한글


Label6.Visible = True
Timer1.Enabled = True
Text1.SetFocus
For i = 0 To k
Label1(i).Caption = "x"
Label1(i).FontSize = Text1.FontSize
Label1(i).Visible = True
Timer2.Enabled = False
Next
Timer3.Enabled = True
cnt = 0
If mumain.Option12 = True Then
mumain.Agent1.Characters("Genie").Speak Label6.Caption
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Timer2.Enabled = True
Label2.Caption = Text1.Text
If KeyAscii = 13 Then
If Trim(Text1.Text) <> "" Then
If cnt <> 0 Then
Select Case cnt
Case 1 To 2
c = 2
Case 3 To 4
c = 1.9
Case 5 To 6
c = 1.8
Case 6 To 7
c = 1.7
Case 8 To 9
c = 1.6
Case 10 To 11
c = 1.5
Case 12 To 13
c = 1.4
Case 14 To 15
c = 1.3
Case 16 To 17
c = 1.2
Case 18 To 19
c = 1.1
Case 20 To 21
c = 1
End Select
Label10.Caption = Round(((k / cnt) * 60) * c * -1) * -1
If Label12.Caption < Label10.Caption Then
Label12.Caption = Label10.Caption
mid.Command = "prev"
mid.Command = "play"
End If

Label13.Caption = Label13.Caption + 1
        For i = 0 To k
        Label1(i).Caption = ""
        Next
       ' Label10.Caption = Int((d / k) * 0.1)
        If Label6.Caption = Trim(Text1.Text) Then
           Label4.Caption = Label4.Caption + 1
           a = Int(Rnd(1) * 100)
           Label6.Caption = 문장(a)
           If Label6.Caption = "" Then
           a = Int(Rnd(1) * 100)
           Label6.Caption = 문장(a)
           End If
           k = Len(문장(a))
           For i = 0 To k
           Label1(i).Caption = "x"
           Next
           Label8.Caption = "맞았다"
        ElseIf v <> Label6.Caption Then
         Label8.Caption = "틀렸다"
           a = Int(Rnd(1) * 100)
           Label6.Caption = 문장(a)
           If Label6.Caption = "" Then
           a = Int(Rnd(1) * 100)
           Label6.Caption = 문장(a)
          End If
           k = Len(문장(a))
           
                      For i = 0 To k
           Label1(i).Caption = "x"
                       
            Beep
           Next
        End If
        End If
    Text1.Text = ""
    Text1.SetFocus
    cnt = 0
End If
If mumain.Option12 = True Then
mumain.Agent1.Characters("Genie").Speak Label6.Caption
End If

End If
End Sub

Private Sub Timer1_Timer()

k = Len(Label6.Caption)

For i = 0 To k
    If Left(Text1.Text, i + 1) = Left(Label6.Caption, i + 1) Then
        Label1(i).Caption = "o"
        End If
         Next
For i = 0 To k
If Left(Text1.Text, i + 1) <> Left(Label6.Caption, i + 1) Then
       Label1(i).Caption = "x"
        End If
        Next
End Sub

Private Sub Timer2_Timer()
'd = Timer2.Interval + Timer2.Interval + 0.1
End Sub

Private Sub Timer3_Timer()
cnt = cnt + 1
End Sub
