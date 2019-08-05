VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form7 
   BorderStyle     =   0  '없음
   Caption         =   "전쟁"
   ClientHeight    =   6000
   ClientLeft      =   105
   ClientTop       =   45
   ClientWidth     =   10500
   Icon            =   "타자 연습 V1.0 7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer17 
      Interval        =   100
      Left            =   5760
      Top             =   0
   End
   Begin VB.CheckBox Check2 
      Height          =   180
      Left            =   7680
      TabIndex        =   35
      Top             =   4640
      Width           =   180
   End
   Begin VB.CheckBox Check1 
      Height          =   180
      Left            =   7680
      TabIndex        =   34
      Top             =   4230
      Value           =   1  '확인
      Width           =   180
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   9000
      TabIndex        =   28
      Top             =   4230
      Width           =   180
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   9000
      TabIndex        =   27
      Top             =   4470
      Value           =   -1  'True
      Width           =   180
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   9000
      TabIndex        =   26
      Top             =   4710
      Width           =   180
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3240
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   4680
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Interval        =   100
      Left            =   5040
      Top             =   0
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   3960
      Top             =   0
   End
   Begin VB.Timer Timer8 
      Interval        =   100
      Left            =   2880
      Top             =   0
   End
   Begin VB.Timer Timer9 
      Interval        =   100
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer10 
      Interval        =   100
      Left            =   3600
      Top             =   0
   End
   Begin VB.Timer Timer11 
      Interval        =   100
      Left            =   4320
      Top             =   0
   End
   Begin VB.Timer Timer12 
      Interval        =   100
      Left            =   2160
      Top             =   0
   End
   Begin VB.Timer Timer13 
      Interval        =   100
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer14 
      Interval        =   100
      Left            =   5400
      Top             =   0
   End
   Begin VB.Timer Timer15 
      Interval        =   100
      Left            =   2520
      Top             =   0
   End
   Begin VB.Timer Timer16 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   3990
      TabIndex        =   2
      Top             =   4550
      Width           =   2700
   End
   Begin MCI.MMControl mid12 
      Height          =   330
      Left            =   5880
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl mid11 
      Height          =   330
      Left            =   5880
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   9000
      MousePointer    =   2  '십자형
      TabIndex        =   33
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   5640
      MousePointer    =   2  '십자형
      TabIndex        =   32
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   4440
      MousePointer    =   2  '십자형
      TabIndex        =   31
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   25
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   24
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   2
      Left            =   8400
      TabIndex        =   23
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   3
      Left            =   8400
      TabIndex        =   22
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   4
      Left            =   8400
      TabIndex        =   21
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   20
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   6
      Left            =   8400
      TabIndex        =   19
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   7
      Left            =   8400
      TabIndex        =   18
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      X1              =   600
      X2              =   10320
      Y1              =   -120
      Y2              =   -120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   8
      Left            =   8400
      TabIndex        =   17
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   9
      Left            =   8400
      TabIndex        =   16
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   10
      Left            =   8400
      TabIndex        =   15
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Index           =   11
      Left            =   8400
      TabIndex        =   14
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '단색
      Height          =   3735
      Left            =   1080
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '투명
      Height          =   420
      Index           =   0
      Left            =   -600
      TabIndex        =   13
      Top             =   960
      Width           =   1020
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '투명
      Height          =   420
      Index           =   1
      Left            =   -600
      TabIndex        =   12
      Top             =   2640
      Width           =   1020
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Caption         =   "적군의 병력"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '투명
      Height          =   975
      Left            =   2040
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "점수 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Caption         =   "확 인 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '투명
      Caption         =   "확 인 :"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "단 어 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   3960
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "타자 연습 V1.0 7.frx":030A
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 낱말(200), k, g, e, p, o, i, u, t, r, m, n, s, q, w, v, f, h, J, l, aa, bb, cc, dd, aaa, bbb, ccc, ddd
Dim b As Integer
Dim qwe
Dim ttt
Dim tttt
Dim yyy
Dim qqq
Dim www
Dim eee
Dim rrr
Dim ytr


Private Sub Check1_Click()
If Check1.Value = 1 Then
Check2.Value = 0
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Check1.Value = 0
End If
End Sub

Private Sub Form_Load()
Open App.Path & "\txt\타자글.txt" For Input As #1
For i = 1 To 200
    Input #1, 낱말(i)
Next
Close #1


a = Int(Rnd(1) * 200)
Label2.Caption = 낱말(a)
Label5.Caption = 0
Label2.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
Timer12.Enabled = False
Timer13.Enabled = False
Timer14.Enabled = False
Timer15.Enabled = False
Timer16.Enabled = False
Label12.Visible = False
mid11.FileName = App.Path + "\sound\THEME MI.wav"
mid11.Command = "open"
mid12.FileName = App.Path + "\sound\THEME AS.wav"
mid12.Command = "open"
For i = 0 To 1
Label13(i).Visible = False
Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
mumain.Agent1.Characters("Genie").Stop

End Sub

Private Sub Label11_Change()
If Check2.Value = 1 Then
        If mumain.Option12 = True Then
        mumain.Agent1.Characters("Genie").Speak Label11.Caption
        End If
End If
End Sub

Private Sub Label14_Click()
Text1.IMEMode = vbIMEModeHangul '한글
Timer1.Enabled = True
Label2.Visible = True
Text1.SetFocus
p = False
q = False
w = False
v = False
f = False
h = False
J = False
l = False
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True
Timer9.Enabled = True
Timer10.Enabled = True
Timer11.Enabled = True
Timer12.Enabled = True
Timer13.Enabled = True
Timer14.Enabled = True
Timer15.Enabled = True
Timer16.Enabled = True
Label8.Visible = True
Label6.Visible = True
Label12.Visible = True
Label11.Caption = "동과 서의 전쟁이 시작되었습니다. 적을 전멸시키는게 목적입니다. 행운을 빌겠습니다."
        If Check1.Value = 1 Then
            If mumain.Option12 = True Then
               mumain.Agent1.Characters("Genie").Speak Label2.Caption
             End If
        End If
End Sub

Private Sub Label15_Click()
mumain.Agent1.Characters("Genie").Stop
Label11.Caption = "전쟁이 다시 발발하려고 합니다"
Shape1.Left = 1080
Timer1.Enabled = False
Label8.Visible = False
Label6.Visible = False
Label5.Visible = 0
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
Timer12.Enabled = False
Timer13.Enabled = False
Timer14.Enabled = False
Timer15.Enabled = False
Timer16.Enabled = False
Label12.Visible = False
Label2.Visible = False
End Sub

Private Sub Label16_Click()
mumain.Show
Unload Me
mumain.Agent1.Characters("Genie").Stop
End Sub

Private Sub Option4_Click()

End Sub

Private Sub Option5_Click()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 32 Then
Label8.Caption = Text1.Text
'Label8.Caption = tirm(Text1.Text)
'p = Text1.Text
        ttt = Len(Text1.Text)
          tttt = ttt * 40
If Trim(Text1.Text) = Label2.Caption Then
    c = Int(Rnd(0) * 11)
    If c = hh Then
        c = Int(Rnd(0) * 11)
            If c = hh Then
            ElseIf c = hh Then
        c = Int(Rnd(0) * 11)
            ElseIf c = hh Then
        c = Int(Rnd(0) * 11)

        c = Int(Rnd(0) * 11)
    If c = hh Then
        c = Int(Rnd(0) * 11)

    ElseIf c = hh Then
        c = Int(Rnd(0) * 11)
    ElseIf c = hh Then
        c = Int(Rnd(0) * 11)
    ElseIf c = hh Then
        c = Int(Rnd(0) * 11)
   End If
   End If
   End If
    hh = c
    s = 0
    o = 1
    i = 2
    u = 3
    t = 4
    r = 5
    n = 6
    m = 7
    aa = 8
    bb = 9
    cc = 10
    dd = 11
    Label1(c).Caption = Label2.Caption
   
    a = Int(Rnd(1) * 200)
    Label2.Caption = 낱말(a)
    If Label2.Caption = "" Then
    a = Int(Rnd(1) * 200)
    Label2.Caption = 낱말(a)
     End If
    g = k * 80
    Label6.Caption = "맞았다"
If c = s Then
p = True
End If
If c = m Then
q = True
End If
If c = n Then
w = True
End If
If c = r Then
v = True
End If
If c = t Then
f = True
End If
If c = u Then
h = True
End If
If c = i Then
J = True
End If
If c = o Then
l = True
End If
If c = aa Then
aaa = True
End If
If c = bb Then
bbb = True
End If
If c = cc Then
ccc = True
End If
If d = dd Then
ddd = True
End If
        Else
If Text1.Text <> "" Then
    yyy = Int(Rnd(0) * 1)
    qqq = 0
    eee = 1
    ytr = yyy
    If yyy = ytr Then
     yyy = Int(Rnd(0) * 1)

    ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
ElseIf yyy = ytr Then
        yyy = Int(Rnd(0) * 1)
  End If
    Label13(yyy).Caption = Text1.Text
    If yyy = qqq Then
    www = True
    End If
    If yyy = eee Then
    rrr = True
    End If
    a = Int(Rnd(1) * 200)
    Label2.Caption = 낱말(a)
    If Label2.Caption = "" Then
     a = Int(Rnd(1) * 200)
     Label2.Caption = 낱말(a)
    End If
    Label6.Caption = "틀렸다"
End If
End If
Text1.Text = ""
Text1.SetFocus
        If Check1.Value = 1 Then
            If mumain.Option12 = True Then
               mumain.Agent1.Characters("Genie").Speak Label2.Caption
             End If
        End If

End If
End Sub

Private Sub Timer1_Timer()

If Option1.Value = True Then
Shape1.Left = Shape1.Left + Int(Rnd(60) * 20)
ElseIf Option2.Value = True Then
Shape1.Left = Shape1.Left + Int(Rnd(40) * 20)
ElseIf Option3.Value = True Then
Shape1.Left = Shape1.Left + Int(Rnd(20) * 5)
Else
Shape1.Left = Shape1.Left + Int(Rnd(25) * 15)
End If
e = Label5.Caption
k = Len(Label2.Caption)

If Shape1.Left >= 8200 Then
MsgBox "전쟁이 끝났습니다"
mumain.Agent1.Characters("Genie").Play "Sad"
MsgBox "당신은 잘 싸워 주었지만 전쟁에서는 패하고 말았습니다." + Chr(13) + Chr(13) _
       + "당신은 적군에게 " & e & " 만큼의 피해를 입혔습니다"
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
Timer12.Enabled = False
Timer13.Enabled = False
Timer14.Enabled = False
Timer15.Enabled = False
Timer16.Enabled = False
Label1(0).Caption = ""
Label1(0).Left = 8400
Label1(1).Caption = ""
Label1(1).Left = 8400
Label1(2).Caption = ""
Label1(2).Left = 8400
Label1(3).Caption = ""
Label1(3).Left = 8400
Label1(4).Caption = ""
Label1(4).Left = 8400
Label1(5).Caption = ""
Label1(5).Left = 8400
Label1(6).Caption = ""
Label1(6).Left = 8400
Label1(7).Caption = ""
Label1(7).Left = 8400
Label1(8).Caption = ""
Label1(8).Left = 8400
Label1(9).Caption = ""
Label1(9).Left = 8400
Label1(10).Caption = ""
Label1(10).Left = 8400
Label1(11).Caption = ""
Label1(11).Left = 8400
Label13(0).Caption = ""
Label13(0).Left = 0
Label13(1).Caption = ""
Label13(1).Left = 0
Label12.Caption = ""
Text1.Text = ""
Shape1.Left = 840
Label5.Caption = 0
Label2.Visible = False
Label11.Caption = "당신은 전쟁에서 패배하였습니다"
End If

If Shape1.Left <= 0 Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
Timer12.Enabled = False
Timer13.Enabled = False
Timer14.Enabled = False
Timer15.Enabled = False
Timer16.Enabled = False
Label1(0).Caption = ""
Label1(0).Left = 8400
Label1(1).Caption = ""
Label1(1).Left = 8400
Label1(2).Caption = ""
Label1(2).Left = 8400
Label1(3).Caption = ""
Label1(3).Left = 8400
Label1(4).Caption = ""
Label1(4).Left = 8400
Label1(5).Caption = ""
Label1(5).Left = 8400
Label1(6).Caption = ""
Label1(6).Left = 8400
Label1(7).Caption = ""
Label1(7).Left = 8400
Label1(8).Caption = ""
Label1(8).Left = 8400
Label1(9).Caption = ""
Label1(9).Left = 8400
Label1(10).Caption = ""
Label1(10).Left = 8400
Label1(11).Caption = ""
Label1(11).Left = 8400
Label13(0).Caption = ""
Label13(0).Left = 0
Label13(1).Caption = ""
Label13(1).Left = 0
Label12.Caption = ""
Text1.Text = ""

Shape1.Left = 840
mumain.Agent1.Characters("Genie").Stop
mumain.Agent1.Characters("Genie").Play "Congratulate"
MsgBox "축하합니다. 당신은 전쟁에서 승리하였습니다" + Chr(13) + Chr(13) _
        + "당신은 " & e & "명의 군사로 승리하셨습니다"
Label5.Caption = 0
Label2.Visible = False
Label11.Caption = "축하합니다. 당신은 전쟁에서 승리하였습니다"
End If
End Sub

Private Sub Timer10_Timer()
Label1(aa).Visible = False

If aaa Then
           If Label1(aa).Left <= 8200 Then
       Label1(aa).Visible = True
        End If
    
    Label1(aa).Left = Label1(aa).Left - 800
    If Label1(aa).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
       mid11.Command = "prev"

       mid11.Command = "play"
               Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
        Label1(aa).Left = 8400
        Label1(aa).Caption = ""
        aaa = False
    End If
End If

End Sub

Private Sub Timer11_Timer()
Label1(bb).Visible = False

If bbb Then
           If Label1(bb).Left <= 8200 Then
       Label1(bb).Visible = True
        End If
    
    Label1(bb).Left = Label1(bb).Left - 800
    If Label1(bb).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
       mid11.Command = "prev"
        mid11.Command = "play"

        Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
        Label1(bb).Left = 8400
        Label1(bb).Caption = ""
        bbb = False
    End If
End If

End Sub

Private Sub Timer12_Timer()
Label1(cc).Visible = False

If ccc Then
           If Label1(cc).Left <= 8200 Then
       Label1(cc).Visible = True
        End If
    
    Label1(cc).Left = Label1(cc).Left - 800
    If Label1(cc).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
        Label1(cc).Left = 8400
        Label1(cc).Caption = ""
        ccc = False
    End If
End If

End Sub

Private Sub Timer13_Timer()
Label1(dd).Visible = False

If ddd Then
           If Label1(dd).Left <= 8200 Then
       Label1(dd).Visible = True
        End If
    
    Label1(dd).Left = Label1(dd).Left - 800
    If Label1(dd).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
        Label1(dd).Left = 8400
        Label1(dd).Caption = ""
        ddd = False
    End If
End If

End Sub

Private Sub Timer14_Timer()
Label12.Caption = Shape1.Left
End Sub

Private Sub Timer15_Timer()
Label13(qqq).Visible = False
If www Then
    Label13(qqq).Visible = True
    Label13(qqq).Left = Label13(qqq).Left + 1000
    If Label13(qqq).Left >= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left + tttt
mid12.Command = "prev"
        
        mid12.Command = "play"

       Label11.Caption = "적군에서 " & tttt & "명의 원군이 도착했습니다"
         Label13(qqq).Caption = Text1.Text
        Shape1.Left = Shape1.Left + tttt
       Label5.Caption = Label5.Caption - tttt
        Label13(qqq).Left = 0
        Label13(qqq).Caption = ""
        www = False
    End If
End If
End Sub

Private Sub Timer16_Timer()
Label13(eee).Visible = False
If rrr Then
    Label13(eee).Visible = True
    Label13(eee).Left = Label13(eee).Left + 1000
    If Label13(eee).Left >= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left + tttt
mid12.Command = "prev"
        
        mid12.Command = "play"

       Label11.Caption = "적군에서 " & tttt & "명의 원군이 도착했습니다"
         Label13(eee).Caption = Text1.Text
        Shape1.Left = Shape1.Left + tttt
       Label5.Caption = Label5.Caption - tttt
        Label13(eee).Left = 0
        Label13(eee).Caption = ""
        rrr = False
    End If
End If

End Sub

Private Sub Timer17_Timer()
If Timer2.Enabled = True Then
Check1.Enabled = False
Check2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Label14.Visible = False
End If
If Timer2.Enabled = False Then
Check1.Enabled = True
Check2.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Label14.Visible = True
End If

End Sub

Private Sub Timer2_Timer()
Label1(o).Visible = False
If l Then
    If Label1(o).Left <= 8200 Then
       Label1(o).Visible = True
       End If
    
    Label1(o).Left = Label1(o).Left - 800
    If Label1(o).Left <= (Shape1.Left + Shape1.Width) Then
      Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

       Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
        Label1(o).Left = 8400
        Label1(o).Caption = ""
        l = False
    End If
End If

End Sub

Private Sub Timer3_Timer()
Label1(i).Visible = False

If J Then
           If Label1(i).Left <= 8200 Then
       Label1(i).Visible = True
      End If
    
    Label1(i).Left = Label1(i).Left - 800
    If Label1(i).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
        Label1(i).Left = 8400
        Label1(i).Caption = ""
        J = False
    End If
End If

End Sub

Private Sub Timer4_Timer()
Label1(u).Visible = False

If h Then
           If Label1(u).Left <= 8200 Then
       Label1(u).Visible = True
        End If
    
    Label1(u).Left = Label1(u).Left - 800
    If Label1(u).Left <= (Shape1.Left + Shape1.Width) Then
        Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
        Label1(u).Left = 8400
        Label1(u).Caption = ""
        h = False
    End If
End If

End Sub

Private Sub Timer5_Timer()
Label1(t).Visible = False

If f Then
           If Label1(t).Left <= 8200 Then
       Label1(t).Visible = True
       End If
    
    Label1(t).Left = Label1(t).Left - 800
    If Label1(t).Left <= (Shape1.Left + Shape1.Width) Then
           Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
        Label1(t).Left = 8400
        Label1(t).Caption = ""
        f = False
    End If
End If

End Sub

Private Sub Timer6_Timer()
Label1(r).Visible = False
If v Then
           If Label1(r).Left <= 8200 Then
       Label1(r).Visible = True
         End If
     
    Label1(r).Left = Label1(r).Left - 800
    If Label1(r).Left <= (Shape1.Left + Shape1.Width) Then
           Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
        Label1(r).Left = 8400
        Label1(r).Caption = ""
        v = False
    End If
End If

End Sub

Private Sub Timer7_Timer()
Label1(n).Visible = False

If w Then
           If Label1(n).Left <= 8200 Then
       Label1(n).Visible = True
       End If
    Label1(n).Left = Label1(n).Left - 800
        If Label1(n).Left <= (Shape1.Left + Shape1.Width) Then
           Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
        Label1(n).Left = 8400
        Label1(n).Caption = ""
        w = False
    End If
End If

End Sub

Private Sub Timer8_Timer()
Label1(m).Visible = False

If q Then
           If Label1(m).Left <= 8200 Then
       Label1(m).Visible = True
        End If
    
    Label1(m).Left = Label1(m).Left - 800
    If Label1(m).Left <= (Shape1.Left + Shape1.Width) Then
           Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
        Label1(m).Left = 8400
        Label1(m).Caption = ""
            q = False
    End If
End If

End Sub

Private Sub Timer9_Timer()
Label1(s).Visible = False

If p Then
                 If Label1(s).Left <= 8200 Then
       Label1(s).Visible = True
        End If
    
    Label1(s).Left = Label1(s).Left - 800
    If Label1(s).Left <= (Shape1.Left + Shape1.Width) Then

       Shape1.Left = Shape1.Left - g
       Label5.Caption = Label5.Caption + g
mid11.Command = "prev"
       
       mid11.Command = "play"

        Label11.Caption = "적군의 병사를 " & g & " 명 만큼 손실시켰습니다"
       Label1(s).Left = 8400
        Label1(s).Caption = ""
        p = False
    End If
End If

End Sub
