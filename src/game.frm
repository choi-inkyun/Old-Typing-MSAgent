VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  '없음
   Caption         =   "일정관리"
   ClientHeight    =   5220
   ClientLeft      =   150
   ClientTop       =   1620
   ClientWidth     =   2955
   DrawStyle       =   1  '대시
   Icon            =   "game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Visible         =   0   'False
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   3600
      TabIndex        =   16
      Text            =   "01"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Caption         =   "감추기"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1250
      Left            =   360
      Top             =   5040
   End
   Begin VB.Timer Timer2 
      Interval        =   2100
      Left            =   720
      Top             =   5040
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1080
      Top             =   5040
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   1440
      Top             =   5040
   End
   Begin VB.Timer Timer5 
      Interval        =   2500
      Left            =   1800
      Top             =   5040
   End
   Begin VB.Timer Timer6 
      Interval        =   1000
      Left            =   2160
      Top             =   5040
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  '평면
      Height          =   310
      Left            =   320
      TabIndex        =   5
      Top             =   3840
      Width           =   2380
   End
   Begin VB.OptionButton Option2 
      Height          =   180
      Left            =   1560
      TabIndex        =   4
      Top             =   3070
      Width           =   180
   End
   Begin VB.OptionButton Option1 
      Height          =   180
      Left            =   480
      TabIndex        =   3
      Top             =   3070
      Value           =   -1  'True
      Width           =   180
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "game.frx":27A2
      Left            =   360
      List            =   "game.frx":27AC
      TabIndex        =   6
      Text            =   "오전"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1850
      TabIndex        =   2
      Top             =   2400
      Width           =   375
   End
   Begin MCI.MMControl midcall 
      Height          =   330
      Left            =   -600
      TabIndex        =   9
      Top             =   -120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl midend 
      Height          =   330
      Left            =   -480
      TabIndex        =   10
      Top             =   -120
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      Height          =   495
      Left            =   2520
      TabIndex        =   17
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   2480
      Picture         =   "game.frx":27BC
      Top             =   1710
      Width           =   525
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   2400
      Top             =   2880
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Height          =   495
      Left            =   960
      MousePointer    =   2  '십자형
      TabIndex        =   15
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Height          =   315
      Left            =   1320
      MousePointer    =   2  '십자형
      TabIndex        =   12
      Top             =   1755
      Width           =   990
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Height          =   300
      Left            =   180
      MousePointer    =   2  '십자형
      TabIndex        =   11
      Top             =   1750
      Width           =   970
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "시"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1620
      TabIndex        =   8
      Top             =   2450
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2330
      TabIndex        =   7
      Top             =   2450
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "바탕체"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   0
      Picture         =   "game.frx":3730
      Top             =   0
      Width           =   3000
   End
   Begin VB.Menu Menu 
      Caption         =   "메뉴"
      Begin VB.Menu aksemsdl 
         Caption         =   "만든이"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu mutime 
         Caption         =   "일정관리"
         Shortcut        =   ^T
      End
      Begin VB.Menu muxkwk 
         Caption         =   "타자"
         Shortcut        =   ^W
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
      Begin VB.Menu mutlrksdkfflackd 
         Caption         =   "시간알림창"
         Shortcut        =   ^C
      End
      Begin VB.Menu muhelp 
         Caption         =   "도움"
         Shortcut        =   ^H
      End
      Begin VB.Menu muhideshow 
         Caption         =   "숨기 - 나오기"
      End
      Begin VB.Menu Line3 
         Caption         =   "-"
      End
      Begin VB.Menu muend 
         Caption         =   "끝내기"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ss As Integer
Dim Button, i
Dim z, X As Integer
Dim Y As String
Dim l, t As Integer
Dim Genie As Object

Private Sub Agent1_Show(ByVal CharacterID As String, ByVal Cause As Integer)
Genie.Top = lasttop
Genie.Left = lastleft

End Sub

Private Sub aksemsdl_Click()
Genie.Stop
    Dim LeStyle As Long
    
    
    If Genie.Visible Then
        LeStyle = StyleEnFonction(lastleft + Genie.Width / 2, lasttop + Genie.Height / 2)
        If LeStyle = 0 Then
            Form2.Top = lasttop * Screen.TwipsPerPixelY - Form2.Height + 150
            Form2.Left = lastleft * Screen.TwipsPerPixelX - Form2.Width + Genie.Width * 5
        End If
        If LeStyle = 1 Then
            Form2.Top = lasttop * Screen.TwipsPerPixelY - Form2.Height + 150
            Form2.Left = lastleft * Screen.TwipsPerPixelX + Genie.Width * 10
        End If

        If LeStyle = 4 Then
            Form2.Top = lasttop * Screen.TwipsPerPixelY + Genie.Height * 15
            Form2.Left = lastleft * Screen.TwipsPerPixelX + Genie.Width * 10
        End If

        If LeStyle = 5 Then
            Form2.Top = lasttop * Screen.TwipsPerPixelY + Genie.Height * 15
            Form2.Left = lastleft * Screen.TwipsPerPixelX - Form2.Width + Genie.Width * 5
        End If
    Else
        Form2.Left = (Screen.Width - Form2.Width) / 2
        Form2.Top = (Screen.Height - Form2.Height) / 2

    End If
    Form2.asBubbleForm1.Style = LeStyle
    Genie.Balloon.Style = 0
    Form2.Visible = True
    Form2.asBubbleForm1.CreateRegion
    Form2.Show
End Sub



Private Sub Form_Load()
Image2.Visible = False
Label8.Visible = False
midcall.Visible = False
Form1.Visible = False
midend.Visible = False
Menu.Visible = False
Agent1.Characters.Load "Genie", "Genie.acs" ' Genie 캐릭터 로드
Set Genie = Agent1.Characters("Genie")
Genie.SoundEffectsOn = True
lastleft = Screen.Width \ 15 - Genie.Width - 20
lasttop = Screen.Height \ 15 - Genie.Height - 20
Genie.Top = lasttop
Genie.Left = lastleft
Genie.LanguageID = &H412
Genie.Show
Genie.Play "Greet"
Genie.Speak "부르셨습니까, 주인님!."
Genie.AutoPopupMenu = False
Timer3.Enabled = True
midend.FileName = App.Path + "\End.wav"
midend.Command = "open"
Timer1.Enabled = False
Timer2.Enabled = False
Timer5.Enabled = False




    Dim LeStyle As Long
    
    
    If Genie.Visible Then
        LeStyle = StyleEnFonction(lastleft + Genie.Width / 2, lasttop + Genie.Height / 2)
        If LeStyle = 0 Then
            Form4.Top = lasttop * Screen.TwipsPerPixelY - Form4.Height + 150
            Form4.Left = lastleft * Screen.TwipsPerPixelX - Form4.Width + Genie.Width * 5
        End If
        If LeStyle = 1 Then
            Form4.Top = lasttop * Screen.TwipsPerPixelY - Form4.Height + 150
            Form4.Left = lastleft * Screen.TwipsPerPixelX + Genie.Width * 10
        End If

        If LeStyle = 4 Then
            Form4.Top = lasttop * Screen.TwipsPerPixelY + Genie.Height * 15
            Form4.Left = lastleft * Screen.TwipsPerPixelX + Genie.Width * 10
        End If

        If LeStyle = 5 Then
            Form4.Top = lasttop * Screen.TwipsPerPixelY + Genie.Height * 15
            Form4.Left = lastleft * Screen.TwipsPerPixelX - Form4.Width + Genie.Width * 5
        End If
    Else
        Form4.Left = (Screen.Width - Form4.Width) / 2
        Form4.Top = (Screen.Height - Form4.Height) / 2

    End If
    Form4.asBubbleForm1.Style = LeStyle
    Genie.Balloon.Style = 0
    Form4.Visible = True
    Form4.asBubbleForm1.CreateRegion
    Form4.Show

End Sub

Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer) ' 에이전트를 클릭하면..
If Button = 2 Then
muhideshow.Caption = IIf(Genie.Visible, "숨기", "보이기")
PopupMenu Form1.Menu
End If
End Sub

Private Sub Label4_Click()
Form3.Show
Form1.Hide


End Sub

Private Sub Label5_Click()
   ss = MsgBox("종료하시겠습니까?", vbCritical + vbYesNo, "종료")
    
    If ss = vbYes Then
    Genie.Stop
    Timer1.Enabled = True
    midend.Command = "play"
    Genie.Play "GetAttentionContinued"
        
    Else
        Exit Sub
    End If

End Sub

Private Sub Label6_Click()
If Check2.Value = 0 Then
Check2.Value = 1
Image2.Visible = True
Label8.Visible = True
End If


End Sub

Private Sub Label7_Click()
Genie.Stop
z = Val(Text1.Text)
X = Val(Text2.Text)
Y = Text3.Text

If z > 12 Then
Genie.Speak "시간은 0~12 사이의 숫자를 사용해주세요"
Text1.Text = ""
ElseIf z < 0 Then
Genie.Speak "시간은 0~12 사이의 숫자를 사용해주세요"
Text1.Text = ""
ElseIf Text1.Text = "" Then
Genie.Speak "시간을 입력해 주세요"
ElseIf X > 59 Then
Genie.Speak "분은 0~59 사이의 숫자를 사용해 주세요"
Text2.Text = ""
ElseIf X < 0 Then
Genie.Speak "분은 0~59 사이의 숫자를 사용해 주세요"
Text2.Text = ""
ElseIf Text2.Text = "" Then
Genie.Speak "분을 입력해 주세요"
Else
Form3.Show
Form1.Hide
End If
If Option2.Value = True Then
If Text3.Text = "" Then
Text3.Text = "메세지 없음"
End If
End If

End Sub

Private Sub Label8_Click()
If Check2.Value = 1 Then
Check2.Value = 0
Image2.Visible = False
Label8.Visible = False

End If

End Sub

Private Sub muend_Click()
    
    ss = MsgBox("종료하시겠습니까?", vbCritical + vbYesNo, "종료")
    
    If ss = vbYes Then
    Genie.Stop
    Timer1.Enabled = True
    midend.Command = "play"
    Genie.Play "GetAttentionContinued"
        
    Else
        Exit Sub
    End If

End Sub

Private Sub mugame1_Click()

End Sub

Private Sub muhelp_Click()
Form1.Visible = False
Form1.Visible = True
Form1.Agent1.Characters("Genie").Stop
Form1.Agent1.Characters("Genie").MoveTo 350, 300
Genie.Play "GestureLeft"
Genie.Speak "이쪽을 봐주세요"
Genie.Speak "현재 시각을 나타내고 있습니다"
Genie.MoveTo 670, 370
Genie.Play "GestureRight"
Genie.Speak "다시 이쪽을 보시면 화살표모양이 있습니다."
Genie.Speak "이것을 클릭하시면 시간을 나타내어 주는부분만 남기고 숨겨지게됩니다."
Genie.MoveTo 370, 370
Genie.Play "GestureLeft"
Genie.Speak "이쪽에 있는 알람정보를 클릭하시면"
Genie.Speak "지금까지 설정한 알람들에 대한 정보를 알수있습니다"
Genie.MoveTo 370, 450
Genie.Play "GestureLeft"
Genie.Speak "이곳에서는 알람을 설정하는 기능을 합니다."
Genie.Speak "오전오후를 선택하고시 시간 분을 선택하세요"
Genie.Speak "그리고 음성으로 할지 소리로 하실지 정하시고 음성으로 하시면 내용을 입력해 주세요"
Genie.MoveTo 390, 570
Genie.Play "GestureLeft"
Genie.Speak "그리고 여기있는 설정버튼을 누르시면 됩니다."
Genie.Speak "이상으로 설명을 마치겠습니다."
Genie.Play "Alert"
End Sub

Private Sub muhideshow_Click()


    If Genie.Visible Then
        Genie.Hide
    Else
        Genie.Show
 End If
End Sub

Private Sub mutime_Click()
Genie.Stop
Genie.Play "Alert"
Genie.Speak "일정관리 프로그램을 실행합니다."
Timer5.Enabled = True
End Sub

Private Sub mutlrksdkfflackd_Click()
Form5.Show
End Sub

Private Sub muxkwk_Click()
Genie.Stop
Genie.Play "Alert"
Genie.Speak "타자프로그램을 실행합니다."
Timer2.Enabled = True
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Timer1_Timer()
End
End Sub

Private Function StyleEnFonction(ByVal X As Long, ByVal Y As Long) As Long
    StyleEnFonction = 0
    If X < Screen.Width / 30 Then StyleEnFonction = StyleEnFonction + 1
    If Y < Screen.Height / 30 Then StyleEnFonction = 5 - StyleEnFonction
End Function

Private Sub Timer2_Timer()
Call Shell(App.Path & "\인균이의 타자연습.exe", 1)
Timer2.Enabled = False
End
End Sub

Private Sub Timer3_Timer()
Genie.Play "Alert"
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
If Check2.Value = 1 Then
Form1.Height = 2500
Else
Form1.Height = 5600
End If
If Option1.Value = True Then
Text3.Enabled = False
ElseIf Option2.Value = True Then
Text3.Enabled = True
End If
End Sub

Private Sub Timer5_Timer()
Form6.Show
Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
If Form1.Visible = False Then
muhelp.Enabled = False
End If
If Form1.Visible = True Then
muhelp.Enabled = True
End If
Label1.Caption = CStr(Time)
For i = 0 To 9
If CStr(Time) = 오전오후(i) & " " & 시간(i) & ":" & 분(i) & ":" & Text4.Text Then
If Form3.Label2(i).Caption <> "소리" Then
Genie.Speak 메세지(i)
MsgBox 메세지(i)
End If
If Form3.Label2(i).Caption = "소리" Then
midcall.FileName = App.Path + "\sound\Call.wav"
midcall.Command = "close"
midcall.Command = "open"
midcall.Command = "play"
MsgBox "알람 시각 입니다"
End If
Form3.Label1(i).Caption = ""
Form3.Label2(i).Caption = ""
End If
Next

End Sub

