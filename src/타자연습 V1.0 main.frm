VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form mumain 
   BackColor       =   &H80000009&
   BorderStyle     =   0  '없음
   Caption         =   "타자 프로그램"
   ClientHeight    =   6030
   ClientLeft      =   105
   ClientTop       =   -180
   ClientWidth     =   10530
   DrawStyle       =   3  '대시-점
   Icon            =   "타자연습 V1.0 main.frx":0000
   LinkTopic       =   "Form2"
   MouseIcon       =   "타자연습 V1.0 main.frx":030A
   ScaleHeight     =   6030
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      Height          =   10095
      Left            =   0
      Picture         =   "타자연습 V1.0 main.frx":0614
      ScaleHeight     =   10035
      ScaleWidth      =   14235
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      Begin VB.CheckBox Check10 
         Caption         =   "Check10"
         Height          =   255
         Left            =   8760
         TabIndex        =   35
         Top             =   4440
         Width           =   180
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Check9"
         Height          =   255
         Left            =   8760
         TabIndex        =   34
         Top             =   3480
         Width           =   180
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check8"
         Height          =   255
         Left            =   8760
         TabIndex        =   33
         Top             =   2520
         Width           =   180
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check7"
         Height          =   255
         Left            =   8760
         TabIndex        =   32
         Top             =   1680
         Width           =   180
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check6"
         Height          =   255
         Left            =   6480
         TabIndex        =   31
         Top             =   4440
         Width           =   180
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check5"
         Height          =   255
         Left            =   6480
         TabIndex        =   30
         Top             =   3480
         Width           =   180
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   255
         Left            =   6480
         TabIndex        =   29
         Top             =   2520
         Width           =   180
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   6480
         TabIndex        =   28
         Top             =   1680
         Value           =   1  '확인
         Width           =   180
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   180
         Left            =   1320
         TabIndex        =   27
         Top             =   3930
         Width           =   180
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   180
         Left            =   1320
         TabIndex        =   26
         Top             =   3420
         Width           =   180
      End
      Begin VB.OptionButton Option13 
         BackColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   1320
         TabIndex        =   25
         Top             =   3720
         Width           =   180
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Option12"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   5040
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Timer Timer12 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer Timer11 
         Interval        =   10
         Left            =   0
         Top             =   1080
      End
      Begin VB.Timer Timer10 
         Interval        =   1000
         Left            =   0
         Top             =   720
      End
      Begin VB.Timer Timer9 
         Interval        =   1000
         Left            =   0
         Top             =   1440
      End
      Begin VB.Timer Timer8 
         Interval        =   1000
         Left            =   0
         Top             =   2520
      End
      Begin VB.Timer Timer7 
         Interval        =   1000
         Left            =   0
         Top             =   3600
      End
      Begin VB.Timer Timer6 
         Interval        =   1000
         Left            =   0
         Top             =   2880
      End
      Begin VB.Timer Timer5 
         Interval        =   1000
         Left            =   0
         Top             =   3240
      End
      Begin VB.Timer Timer4 
         Interval        =   1000
         Left            =   0
         Top             =   1800
      End
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   0
         Top             =   3960
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   0
         Top             =   2160
      End
      Begin MCI.MMControl mid3 
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   1440
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   661
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.Timer Timer1 
         Interval        =   1250
         Left            =   0
         Top             =   360
      End
      Begin MCI.MMControl mid9 
         Height          =   330
         Left            =   600
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid8 
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   661
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid7 
         Height          =   330
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid6 
         Height          =   330
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid5 
         Height          =   330
         Left            =   720
         TabIndex        =   6
         Top             =   -120
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid4 
         Height          =   330
         Left            =   480
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid2 
         Height          =   330
         Left            =   720
         TabIndex        =   8
         Top             =   2160
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MCI.MMControl mid10 
         Height          =   330
         Left            =   3840
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin AgentObjectsCtl.Agent Agent1 
         Left            =   3600
         Top             =   1560
         _cx             =   847
         _cy             =   847
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Height          =   495
         Left            =   9120
         MousePointer    =   2  '십자형
         TabIndex        =   23
         Top             =   5400
         Width           =   975
      End
      Begin VB.Image Image5 
         Height          =   885
         Left            =   8940
         Picture         =   "타자연습 V1.0 main.frx":CD798
         Top             =   5120
         Width           =   1560
      End
      Begin VB.Image Image4 
         Height          =   6000
         Left            =   0
         Picture         =   "타자연습 V1.0 main.frx":D1FC4
         Top             =   0
         Width           =   10500
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   8040
         MousePointer    =   2  '십자형
         TabIndex        =   21
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '투명
         Height          =   735
         Left            =   9600
         MousePointer    =   2  '십자형
         TabIndex        =   20
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   8520
         MousePointer    =   2  '십자형
         TabIndex        =   19
         Top             =   5520
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   8280
         MousePointer    =   2  '십자형
         TabIndex        =   18
         Top             =   4500
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   8280
         MousePointer    =   2  '십자형
         TabIndex        =   17
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   8160
         MousePointer    =   2  '십자형
         TabIndex        =   16
         Top             =   3200
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   8160
         MousePointer    =   2  '십자형
         TabIndex        =   15
         Top             =   2550
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   4110
         Left            =   7520
         Picture         =   "타자연습 V1.0 main.frx":19F148
         Top             =   1895
         Width           =   2985
      End
      Begin VB.Label Label14 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   1080
         MousePointer    =   2  '십자형
         TabIndex        =   14
         Top             =   4290
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   1080
         MousePointer    =   2  '십자형
         TabIndex        =   13
         Top             =   3660
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   1080
         MousePointer    =   2  '십자형
         TabIndex        =   12
         Top             =   3320
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   960
         MousePointer    =   2  '십자형
         TabIndex        =   11
         Top             =   2670
         Width           =   1935
      End
      Begin VB.Image Image2 
         Height          =   4095
         Left            =   0
         Picture         =   "타자연습 V1.0 main.frx":1C73BC
         Top             =   1910
         Width           =   3120
      End
      Begin VB.Image Image1 
         Height          =   765
         Left            =   7830
         Picture         =   "타자연습 V1.0 main.frx":1F0D70
         Top             =   5230
         Width           =   2670
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   4280
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Height          =   210
         Left            =   1080
         TabIndex        =   9
         Top             =   3660
         Width           =   975
      End
   End
End
Attribute VB_Name = "mumain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lastleft As Long
Dim lasttop As Long
Dim Button
Dim Genie As Object
Dim a As Integer
Dim cnt As Byte
Dim CNT1 As Byte
Dim cnt2 As Byte
Dim cnt3 As Byte
Dim cnt4 As Byte
Dim cnt5 As Byte
Dim cnt6 As Byte
Dim cnt7 As Byte
Dim cnt8 As Byte
Dim aaa As Byte
Dim bbb As Byte
Dim ccc As Byte
Dim optiona, optionb

Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
If Button = 2 Then
Form2.muhideshow.Caption = IIf(Genie.Visible, "숨기", "보이기")
PopupMenu Form2.Menu
End If
End Sub


Private Sub Check1_Click()
If Check1.Value = 0 Then
Option13.Value = False
Option12.Value = True
End If
If Check1.Value = 1 Then
Option13.Value = True
Option13.Value = False
End If

End Sub

Private Sub Check2_Click()
If optiona = 2 Then
Check3.Enabled = True
Check4.Enabled = True
Check5.Enabled = True
Check6.Enabled = True
Check7.Enabled = True
Check8.Enabled = True
Check9.Enabled = True
Check10.Enabled = True

End If
If optiona = 1 Then
'If Option6.Value = True Then
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
Check3.Enabled = False
Check4.Enabled = False
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Check8.Enabled = False
Check9.Enabled = False
Check10.Enabled = False
End If

If optiona = 1 Then
optiona = 2
GoTo 1
End If
If optiona = 2 Then
optiona = 1
End If
1:


End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
mid4.Command = "close"
mid5.Command = "close"
mid2.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"


mid6.FileName = App.Path + "\sound\2-02 Ahead on Our Way Midi.mid"
mid6.Command = "open"
mid6.Command = "play"
Check4.Value = 0
Check5.Value = 0
Check3.Value = 0
Check7.Value = 0
Check8.Value = 0
Check9.Value = 0
Check10.Value = 0
End If

End Sub

Private Sub Form_Load()
optionb = 0
optiona = 1
Check6.Visible = False
Check1.Visible = False
Check2.Visible = False
Label1.Visible = False
Image5.Visible = False
Option12.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Agent1.Characters.Load "Genie", "Genie.acs" ' Genie 캐릭터 로드
Set Genie = Agent1.Characters("Genie")
Genie.SoundEffectsOn = True
lastleft = Screen.Width \ 15 - Genie.Width - 20
lasttop = Screen.Height \ 15 - Genie.Height - 20
Genie.Top = lasttop
Genie.Left = lastleft
Genie.LanguageID = &H412
Genie.Show
Genie.AutoPopupMenu = False
Timer1.Enabled = False
Timer3.Enabled = True
Timer4.Enabled = True
Timer2.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True
Timer9.Enabled = True
Timer12.Enabled = True
mid2.Visible = False
mid2.Enabled = False
mid3.Visible = False
mid3.Enabled = False
mid4.Visible = False
mid4.Enabled = False
mid5.Visible = False
mid5.Enabled = False
mid6.Visible = False
mid6.Enabled = False
mid7.Visible = False
mid7.Enabled = False
mid8.Visible = False
mid8.Enabled = False
mid9.Visible = False
mid9.Enabled = False
mid10.Visible = False
mid10.Enabled = False

Check3.Visible = False
Check4.Visible = False
Check5.Visible = False
Check6.Visible = False
Check7.Visible = False
Check8.Visible = False
Check9.Visible = False
Check10.Visible = False

Option13.Visible = False
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
If Check3.Value = True Then
mid2.FileName = App.Path + "\sound\Toheart.mid"
mid2.Command = "open"
mid2.Command = "play"
End If
If Check4.Value = True Then
mid4.FileName = App.Path + "\sound\31.mid"
mid4.Command = "open"
mid4.Command = "play"
End If
If Check3.Value = True Then
mid5.FileName = App.Path + "\sound\11.mid"
mid5.Command = "open"
mid5.Command = "play"
End If
If Check6.Value = True Then
mid6.FileName = App.Path + "\sound\2-02 Ahead on Our Way Midi.mid"
mid6.Command = "open"
mid6.Command = "play"
End If
If Check7.Value = True Then
mid7.FileName = App.Path + "\sound\12.mid"
mid7.Command = "open"
mid7.Command = "play"
End If
If Check8.Value = True Then
mid8.FileName = App.Path + "\sound\34.mid"
mid8.Command = "open"
mid8.Command = "play"
End If
If Check9.Value = True Then
mid9.FileName = App.Path + "\sound\Hero.mid"
mid9.Command = "open"
mid9.Command = "play"
End If
If Check10.Value = True Then
mid10.FileName = App.Path + "\sound\Yestrday.mid"
mid10.Command = "open"
mid10.Command = "play"
End If
mid3.FileName = App.Path + "\sound\End.wav"
mid3.Command = "open"



    Dim LeStyle As Long
    
    
    If Genie.Visible Then
        LeStyle = StyleEnFonction(lastleft + Genie.Width / 2, lasttop + Genie.Height / 2)
        If LeStyle = 0 Then
            Form12.Top = lasttop * Screen.TwipsPerPixelY - Form12.Height + 150
            Form12.Left = lastleft * Screen.TwipsPerPixelX - Form12.Width + Genie.Width * 5
        End If
        If LeStyle = 1 Then
            Form12.Top = lasttop * Screen.TwipsPerPixelY - Form12.Height + 150
            Form12.Left = lastleft * Screen.TwipsPerPixelX + Genie.Width * 10
        End If

        If LeStyle = 4 Then
            Form12.Top = lasttop * Screen.TwipsPerPixelY + Genie.Height * 15
            Form12.Left = lastleft * Screen.TwipsPerPixelX + Genie.Width * 10
        End If

        If LeStyle = 5 Then
            Form12.Top = lasttop * Screen.TwipsPerPixelY + Genie.Height * 15
            Form12.Left = lastleft * Screen.TwipsPerPixelX - Form12.Width + Genie.Width * 5
        End If
    Else
        Form12.Left = (Screen.Width - Form2.Width) / 2
        Form12.Top = (Screen.Height - Form2.Height) / 2

    End If
    Form12.asBubbleForm1.Style = LeStyle
    Genie.Balloon.Style = 0
    Form12.Visible = True
    Form12.asBubbleForm1.CreateRegion
    Form12.Show
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
End Sub

Private Sub Label1_Click()
Label1.Visible = False
Check1.Visible = False
Check2.Visible = False
mid2.Visible = False
mid2.Enabled = False
mid3.Visible = False
mid3.Enabled = False
mid4.Visible = False
mid4.Enabled = False
mid5.Visible = False
mid5.Enabled = False
mid6.Visible = False
mid6.Enabled = False
mid7.Visible = False
mid7.Enabled = False
mid8.Visible = False
mid8.Enabled = False
mid9.Visible = False
mid9.Enabled = False
mid10.Visible = False
mid10.Enabled = False
Image4.Visible = False
Image5.Visible = False
Check7.Visible = False
Check4.Visible = False
Check5.Visible = False
Check6.Visible = False
Check8.Visible = False
Check9.Visible = False
Check10.Visible = False
Check3.Visible = False
Option13.Visible = False
Option12.Visible = False


Check6.Visible = False

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = True
End Sub

Private Sub Label10_Click()
Image4.Visible = True
Check3.Visible = True
Check4.Visible = True
Check5.Visible = True
Check6.Visible = True
Check7.Visible = True
Check8.Visible = True
Check9.Visible = True
Check10.Visible = True
Check1.Visible = True
Check2.Visible = True
Label1.Visible = True
Check6.Visible = True
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\image\t-a-4.bmp")
Image1.Visible = True
End Sub

Private Sub Label11_Click()

     mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"

'mid3.Command = "play"
'Timer1.Enabled = True
End
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\image\t-a-2.bmp")
Image1.Visible = True
End Sub

Private Sub Label12_Click()
          If mumain.Option12 = True Then
               mumain.Agent1.Characters("Genie").Speak "이 프로그램은 누구나 쉽게 이용할수 있게" + Chr(13) + Chr(13) _
       + "만들었습니다.^^ 그래서 도움말은 생략하겠습니다." + Chr(13) + Chr(13) _
       + "궁금하신점이 있으시면 메일로 보내주세요." + Chr(13) + Chr(13) _

End If

End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\image\t-a-3.bmp")
Image1.Visible = True
End Sub

Private Sub Label13_Click()

Form10.Show

End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\image\t2-a-4.bmp")
Image2.Visible = True
End Sub

Private Sub Label14_Click()
Form11.Show
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\image\t2-a-2.bmp")
Image2.Visible = True
End Sub

Private Sub Label2_Click()
Form6.Show
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\image\t2-a-1.bmp")
Image2.Visible = True
End Sub

Private Sub Label3_Click()
Form3.Show
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\image\t2-a-3.bmp")
Image2.Visible = True
End Sub

Private Sub Label4_Click()
Form4.Show
End Sub

Private Sub Label5_Click()
Form5.Show
End Sub

Private Sub Label6_Click()
Form1.Show
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\image\t3-a-1.bmp")
Image3.Visible = True
End Sub

Private Sub Label7_Click()
Form7.Show
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\image\t3-a-2.bmp")
Image3.Visible = True
End Sub

Private Sub Label8_Click()
Form8.Show
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\image\t3-a-3.bmp")
Image3.Visible = True
End Sub

Private Sub Label9_Click()
Form9.Show
End Sub

Private Sub mu111_Click()

End Sub

Private Sub mu111help_Click()
       
       

End Sub

Private Sub muekfflrl_Click()

End Sub

Private Sub muend_Click()


End Sub

Private Sub muhelp1_Click()
 
End Sub

Private Sub mulong_Click()

End Sub

Private Sub mume_Click()
           If mumain.Option12 = True Then
               mumain.Agent1.Characters("Genie").Speak "안녕하세요..^^" + Chr(13) + Chr(13) _
                  + "이 프로그램의 제작자인 최인균 이라고 합니다." + Chr(13) + Chr(13) _
                  + "현재 한국애니메이션 고등학교 컴퓨터게임제작과 1학년에 재학중입니다." + Chr(13) + Chr(13) _
                  + "86년 2월22일 태어났고요, 소년입니다." + Chr(13) + Chr(13) _
                  + "E-Mail = dingpong@hitel.net" + Chr(13) + Chr(13) _
                  + "그럼 재미있게 즐기세요..^^" + Chr(13) + Chr(13) _
                  + "1.4V 제작 : 2001년 2월"
End If
MsgBox "안녕하세요..^^" + Chr(13) + Chr(13) _
                  + "이 프로그램의 제작자인 최인균 이라고 합니다." + Chr(13) + Chr(13) _
                  + "비록 잘 못짠 프로그램 일지라도 잘 봐주세요." + Chr(13) + Chr(13) _
                  + "제 소개를 할께요." + Chr(13) + Chr(13) _
                  + "86년 2월22일 태어났고요, 소년입니다." + Chr(13) + Chr(13) _
                  + "E-Mail = dingpong@hitel.net" + Chr(13) + Chr(13) _
                  + "그럼 재미있게 즐기세요..^^" + Chr(13) + Chr(13) _
                  + "1.4V 제작 : 2001년 2월"

End Sub

Private Sub murjawjd_Click()
End Sub

Private Sub murmf_Click()

End Sub

Private Sub musdsd_Click()
Form7.Show
End Sub

Private Sub muso_Click()

End Sub

Private Sub muver_Click()
muloag.Show
End Sub

Private Sub muwkfl_Click()

End Sub

Private Sub muwkqrl_Click()

End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\image\t3-a-4.bmp")
Image3.Visible = True
End Sub

Private Sub check3_Click()
If Check3.Value = 1 Then
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"

mid2.FileName = App.Path + "\sound\Toheart.mid"
mid2.Command = "open"
mid2.Command = "play"
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0
Check9.Value = 0
Check10.Value = 0
End If
End Sub

Private Sub check10_Click()
If Check10.Value = 1 Then
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"


mid10.FileName = App.Path + "\sound\Yestrday.mid"
mid10.Command = "open"
mid10.Command = "play"
Check4.Value = 0
Check3.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0
Check9.Value = 0
Check5.Value = 0
End If

End Sub

Private Sub check31_Click()


End Sub

Private Sub check4_Click()
If Check4.Value = 1 Then
mid2.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"


mid4.FileName = App.Path + "\sound\31.mid"
mid4.Command = "open"
mid4.Command = "play"
Check3.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0
Check9.Value = 0
Check10.Value = 0
End If

End Sub

Private Sub check5_Click()
If Check5.Value = 1 Then
mid2.Command = "close"
mid4.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"

mid5.FileName = App.Path + "\sound\11.mid"
mid5.Command = "open"
mid5.Command = "play"
Check4.Value = 0
Check3.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0
Check9.Value = 0
Check10.Value = 0
End If

End Sub

Private Sub Option5_Click()
Check3.Enabled = True
Check4.Enabled = True
Check5.Enabled = True
Check6.Enabled = True
Check7.Enabled = True
Check8.Enabled = True
Check9.Enabled = True
Check10.Enabled = True
If Check3.Value = 1 Then
mid2.FileName = App.Path + "\sound\Toheart.mid"
mid2.Command = "open"
mid2.Command = "play"
End If
If Check4.Value = 1 Then
mid4.FileName = App.Path + "\sound\31.mid"
mid4.Command = "open"
mid4.Command = "play"
End If
If Check3.Value = 1 Then
mid5.FileName = App.Path + "\sound\11.mid"
mid5.Command = "open"
mid5.Command = "play"
End If
If Check6.Value = 1 Then
mid6.FileName = App.Path + "\sound\2-02 Ahead on Our Way Midi.mid"
mid6.Command = "open"
mid6.Command = "play"
End If
If Check7.Value = 1 Then
mid7.FileName = App.Path + "\sound\12.mid"
mid7.Command = "open"
mid7.Command = "play"
End If
If Check8.Value = 1 Then
mid8.FileName = App.Path + "\sound\34.mid"
mid8.Command = "open"
mid8.Command = "play"
End If
If Check9.Value = 1 Then
mid9.FileName = App.Path + "\sound\Hero.mid"
mid9.Command = "open"
mid9.Command = "play"
End If
If Check10.Value = 1 Then
mid10.FileName = App.Path + "\sound\Yestrday"
mid10.Command = "open"
mid10.Command = "play"
End If


End Sub

Private Sub check7_Click()
If Check7.Value = 1 Then
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid8.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
mid7.FileName = App.Path + "\sound\12.mid"
mid7.Command = "open"
mid7.Command = "play"
Check4.Value = 0
Check3.Value = 0
Check6.Value = 0
Check5.Value = 0
Check8.Value = 0
Check9.Value = 0
Check10.Value = 0
End If

End Sub

Private Sub check8_Click()
If Check8.Value = 1 Then
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid9.Command = "close"
mid10.Command = "close"
mid8.FileName = App.Path + "\sound\34.mid"
mid8.Command = "open"
mid8.Command = "play"
Check4.Value = 0
Check3.Value = 0
Check6.Value = 0
Check7.Value = 0
Check5.Value = 0
Check9.Value = 0
Check10.Value = 0
End If

End Sub

Private Sub check9_Click()

If Check9.Value = 1 Then
mid2.Command = "close"
mid4.Command = "close"
mid5.Command = "close"
mid6.Command = "close"
mid7.Command = "close"
mid8.Command = "close"
mid10.Command = "close"
mid9.FileName = App.Path + "\sound\Hero.mid"
mid9.Command = "open"
mid9.Command = "play"
Check4.Value = 0
Check3.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0
Check5.Value = 0
Check10.Value = 0
End If

End Sub


Private Sub Option6_Click()

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Image3.Visible = False
Image1.Visible = False
End Sub

Private Sub Timer1_Timer()
End
End Sub

Private Sub Timer11_Timer()
'If Option6.Value = 1 Then
'mid2.Command = "close"
'mid4.Command = "close"
'mid5.Command = "close"
'mid6.Command = "close"
'mid7.Command = "close"
'mid8.Command = "close"
'mid9.Command = "close"
'mid10.Command = "close"


'End If

End Sub

Private Sub Timer2_Timer()
If Check3.Value = 1 Then
cnt = cnt + 1
If cnt = 175 Then
mid2.Command = "close"
mid2.Command = "open"
mid2.Command = "play"
cnt = 0
End If
End If
End Sub

Private Sub Timer3_Timer()
If Check4.Value = 1 Then
CNT1 = CNT1 + 1
If CNT1 = 120 Then
mid4.Command = "close"
mid4.Command = "open"
mid4.Command = "play"
CNT1 = 0
End If
End If

End Sub

Private Sub Timer4_Timer()
If Check3.Value = 1 Then
cnt2 = cnt2 + 1
If cnt2 = 155 Then
mid5.Command = "close"
mid5.Command = "open"
mid5.Command = "play"
cnt2 = 0
End If
End If
End Sub

Private Sub Timer5_Timer()
If Check6.Value = 1 Then
cnt3 = cnt3 + 1
If cnt3 = 100 Then
mid6.Command = "close"
mid6.Command = "open"
mid6.Command = "play"
cnt3 = 0
End If

End If
End Sub

Private Sub Timer6_Timer()
If Check7.Value = 1 Then
cnt4 = cnt4 + 1
If cnt4 = 60 Then
mid7.Command = "close"
mid7.Command = "open"
mid7.Command = "play"
cnt4 = 0
End If
End If
End Sub

Private Sub Timer7_Timer()
If Check8.Value = 1 Then
cnt5 = cnt5 + 1
If cnt5 = 120 Then
mid8.Command = "close"
mid8.Command = "open"
mid8.Command = "play"
cnt5 = 0
End If
End If
End Sub

Private Sub Timer8_Timer()
If Check9.Value = 1 Then
cnt6 = cnt6 + 1
If cnt6 = 250 Then
mid9.Command = "close"
mid9.Command = "open"
mid9.Command = "play"
cnt6 = 0
End If
End If
End Sub

Private Sub Timer9_Timer()
If Check10.Value = 1 Then
cnt7 = cnt7 + 1
If cnt7 = 120 Then
mid10.Command = "close"
mid10.Command = "open"
mid10.Command = "play"
cnt7 = 0
End If
End If
End Sub
Private Function StyleEnFonction(ByVal X As Long, ByVal Y As Long) As Long
    StyleEnFonction = 0
    If X < Screen.Width / 30 Then StyleEnFonction = StyleEnFonction + 1
    If Y < Screen.Height / 30 Then StyleEnFonction = 5 - StyleEnFonction
End Function
Private Sub Agent1_Show(ByVal CharacterID As String, ByVal Cause As Integer)
Genie.Top = lasttop
Genie.Left = lastleft

End Sub

