VERSION 5.00
Object = "{4034C11D-8602-11D1-9840-002078110E7D}#1.0#0"; "ASASSISTANTPOPUP.OCX"
Object = "{D19CC187-8393-11D1-983F-002078110E7D}#1.0#0"; "ASBUBBLEFORM.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000005&
   BorderStyle     =   0  '없음
   Caption         =   "Form2"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   LinkTopic       =   "Form2"
   ScaleHeight     =   3405
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Visible         =   0   'False
   Begin asBubbleWindow.asBubbleForm asBubbleForm1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5953
      BorderColor     =   0
      BackColor       =   12648384
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   0
         Top             =   0
      End
      Begin asAssistantPopup.asAssisPopup asAssisPopup1 
         Height          =   495
         Left            =   3120
         TabIndex        =   1
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BackColor       =   12648384
         Caption         =   "확 인"
         Picture         =   "game2.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "바탕체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   60
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pp As Byte
Private Sub asAssisPopup1_Click()
Form1.Agent1.Characters("Genie").Stop
Form1.Agent1.Characters("Genie").Play "Alert"
Form1.Agent1.Characters("Genie").LanguageID = &H412
Form2.Hide
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
pp = 0

Timer1.Enabled = True

Label1.Caption = Chr(13) + "이름 : 최인균" _
                + Chr(13) + Chr(13) + "학교 : 한국애니메이션고등학교 2기생" _
                + Chr(13) + Chr(13) + "컴퓨터 게임제작과 1학년 재학중" _
                + Chr(13) + Chr(13) + "제작 : 2001년 2월." _
                + Chr(13) + Chr(13) + "E-Mail = dingpong@hitel.net" _
                + Chr(13) + Chr(13) + "HomePage = http://www.teamnemo.ce.ro"
                
 
 
End Sub

Private Sub Timer1_Timer()
pp = pp + 1
If pp = 1 Then
Form1.Agent1.Characters("Genie").Play "Read"
Form1.Agent1.Characters("Genie").LanguageID = &H412
Form1.Agent1.Characters("Genie").Speak "최인균. 한국애니메이션고등학교 2기생. 컴퓨터 게임제작과 1학년 재학중. 2001년 2월 제작"
End If
If pp = 21 Then
Form1.Agent1.Characters("Genie").LanguageID = &H409
Form1.Agent1.Characters("Genie").Speak " dingpong@hitel.net, http://user.chollian.net/~redcmk"
End If
If pp = 33 Then
Form1.Agent1.Characters("Genie").Play "Alert"
pp = 0
Timer1.Enabled = False
End If
End Sub

