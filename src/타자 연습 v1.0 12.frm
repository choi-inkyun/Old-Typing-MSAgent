VERSION 5.00
Object = "{4034C11D-8602-11D1-9840-002078110E7D}#1.0#0"; "ASASSISTANTPOPUP.OCX"
Object = "{D19CC187-8393-11D1-983F-002078110E7D}#1.0#0"; "ASBUBBLEFORM.OCX"
Begin VB.Form Form12 
   BorderStyle     =   0  '없음
   Caption         =   "Form12"
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   LinkTopic       =   "Form12"
   ScaleHeight     =   1995
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin asBubbleWindow.asBubbleForm asBubbleForm1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3625
      BackColor       =   12648447
      Begin asAssistantPopup.asAssisPopup asAssisPopup3 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   12648447
         Caption         =   "종료"
         Picture         =   "타자 연습 v1.0 12.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin asAssistantPopup.asAssisPopup asAssisPopup2 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackColor       =   12648447
         Caption         =   "확인"
         Picture         =   "타자 연습 v1.0 12.frx":059A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin asAssistantPopup.asAssisPopup asAssisPopup1 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BackColor       =   12648447
         Caption         =   "메인"
         Picture         =   "타자 연습 v1.0 12.frx":0B34
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "케릭터에 마우스 오른쪽 버튼을 누르면 메뉴에서도 선택하실수 있습니다."
         ForeColor       =   &H00FF00FF&
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asAssisPopup1_Click()

mumain.Agent1.Characters("Genie").Stop
mumain.Agent1.Characters("Genie").Speak "메인프로그램을 실행합니다."
Form2.Timer1.Enabled = True
End Sub

Private Sub asAssisPopup2_Click()
Unload Me
End Sub

Private Sub asAssisPopup3_Click()
mumain.Agent1.Characters("Genie").Stop
    
    ss = MsgBox("종료하시겠습니까?", vbCritical + vbYesNo, "종료")
    
    If ss = vbYes Then
    Form2.Timer1.Enabled = True
    mumain.Timer1.Enabled = True
    mumain.Agent1.Characters("Genie").Play "GetAttentionContinued"
        
    Else
        Exit Sub
    End If
End Sub

