VERSION 5.00
Object = "{4034C11D-8602-11D1-9840-002078110E7D}#1.0#0"; "ASASSISTANTPOPUP.OCX"
Object = "{D19CC187-8393-11D1-983F-002078110E7D}#1.0#0"; "ASBUBBLEFORM.OCX"
Begin VB.Form Form4 
   BorderStyle     =   0  '����
   Caption         =   "Form4"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2655
   LinkTopic       =   "Form4"
   ScaleHeight     =   2310
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �⺻��
   Begin asBubbleWindow.asBubbleForm asBubbleForm1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4048
      BackColor       =   16761024
      Begin asAssistantPopup.asAssisPopup asAssisPopup4 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16761024
         Caption         =   "������"
         Picture         =   "game4.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin asAssistantPopup.asAssisPopup asAssisPopup3 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16761024
         Caption         =   "Ȯ��"
         Picture         =   "game4.frx":059A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BackColor       =   16761024
         Caption         =   "��������"
         Picture         =   "game4.frx":0B34
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BackColor       =   16761024
         Caption         =   "Ÿ��"
         Picture         =   "game4.frx":10CE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�ɸ��Ϳ� ���콺 ������ ��ư�� ������ �޴������� �����ϽǼ� �ֽ��ϴ�."
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asAssisPopup1_Click()

Form1.Agent1.Characters("Genie").Stop
Form1.Agent1.Characters("Genie").Play "Alert"
Form1.Agent1.Characters("Genie").Speak "Ÿ�����α׷��� �����մϴ�."
Form1.Timer2.Enabled = True
Unload Me
End Sub

Private Sub asAssisPopup2_Click()
Form1.Agent1.Characters("Genie").Stop
Form1.Agent1.Characters("Genie").Play "Alert"
Form1.Agent1.Characters("Genie").Speak "�������� ���α׷��� �����մϴ�."
Form1.Timer5.Enabled = True
Unload Me
End Sub

Private Sub asAssisPopup3_Click()
Unload Me
End Sub

Private Sub asAssisPopup4_Click()
    ss = MsgBox("�����Ͻðڽ��ϱ�?", vbCritical + vbYesNo, "����")
    
    If ss = vbYes Then
    End
    End If
End Sub
