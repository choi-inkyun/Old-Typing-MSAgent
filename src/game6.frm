VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form6 
   BorderStyle     =   0  '없음
   Caption         =   "Form6"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   LinkTopic       =   "Form6"
   ScaleHeight     =   2985
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   4800
      Top             =   0
   End
   Begin MCI.MMControl mid1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   0
      Picture         =   "game6.frx":0000
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
mid1.FileName = App.Path + "\sound\Start.wav"
mid1.Command = "open"
mid1.Command = "play"
End Sub

Private Sub Label3_Click()
End Sub

Private Sub Timer1_Timer()
Unload Me
Form1.Visible = True
End Sub


