VERSION 5.00
Begin VB.Form muloag 
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5385
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2400
      Top             =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      X1              =   240
      X2              =   5160
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      X1              =   5160
      X2              =   5160
      Y1              =   240
      Y2              =   3120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   240
      X2              =   5160
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   3120
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF00FF&
      BorderColor     =   &H00FF8080&
      FillColor       =   &H0000FF00&
      Height          =   2895
      Left            =   240
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "muloag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Label1.Caption = "타자 연습" + Chr(13) + Chr(13) + "      프로그램 Ver 1.0"
Label2.Caption = "제작 : 최인균"
Label3.Caption = "이 프로그램은 공개 프로그램 입니다."
End Sub

Private Sub Timer1_Timer()
Unload muloag
mumain.Show
End Sub

