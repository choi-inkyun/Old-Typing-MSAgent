VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "´Ü¾î Àâ±â"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   Icon            =   "Å¸ÀÚ ¿¬½À V1.0 8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.TextBox Text1 
      Appearance      =   0  'Æò¸é
      Height          =   460
      Left            =   3870
      TabIndex        =   4
      Top             =   4170
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9120
      TabIndex        =   0
      Text            =   "100"
      Top             =   4280
      Width           =   615
   End
   Begin VB.Timer Timer33 
      Interval        =   1000
      Left            =   10080
      Top             =   1440
   End
   Begin VB.Timer Timer32 
      Interval        =   100
      Left            =   10080
      Top             =   1080
   End
   Begin VB.Timer Timer31 
      Interval        =   100
      Left            =   10080
      Top             =   720
   End
   Begin VB.Timer Timer30 
      Interval        =   100
      Left            =   10080
      Top             =   360
   End
   Begin VB.Timer Timer29 
      Interval        =   100
      Left            =   10080
      Top             =   0
   End
   Begin VB.Timer Timer28 
      Left            =   9720
      Top             =   0
   End
   Begin VB.Timer Timer27 
      Interval        =   100
      Left            =   9360
      Top             =   0
   End
   Begin VB.Timer Timer26 
      Interval        =   100
      Left            =   9360
      Top             =   2640
   End
   Begin VB.Timer Timer25 
      Interval        =   100
      Left            =   9360
      Top             =   2280
   End
   Begin VB.Timer Timer24 
      Interval        =   100
      Left            =   9360
      Top             =   1800
   End
   Begin VB.Timer Timer23 
      Interval        =   100
      Left            =   9720
      Top             =   1440
   End
   Begin VB.Timer Timer22 
      Interval        =   100
      Left            =   9720
      Top             =   1080
   End
   Begin VB.Timer Timer21 
      Interval        =   100
      Left            =   9720
      Top             =   720
   End
   Begin VB.Timer Timer20 
      Interval        =   100
      Left            =   9720
      Top             =   360
   End
   Begin VB.Timer Timer19 
      Interval        =   100
      Left            =   9720
      Top             =   4560
   End
   Begin VB.Timer Timer18 
      Interval        =   100
      Left            =   9720
      Top             =   4200
   End
   Begin VB.Timer Timer17 
      Interval        =   100
      Left            =   9720
      Top             =   3840
   End
   Begin VB.Timer Timer16 
      Interval        =   100
      Left            =   9720
      Top             =   3480
   End
   Begin VB.Timer Timer15 
      Interval        =   100
      Left            =   9720
      Top             =   3120
   End
   Begin VB.Timer Timer14 
      Interval        =   100
      Left            =   9720
      Top             =   2760
   End
   Begin VB.Timer Timer13 
      Interval        =   100
      Left            =   9720
      Top             =   2400
   End
   Begin VB.Timer Timer12 
      Interval        =   100
      Left            =   10080
      Top             =   2040
   End
   Begin VB.Timer Timer11 
      Interval        =   100
      Left            =   9720
      Top             =   2040
   End
   Begin VB.Timer Timer10 
      Interval        =   100
      Left            =   10080
      Top             =   2400
   End
   Begin VB.Timer Timer9 
      Interval        =   100
      Left            =   10080
      Top             =   2760
   End
   Begin VB.Timer Timer8 
      Interval        =   100
      Left            =   10080
      Top             =   3120
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   10080
      Top             =   3480
   End
   Begin VB.Timer Timer6 
      Interval        =   100
      Left            =   10080
      Top             =   3840
   End
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   10080
      Top             =   4200
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   10080
      Top             =   4560
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   10080
      Top             =   4920
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   10080
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10080
      Top             =   5640
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "³²Àº ½Ã°£ :"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   25
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Åõ¸í
      Height          =   615
      Left            =   4440
      MousePointer    =   2  '½ÊÀÚÇü
      TabIndex        =   24
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Åõ¸í
      Height          =   735
      Left            =   9000
      MousePointer    =   2  '½ÊÀÚÇü
      TabIndex        =   22
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Åõ¸í
      Height          =   615
      Left            =   5640
      MousePointer    =   2  '½ÊÀÚÇü
      TabIndex        =   21
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   8040
      TabIndex        =   20
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   5880
      TabIndex        =   19
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   3720
      TabIndex        =   18
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   1440
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   8040
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   5880
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   3720
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1440
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   8040
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5880
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3720
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1440
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8040
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¸ÂÀº °¹¼ö :"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Åõ¸í
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   1
      Top             =   4270
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "Å¸ÀÚ ¿¬½À V1.0 8.frx":030A
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ³¹¸»(200), b, bb, c, d, e, f, g, h, i, k, l, m, n, o, p, q, r, cc, dd, ee, ff, gg, hh, ii, kk, ll, mm, nn, oo, pp, qq, rr, aaa, ab


Private Sub Form_Load()
Randomize
Open App.Path & "\txt\Å¸ÀÚ±Û.txt" For Input As #1
For i = 1 To 200
    Input #1, ³¹¸»(i)
Next
Close #1

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
Timer17.Enabled = False
Timer18.Enabled = False
Timer19.Enabled = False
Timer20.Enabled = False
Timer21.Enabled = False
Timer22.Enabled = False
Timer23.Enabled = False
Timer24.Enabled = False
Timer25.Enabled = False
Timer26.Enabled = False
Timer27.Enabled = False
Timer28.Enabled = False
Timer29.Enabled = False
Timer30.Enabled = False
Timer31.Enabled = False
Timer32.Enabled = False
Timer33.Enabled = False
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0
Timer5.Interval = 0
Timer6.Interval = 0
Timer7.Interval = 0
Timer8.Interval = 0
Timer9.Interval = 0
Timer10.Interval = 0
Timer12.Interval = 0
Timer13.Interval = 0
Timer14.Interval = 0
Timer15.Interval = 0
Timer16.Interval = 0
Timer17.Interval = 0
Timer18.Interval = 0
Timer19.Interval = 0
Timer20.Interval = 0
Timer21.Interval = 0
Timer22.Interval = 0
Timer23.Interval = 0
Timer24.Interval = 0
Timer25.Interval = 0
Timer26.Interval = 0
Timer27.Interval = 0
Timer28.Interval = 0
Timer29.Interval = 0
Timer30.Interval = 0
Timer31.Interval = 0
Timer32.Interval = 0
Label3.Caption = 0
Label5.Enabled = False
Label5.Visible = False
Text2.Enabled = True
Text2.Visible = True

End Sub

Private Sub Label6_Click()
Randomize
mumain.Agent1.Characters("Genie").Stop
aaa = Val(Text2.Text)
Label3.Caption = 0
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
Timer17.Enabled = False
Timer18.Enabled = False
Timer19.Enabled = False
Timer20.Enabled = False
Timer21.Enabled = False
Timer22.Enabled = False
Timer23.Enabled = False
Timer24.Enabled = False
Timer25.Enabled = False
Timer26.Enabled = False
Timer27.Enabled = False
Timer28.Enabled = False
Timer29.Enabled = False
Timer30.Enabled = False
Timer31.Enabled = False
Timer32.Enabled = False
Timer33.Enabled = False
For i = 0 To 15
Label1(i).Visible = False
a = Int(Rnd(1) * 200)
Label1(i).Caption = ³¹¸»(a)
Next
Label1(0).Top = 840
Label1(1).Top = 840
Label1(2).Top = 840
Label1(3).Top = 840
Label1(4).Top = 2160
Label1(5).Top = 2160
Label1(6).Top = 2160
Label1(7).Top = 2160
Label1(8).Top = 3480
Label1(9).Top = 3480
Label1(10).Top = 3480
Label1(11).Top = 3480
Label1(12).Top = 4800
Label1(13).Top = 4800
Label1(14).Top = 4800
Label1(15).Top = 4800
Label5.Visible = False
Label5.Enabled = False
Text2.Visible = True
Text2.Enabled = True

End Sub

Private Sub Label7_Click()
mumain.Agent1.Characters("Genie").Stop
mumain.Show
Unload Me
End Sub

Private Sub Label8_Click()
Randomize
Text1.IMEMode = vbIMEModeHangul 'ÇÑ±Û
aaa = Val(Text2.Text)
Timer1.Enabled = True
Timer1.Enabled = True
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
Timer33.Enabled = True
For i = 0 To 15
a = Int(Rnd(1) * 200)
Label1(i).Caption = ³¹¸»(a)
Label1(i).Visible = False
Next

Timer1.Interval = Int(Rnd(500) * 2000)
Timer2.Interval = Int(Rnd(500) * 2000)
Timer3.Interval = Int(Rnd(500) * 2000)
Timer4.Interval = Int(Rnd(500) * 2000)
Timer5.Interval = Int(Rnd(500) * 2000)
Timer6.Interval = Int(Rnd(500) * 2000)
Timer7.Interval = Int(Rnd(500) * 2000)
Timer8.Interval = Int(Rnd(500) * 2000)
Timer9.Interval = Int(Rnd(500) * 2000)
Timer10.Interval = Int(Rnd(500) * 2500)
Timer11.Interval = Int(Rnd(500) * 2500)
Timer12.Interval = Int(Rnd(500) * 2500)
Timer13.Interval = Int(Rnd(500) * 2500)
Timer14.Interval = Int(Rnd(500) * 2500)
Timer15.Interval = Int(Rnd(500) * 2500)
Timer16.Interval = Int(Rnd(500) * 2500)
Timer17.Interval = Int(Rnd(500) * 3000)
Timer18.Interval = Int(Rnd(500) * 3000)
Timer19.Interval = Int(Rnd(500) * 3000)
Timer20.Interval = Int(Rnd(500) * 3000)
Timer21.Interval = Int(Rnd(500) * 3000)
Timer22.Interval = Int(Rnd(500) * 3000)
Timer23.Interval = Int(Rnd(500) * 3500)
Timer24.Interval = Int(Rnd(500) * 3500)
Timer25.Interval = Int(Rnd(500) * 3500)
Timer26.Interval = Int(Rnd(500) * 3500)
Timer27.Interval = Int(Rnd(500) * 3500)
Timer28.Interval = Int(Rnd(500) * 3500)
Timer29.Interval = Int(Rnd(500) * 4000)
Timer30.Interval = Int(Rnd(500) * 4000)
Timer31.Interval = Int(Rnd(500) * 4000)
Timer32.Interval = Int(Rnd(500) * 4000)
Label5.Visible = True
Label5.Enabled = True
Text2.Enabled = False
Text2.Visible = False
Text1.SetFocus
Text1.Text = ""

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 32 Then
    For i = 0 To 15
        If Trim(Text1.Text) = Label1(i).Caption Then
            If Text1.Text <> "" Then
            Label3.Caption = Label3.Caption + 1
            Label1(i).Caption = ""
            End If
         End If
    Next
    Text1.Text = ""
ab = Label3.Caption

End If
End Sub

Private Sub Timer1_Timer()
b = Int(Rnd(10) * 100)
Label1(0).Top = Label1(0).Top - b
If Label1(0).Top <= 50 Then
Label1(0).Visible = True
End If
If Label1(0).Top <= 0 Then
b = 0
Label1(0).Top = 0
Timer1.Enabled = False
Timer17.Enabled = True
End If
End Sub

Private Sub Timer10_Timer()
l = Int(Rnd(10) * 100)
Label1(9).Top = Label1(9).Top - l
If Label1(9).Top <= 1560 Then
Label1(9).Visible = True
End If
If Label1(9).Top <= 1440 Then
l = 0
Label1(9).Top = 1440
Timer10.Enabled = False
Timer26.Enabled = True
End If

End Sub

Private Sub Timer11_Timer()
m = Int(Rnd(10) * 100)
Label1(10).Top = Label1(10).Top - m
If Label1(10).Top <= 1560 Then
Label1(10).Visible = True
End If
If Label1(10).Top <= 1440 Then
m = 0
Label1(10).Top = 1440
Timer11.Enabled = False
Timer27.Enabled = True
End If

End Sub

Private Sub Timer12_Timer()
n = Int(Rnd(10) * 100)
Label1(11).Top = Label1(11).Top - n
If Label1(11).Top <= 1560 Then
Label1(11).Visible = True
End If
If Label1(11).Top <= 1440 Then
n = 0
Label1(11).Top = 1440
Timer12.Enabled = False
Timer28.Enabled = True
End If

End Sub

Private Sub Timer13_Timer()
o = Int(Rnd(10) * 100)
Label1(12).Top = Label1(12).Top - o
If Label1(12).Top <= 2400 Then
Label1(12).Visible = True
End If
If Label1(12).Top <= 2280 Then
o = 0
Label1(12).Top = 2280
Timer13.Enabled = False
Timer29.Enabled = True
End If

End Sub

Private Sub Timer14_Timer()
p = Int(Rnd(10) * 100)
Label1(13).Top = Label1(13).Top - p
If Label1(13).Top <= 2400 Then
Label1(13).Visible = True
End If
If Label1(13).Top <= 2280 Then
p = 0
Label1(13).Top = 2280
Timer14.Enabled = False
Timer30.Enabled = True
End If

End Sub

Private Sub Timer15_Timer()
q = Int(Rnd(10) * 100)
Label1(14).Top = Label1(14).Top - q
If Label1(14).Top <= 2400 Then
Label1(14).Visible = True
End If
If Label1(14).Top <= 2280 Then
q = 0
Label1(14).Top = 2280
Timer15.Enabled = False
Timer31.Enabled = True
End If

End Sub

Private Sub Timer16_Timer()
r = Int(Rnd(10) * 100)
Label1(15).Top = Label1(15).Top - r
If Label1(15).Top <= 2400 Then
Label1(15).Visible = True
End If
If Label1(15).Top <= 2280 Then
r = 0
Label1(15).Top = 2280
Timer16.Enabled = False
Timer32.Enabled = True
End If

End Sub

Private Sub Timer17_Timer()
bb = Int(Rnd(10) * 100)
Label1(0).Top = Label1(0).Top + bb
If Label1(0).Top >= 50 Then
Label1(0).Visible = False
End If
If Label1(0).Top >= 600 Then
Timer17.Enabled = False
Timer1.Enabled = True
Label1(0).Caption = ""
a = Int(Rnd(1) * 200)
Label1(0).Caption = ³¹¸»(a)
End If
End Sub

Private Sub Timer18_Timer()
cc = Int(Rnd(10) * 100)
Label1(1).Top = Label1(1).Top + cc
If Label1(1).Top >= 50 Then
Label1(1).Visible = False
End If
If Label1(1).Top >= 600 Then
Timer18.Enabled = False
Timer2.Enabled = True
Label1(1).Caption = ""
a = Int(Rnd(1) * 200)
Label1(1).Caption = ³¹¸»(a)
End If

End Sub

Private Sub Timer19_Timer()
dd = Int(Rnd(10) * 100)
Label1(2).Top = Label1(2).Top + dd
If Label1(2).Top >= 50 Then
Label1(2).Visible = False
End If
If Label1(2).Top >= 600 Then
Timer19.Enabled = False
Timer3.Enabled = True
Label1(2).Caption = ""
a = Int(Rnd(2) * 200)
Label1(2).Caption = ³¹¸»(a)
End If

End Sub

Private Sub Timer2_Timer()
c = Int(Rnd(10) * 100)
Label1(1).Top = Label1(1).Top - c
If Label1(1).Top <= 50 Then
Label1(1).Visible = True
End If
If Label1(1).Top <= 0 Then
c = 0
Label1(1).Top = 0
Timer2.Enabled = False
Timer18.Enabled = True
End If
End Sub

Private Sub Timer20_Timer()
ee = Int(Rnd(10) * 100)
Label1(3).Top = Label1(3).Top + ee
If Label1(3).Top >= 50 Then
Label1(3).Visible = False
End If
If Label1(3).Top >= 600 Then
Timer20.Enabled = False
Timer4.Enabled = True
Label1(3).Caption = ""
a = Int(Rnd(1) * 200)
Label1(3).Caption = ³¹¸»(a)
End If

End Sub

Private Sub Timer21_Timer()
ff = Int(Rnd(10) * 100)
Label1(4).Top = Label1(4).Top + ff
If Label1(4).Top >= 720 Then
Label1(4).Visible = False
End If
If Label1(4).Top >= 1440 Then
Timer21.Enabled = False
Timer5.Enabled = True
Label1(4).Caption = ""
a = Int(Rnd(1) * 200)
Label1(4).Caption = ³¹¸»(a)
End If

End Sub

Private Sub Timer22_Timer()
gg = Int(Rnd(10) * 100)
Label1(5).Top = Label1(5).Top + gg
If Label1(5).Top >= 720 Then
Label1(5).Visible = False
End If
If Label1(5).Top >= 1440 Then
Timer22.Enabled = False
Timer6.Enabled = True
Label1(5).Caption = ""
a = Int(Rnd(1) * 200)
Label1(5).Caption = ³¹¸»(a)
End If

End Sub
    
Private Sub Timer23_Timer()
hh = Int(Rnd(10) * 100)
Label1(6).Top = Label1(6).Top + hh
If Label1(6).Top >= 720 Then
Label1(6).Visible = False
End If
If Label1(6).Top >= 1440 Then
Timer23.Enabled = False
Timer7.Enabled = True
Label1(6).Caption = ""
a = Int(Rnd(1) * 200)
Label1(6).Caption = ³¹¸»(a)
End If
End Sub

Private Sub Timer24_Timer()
ii = Int(Rnd(10) * 100)
Label1(7).Top = Label1(7).Top + ii
If Label1(7).Top >= 720 Then
Label1(7).Visible = False
End If
If Label1(7).Top >= 1440 Then
Timer24.Enabled = False
Timer8.Enabled = True
Label1(7).Caption = ""
a = Int(Rnd(1) * 200)
Label1(7).Caption = ³¹¸»(a)
End If

End Sub

Private Sub Timer25_Timer()
kk = Int(Rnd(10) * 100)
Label1(8).Top = Label1(8).Top + kk
If Label1(8).Top >= 1560 Then
Label1(8).Visible = False
End If
If Label1(8).Top >= 2280 Then
Timer25.Enabled = False
Timer9.Enabled = True
Label1(8).Caption = ""
a = Int(Rnd(1) * 200)
Label1(8).Caption = ³¹¸»(a)
End If

End Sub

Private Sub Timer26_Timer()
ll = Int(Rnd(10) * 100)
Label1(9).Top = Label1(9).Top + ll
If Label1(9).Top >= 1560 Then
Label1(9).Visible = False
End If
If Label1(9).Top >= 2280 Then
Timer26.Enabled = False
Timer10.Enabled = True
Label1(9).Caption = ""
a = Int(Rnd(1) * 200)
Label1(9).Caption = ³¹¸»(a)
End If

End Sub

Private Sub Timer27_Timer()
mm = Int(Rnd(10) * 100)
Label1(10).Top = Label1(10).Top + mm
If Label1(10).Top >= 1560 Then
Label1(10).Visible = False
End If
If Label1(10).Top >= 2280 Then
Timer27.Enabled = False
Timer11.Enabled = True
Label1(10).Caption = ""
a = Int(Rnd(1) * 200)
Label1(10).Caption = ³¹¸»(a)
End If
End Sub

Private Sub Timer28_Timer()
nn = Int(Rnd(10) * 100)
Label1(11).Top = Label1(11).Top + nn
If Label1(11).Top >= 1560 Then
Label1(11).Visible = False
End If
If Label1(11).Top >= 2280 Then
Timer28.Enabled = False
Timer12.Enabled = True
Label1(11).Caption = ""
a = Int(Rnd(1) * 200)
Label1(11).Caption = ³¹¸»(a)
End If

End Sub

Private Sub Timer29_Timer()
oo = Int(Rnd(10) * 100)
Label1(12).Top = Label1(12).Top + oo
If Label1(12).Top >= 2400 Then
Label1(12).Visible = False
End If
If Label1(12).Top >= 3120 Then
Timer29.Enabled = False
Timer13.Enabled = True
Label1(12).Caption = ""
a = Int(Rnd(1) * 200)
Label1(12).Caption = ³¹¸»(a)
End If

End Sub

Private Sub Timer3_Timer()
d = Int(Rnd(10) * 100)
Label1(2).Top = Label1(2).Top - d
If Label1(2).Top <= 50 Then
Label1(2).Visible = True
End If
If Label1(2).Top <= 0 Then
d = 0
Label1(2).Top = 0
Timer3.Enabled = False
Timer19.Enabled = True
End If

End Sub

Private Sub Timer30_Timer()
pp = Int(Rnd(10) * 100)
Label1(13).Top = Label1(13).Top + pp
If Label1(13).Top >= 2400 Then
Label1(13).Visible = False
End If
If Label1(13).Top >= 3120 Then
Timer30.Enabled = False
Timer14.Enabled = True
Label1(13).Caption = ""
a = Int(Rnd(1) * 200)
Label1(13).Caption = ³¹¸»(a)
End If


End Sub

Private Sub Timer31_Timer()
qq = Int(Rnd(10) * 100)
Label1(14).Top = Label1(14).Top + qq
If Label1(14).Top >= 2400 Then
Label1(14).Visible = False
End If
If Label1(14).Top >= 3120 Then
Timer31.Enabled = False
Timer15.Enabled = True
Label1(14).Caption = ""
a = Int(Rnd(1) * 200)
Label1(14).Caption = ³¹¸»(a)
End If


End Sub

Private Sub Timer32_Timer()
rr = Int(Rnd(10) * 100)
Label1(15).Top = Label1(15).Top + rr
If Label1(15).Top >= 2400 Then
Label1(15).Visible = False
End If
If Label1(15).Top >= 3120 Then
Timer32.Enabled = False
Timer16.Enabled = True
Label1(15).Caption = ""
a = Int(Rnd(1) * 200)
Label1(15).Caption = ³¹¸»(a)
End If


End Sub

Private Sub Timer33_Timer()
aaa = aaa - 1
Label5.Caption = aaa
If Label5 = 0 Then
MsgBox "½Ã°£ÀÌ ´Ù µÇ¾ú½À´Ï´Ù."
MsgBox "´ç½ÅÀÇ ¸ÂÀº °¹¼ö´Â " & ab & "°³ ÀÔ´Ï´Ù."
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
Timer17.Enabled = False
Timer18.Enabled = False
Timer19.Enabled = False
Timer20.Enabled = False
Timer21.Enabled = False
Timer22.Enabled = False
Timer23.Enabled = False
Timer24.Enabled = False
Timer25.Enabled = False
Timer26.Enabled = False
Timer27.Enabled = False
Timer28.Enabled = False
Timer29.Enabled = False
Timer30.Enabled = False
Timer31.Enabled = False
Timer32.Enabled = False
Timer33.Enabled = False
For i = 0 To 15
Label1(i).Visible = False
a = Int(Rnd(1) * 200)
Label1(i).Caption = ³¹¸»(a)

Next
Label1(0).Top = 600
Label1(1).Top = 600
Label1(2).Top = 600
Label1(3).Top = 600
Label1(4).Top = 1440
Label1(5).Top = 1440
Label1(6).Top = 1440
Label1(7).Top = 1440
Label1(8).Top = 2280
Label1(9).Top = 2280
Label1(10).Top = 2280
Label1(11).Top = 2280
Label1(12).Top = 3120
Label1(13).Top = 3120
Label1(14).Top = 3120
Label1(15).Top = 3120

Label5.Visible = False
Label5.Enabled = False
Text2.Enabled = True
Text2.Visible = True
Text1.Text = ""
Label3.Caption = 0
End If
End Sub

Private Sub Timer4_Timer()
e = Int(Rnd(10) * 100)
Label1(3).Top = Label1(3).Top - e
If Label1(3).Top <= 50 Then
Label1(3).Visible = True
End If
If Label1(3).Top <= 0 Then
e = 0
Label1(3).Top = 50
Timer4.Enabled = False
Timer20.Enabled = True
End If

End Sub

Private Sub Timer5_Timer()
f = Int(Rnd(10) * 100)
Label1(4).Top = Label1(4).Top - f
If Label1(4).Top <= 720 Then
Label1(4).Visible = True
End If
If Label1(4).Top <= 600 Then
f = 0
Label1(4).Top = 600
Timer5.Enabled = False
Timer21.Enabled = True
End If

End Sub

Private Sub Timer6_Timer()
g = Int(Rnd(10) * 100)
Label1(5).Top = Label1(5).Top - g
If Label1(5).Top <= 720 Then
Label1(5).Visible = True
End If
If Label1(5).Top <= 600 Then
g = 0
Label1(5).Top = 600
Timer6.Enabled = False
Timer22.Enabled = True
End If

End Sub

Private Sub Timer7_Timer()
h = Int(Rnd(10) * 100)
Label1(6).Top = Label1(6).Top - h
If Label1(6).Top <= 720 Then
Label1(6).Visible = True
End If
If Label1(6).Top <= 600 Then
h = 0
Label1(6).Top = 600
Timer7.Enabled = False
Timer23.Enabled = True
End If

End Sub

Private Sub Timer8_Timer()
i = Int(Rnd(10) * 100)
Label1(7).Top = Label1(7).Top - i
If Label1(7).Top <= 720 Then
Label1(7).Visible = True
End If
If Label1(7).Top <= 600 Then
i = 0
Label1(7).Top = 600
Timer8.Enabled = False
Timer24.Enabled = True
End If
End Sub

Private Sub Timer9_Timer()
k = Int(Rnd(10) * 100)
Label1(8).Top = Label1(8).Top - k
If Label1(8).Top <= 1560 Then
Label1(8).Visible = True
End If
If Label1(8).Top <= 1440 Then
k = 0
Label1(8).Top = 1440
Timer9.Enabled = False
Timer25.Enabled = True
End If

End Sub
