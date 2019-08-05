VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "Å¸ÀÚ°ËÁ¤"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   Icon            =   "Å¸ÀÚ ¿¬½À V1.0 5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
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
      Height          =   255
      Left            =   1680
      TabIndex        =   115
      Text            =   "100"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   3690
      TabIndex        =   110
      Top             =   4620
      Width           =   5565
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3690
      TabIndex        =   109
      Top             =   3690
      Width           =   5565
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3690
      TabIndex        =   108
      Top             =   2750
      Width           =   5565
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3690
      TabIndex        =   107
      Top             =   1800
      Width           =   5565
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6840
      Top             =   0
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   2
      Left            =   3960
      TabIndex        =   123
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   1
      Left            =   3840
      TabIndex        =   122
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Height          =   615
      Left            =   3000
      MousePointer    =   2  '½ÊÀÚÇü
      TabIndex        =   121
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Åõ¸í
      Height          =   735
      Left            =   9000
      MousePointer    =   2  '½ÊÀÚÇü
      TabIndex        =   120
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   119
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "³²Àº ½Ã°£ :"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   118
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÃÊ"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   117
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "½Ã°£ :"
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
      Left            =   960
      TabIndex        =   116
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   114
      Top             =   4330
      Width           =   6615
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   113
      Top             =   3450
      Width           =   6615
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   112
      Top             =   2500
      Width           =   6615
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   111
      Top             =   1545
      Width           =   5775
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
      Height          =   300
      Index           =   26
      Left            =   6840
      TabIndex        =   106
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   25
      Left            =   6720
      TabIndex        =   105
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   24
      Left            =   6600
      TabIndex        =   104
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   23
      Left            =   6480
      TabIndex        =   103
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   22
      Left            =   6360
      TabIndex        =   102
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   21
      Left            =   6240
      TabIndex        =   101
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   20
      Left            =   6120
      TabIndex        =   100
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   19
      Left            =   6000
      TabIndex        =   99
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   18
      Left            =   5880
      TabIndex        =   98
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   17
      Left            =   5760
      TabIndex        =   97
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   16
      Left            =   5640
      TabIndex        =   96
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   15
      Left            =   5520
      TabIndex        =   95
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   14
      Left            =   5400
      TabIndex        =   94
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   13
      Left            =   5280
      TabIndex        =   93
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   12
      Left            =   5160
      TabIndex        =   92
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   11
      Left            =   5040
      TabIndex        =   91
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   10
      Left            =   4920
      TabIndex        =   90
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   9
      Left            =   4800
      TabIndex        =   89
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   8
      Left            =   4680
      TabIndex        =   88
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   7
      Left            =   4560
      TabIndex        =   87
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   6
      Left            =   4440
      TabIndex        =   86
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   5
      Left            =   4320
      TabIndex        =   85
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   4
      Left            =   4200
      TabIndex        =   84
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   3
      Left            =   4080
      TabIndex        =   83
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   2
      Left            =   3960
      TabIndex        =   82
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   1
      Left            =   3840
      TabIndex        =   81
      Top             =   4980
      Width           =   195
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
      Height          =   300
      Index           =   0
      Left            =   3720
      TabIndex        =   80
      Top             =   4980
      Width           =   200
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   26
      Left            =   6840
      TabIndex        =   79
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   25
      Left            =   6720
      TabIndex        =   78
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   24
      Left            =   6600
      TabIndex        =   77
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   23
      Left            =   6480
      TabIndex        =   76
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   22
      Left            =   6360
      TabIndex        =   75
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   21
      Left            =   6240
      TabIndex        =   74
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   20
      Left            =   6120
      TabIndex        =   73
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   19
      Left            =   6000
      TabIndex        =   72
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   18
      Left            =   5880
      TabIndex        =   71
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   17
      Left            =   5760
      TabIndex        =   70
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   16
      Left            =   5640
      TabIndex        =   69
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   15
      Left            =   5520
      TabIndex        =   68
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   14
      Left            =   5400
      TabIndex        =   67
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   13
      Left            =   5280
      TabIndex        =   66
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   12
      Left            =   5160
      TabIndex        =   65
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   11
      Left            =   5040
      TabIndex        =   64
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   10
      Left            =   4920
      TabIndex        =   63
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   9
      Left            =   4800
      TabIndex        =   62
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   8
      Left            =   4680
      TabIndex        =   61
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   7
      Left            =   4560
      TabIndex        =   60
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   6
      Left            =   4440
      TabIndex        =   59
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   5
      Left            =   4320
      TabIndex        =   58
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   4
      Left            =   4200
      TabIndex        =   57
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   3
      Left            =   4080
      TabIndex        =   56
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   2
      Left            =   3960
      TabIndex        =   55
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   1
      Left            =   3840
      TabIndex        =   54
      Top             =   4035
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Index           =   0
      Left            =   3720
      TabIndex        =   53
      Top             =   4040
      Width           =   200
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   26
      Left            =   6840
      TabIndex        =   52
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   25
      Left            =   6720
      TabIndex        =   51
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   24
      Left            =   6600
      TabIndex        =   50
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   23
      Left            =   6480
      TabIndex        =   49
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   22
      Left            =   6360
      TabIndex        =   48
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   21
      Left            =   6240
      TabIndex        =   47
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   20
      Left            =   6120
      TabIndex        =   46
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   19
      Left            =   6000
      TabIndex        =   45
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   18
      Left            =   5880
      TabIndex        =   44
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   17
      Left            =   5760
      TabIndex        =   43
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   16
      Left            =   5640
      TabIndex        =   42
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   15
      Left            =   5520
      TabIndex        =   41
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   14
      Left            =   5400
      TabIndex        =   40
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   13
      Left            =   5280
      TabIndex        =   39
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   12
      Left            =   5160
      TabIndex        =   38
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   11
      Left            =   5040
      TabIndex        =   37
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   10
      Left            =   4920
      TabIndex        =   36
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   9
      Left            =   4800
      TabIndex        =   35
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   8
      Left            =   4680
      TabIndex        =   34
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   7
      Left            =   4560
      TabIndex        =   33
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   6
      Left            =   4440
      TabIndex        =   32
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   5
      Left            =   4320
      TabIndex        =   31
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   4
      Left            =   4200
      TabIndex        =   30
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   3
      Left            =   4080
      TabIndex        =   29
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   2
      Left            =   3960
      TabIndex        =   28
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   1
      Left            =   3840
      TabIndex        =   27
      Top             =   3110
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   0
      Left            =   3720
      TabIndex        =   26
      Top             =   3110
      Width           =   200
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   27
      Left            =   6960
      TabIndex        =   25
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   26
      Left            =   6840
      TabIndex        =   24
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   25
      Left            =   6720
      TabIndex        =   23
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   24
      Left            =   6600
      TabIndex        =   22
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   23
      Left            =   6480
      TabIndex        =   21
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   22
      Left            =   6360
      TabIndex        =   20
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   21
      Left            =   6240
      TabIndex        =   19
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   20
      Left            =   6120
      TabIndex        =   18
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   19
      Left            =   6000
      TabIndex        =   17
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   18
      Left            =   5880
      TabIndex        =   16
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   17
      Left            =   5760
      TabIndex        =   15
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   16
      Left            =   5640
      TabIndex        =   14
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   15
      Left            =   5520
      TabIndex        =   13
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   14
      Left            =   5400
      TabIndex        =   12
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   13
      Left            =   5280
      TabIndex        =   11
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   12
      Left            =   5160
      TabIndex        =   10
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   11
      Left            =   5040
      TabIndex        =   9
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   10
      Left            =   4920
      TabIndex        =   8
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   9
      Left            =   4800
      TabIndex        =   7
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   8
      Left            =   4680
      TabIndex        =   6
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   7
      Left            =   4560
      TabIndex        =   5
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   6
      Left            =   4440
      TabIndex        =   4
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   5
      Left            =   4320
      TabIndex        =   3
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   4
      Left            =   4200
      TabIndex        =   2
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   3
      Left            =   4080
      TabIndex        =   1
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Index           =   0
      Left            =   3720
      TabIndex        =   0
      Top             =   2160
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "Å¸ÀÚ ¿¬½À V1.0 5.frx":030A
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Form_Load()
Timer2.Enabled = False
Timer1.Enabled = False
For i = 0 To 3
lb(i).Visible = False
Text1(i).Visible = False
lb(i).Enabled = False
Text1(i).Enabled = False
Next
For i = 0 To 26
Label2(i).Visible = False
Label2(i).Enabled = False
Label3(i).Visible = False
Label3(i).Enabled = False
Label4(i).Visible = False
Label4(i).Enabled = False
Label5(i).Visible = False
Label5(i).Enabled = False
Next

k = 1
End Sub

Private Sub Label10_Click()
mumain.Agent1.Characters("Genie").Stop
mumain.Show
Unload Me
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If rr = -1 Then
    If KeyAscii = 13 Then
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    For i = 0 To 3
        lb(i).Caption = mid(±ä±Û, k, 27)
        k = k + 27
    Next
    End If
End If


If KeyAscii = 13 Then
    
    rr = rr + 1
    Text1(rr).SetFocus
    If rr = 3 Then
    rr = -1
'       If KeyAscii = 13 Then
 '           MsgBox "11"
  '      End If
     End If
             If rr = 0 Then
        If mumain.Option12 = True Then
        mumain.Agent1.Characters("Genie").Speak lb(0).Caption
        End If
        End If
        If rr = 1 Then
        If mumain.Option12 = True Then
        mumain.Agent1.Characters("Genie").Speak lb(1).Caption
        End If
        End If
        If rr = 2 Then
        If mumain.Option12 = True Then
        mumain.Agent1.Characters("Genie").Speak lb(2).Caption
        End If
        End If
        If rr = -1 Then
        If mumain.Option12 = True Then
        mumain.Agent1.Characters("Genie").Speak lb(3).Caption
        End If
        End If
        
End If

End Sub

Private Sub Label1_Click()
Timer2.Enabled = True
Timer1.Enabled = True
For i = 0 To 3
lb(i).Visible = True
Text1(i).Visible = True
lb(i).Enabled = True
Text1(i).Enabled = True
Next
Text1(0).SetFocus
For i = 0 To 26
Label2(i).Caption = "x"
Label3(i).Caption = "x"
Label4(i).Caption = "x"
Label5(i).Caption = "x"
Label2(i).Visible = True
Label3(i).Visible = True
Label4(i).Visible = True
Label5(i).Visible = True
Next
Label6.Caption = Val(Text2.Text)
gg = Val(Label6.Caption)
Timer2.Enabled = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
If Text2.Text = "" Then
Label6.Caption = 300
End If

If mumain.Option12 = True Then
mumain.Agent1.Characters("Genie").Speak lb(0).Caption
End If

End Sub

Private Sub Timer1_Timer()
For i = 0 To 26
If Left(Text1(0).Text, i + 1) = Left(lb(0).Caption, i + 1) Then
        Label2(i).Caption = "o"
        End If
         Next
For i = 0 To 26
If Left(Text1(0).Text, i + 1) <> Left(lb(0).Caption, i + 1) Then
       Label2(i).Caption = "x"
        End If
        Next
For i = 0 To 26
If Left(Text1(1).Text, i + 1) = Left(lb(1).Caption, i + 1) Then
        Label3(i).Caption = "o"
        End If
         Next
For i = 0 To 26
If Left(Text1(1).Text, i + 1) <> Left(lb(1).Caption, i + 1) Then
       Label3(i).Caption = "x"
        End If
        Next
For i = 0 To 26
If Left(Text1(2).Text, i + 1) = Left(lb(2).Caption, i + 1) Then
        Label4(i).Caption = "o"
        End If
         Next
For i = 0 To 26
If Left(Text1(2).Text, i + 1) <> Left(lb(2).Caption, i + 1) Then
       Label4(i).Caption = "x"
        End If
        Next
For i = 0 To 26
If Left(Text1(3).Text, i + 1) = Left(lb(3).Caption, i + 1) Then
        Label5(i).Caption = "o"
        End If
         Next
For i = 0 To 26
If Left(Text1(3).Text, i + 1) <> Left(lb(3).Caption, i + 1) Then
       Label5(i).Caption = "x"
        End If
        Next
End Sub

Private Sub Timer2_Timer()
gg = gg - 1
Label6.Caption = gg
If Label6.Caption = 0 Then
mumain.Agent1.Characters("Genie").Speak "½Ã°£ÀÌ ´Ù µÇ¾ú½À´Ï´Ù."
MsgBox "½Ã°£ÀÌ ´Ù µÇ¾ú½À´Ï´Ù."
Text1(0).SetFocus
Timer1.Enabled = False
For i = 0 To 3
lb(i).Visible = False
Text1(i).Visible = True
lb(i).Enabled = False
Text1(i).Enabled = True
Text1(i).Text = ""
Next
For i = 0 To 26
Label2(i).Visible = False
Label3(i).Visible = False
Label4(i).Visible = False
Label5(i).Visible = False
Next
Timer2.Enabled = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
End If

End Sub
