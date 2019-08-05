VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  '없음
   Caption         =   "Form5"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   ClipControls    =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   4800
      Left            =   0
      ScaleHeight     =   4800
      ScaleWidth      =   4125
      TabIndex        =   2
      Top             =   0
      Width           =   4125
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1080
         Top             =   840
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "일정 관리 창"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   1215
      Left            =   1680
      ScaleHeight     =   1215
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   4230
      Left            =   5400
      ScaleHeight     =   4230
      ScaleWidth      =   3030
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   3030
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FoxAlphaBlend Lib "FoxCBmp3.dll" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal alpha As Byte, Optional ByVal MaskColor As Long, Optional ByVal Flags As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal HDC As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim CurX As Single, CurY As Single
Dim WH As Long, WD As Long

Private Sub Form_Load()
Picture3.Picture = LoadPicture(App.Path & "\image\aa.bmp")
    Width = Picture3.Width
    Height = Picture3.Height
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '   Picture1_MouseDown Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '  Picture1_MouseMove Button, Shift, X, Y
End Sub

Private Sub Image1_Click()
Unload Me
 End Sub

Private Sub Image2_Click()
  End Sub

Private Sub Image3_Click()

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '   If Button = 1 Then
 '       CurX = X
 '       CurY = Y
 '   End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '  Dim DeltaX As Long, DeltaY As Long
  '  Dim WH As Long, WD As Long
  '  If Button = 1 Then
  '      WH = GetDesktopWindow
  '      WD = GetDC(WH)
  '      DeltaX = X - CurX
  '      DeltaY = Y - CurY
  '      BitBlt Picture2.HDC, 0, 0, Width \ 15, Height \ 15, Picture2.HDC, DeltaX \ 15, DeltaY \ 15, vbSrcCopy
  '      If DeltaX > 0 Then
  '          BitBlt Picture2.HDC, (Width - DeltaX) \ 15, 0, DeltaX \ 15, ScaleHeight \ 15, WD, (Left + Width) \ 15, (Top + DeltaY) \ 15, vbSrcCopy
  '      ElseIf DeltaX < 0 Then
  '          BitBlt Picture2.HDC, 0, 0, -DeltaX \ 15, Height \ 15, WD, (Left + DeltaX) \ 15, (Top + DeltaY) \ 15, vbSrcCopy
  '      End If
  '      If DeltaY > 0 Then
  '          BitBlt Picture2.HDC, 0, (ScaleHeight - DeltaY) \ 15, ScaleWidth \ 15, DeltaY \ 15, WD, (Left + DeltaX) \ 15, (Top + Height) \ 15, vbSrcCopy
  '      ElseIf DeltaY < 0 Then
  '          BitBlt Picture2.HDC, 0, 0, ScaleWidth \ 15, -DeltaY \ 15, WD, (Left + DeltaX) \ 15, (Top + DeltaY) \ 15, vbSrcCopy
  '      End If
  '      Picture2.Refresh
  '      BitBlt Picture1.HDC, 0, 0, Width \ 15, Height \ 15, Picture2.HDC, 0, 0, vbSrcCopy
  '      FoxAlphaBlend Picture1.HDC, 0, 0, Width \ 15, Height \ 15, Picture3.HDC, 0, 0, 128, &HFF00FF, 1
  '      Move Left + DeltaX, Top + DeltaY
  '      Picture1.Refresh
  '      BitBlt Me.HDC, 0, 0, Width \ 15, Height \ 15, Picture1.HDC, 0, 0, vbSrcCopy
  '      Sleep 10
  '      ReleaseDC WH, WD
  '  End If
End Sub

Private Sub Form_Resize()
    Picture1.Move 0, 0, Width, Height
    Picture2.Move 0, 0, Width, Height
    WH = GetDesktopWindow
    WD = GetDC(WH)
    BitBlt Picture2.HDC, 0, 0, Width \ 15, Height \ 15, WD, Left \ 15, Top \ 15, vbSrcCopy
    BitBlt Picture1.HDC, 0, 0, Width \ 15, Height \ 15, Picture2.HDC, 0, 0, vbSrcCopy
    FoxAlphaBlend Picture1.HDC, 0, 0, Width \ 15, Height \ 15, Picture3.HDC, 0, 0, 128, &HFF00FF, 1
    ReleaseDC WH, WD
    Picture2.Refresh
End Sub

Private Sub Picture1_Resize()
    'cmdExit.Move Picture1.Width - cmdExit.Width - 45, 45
    'cmdMinimize.Move cmdExit.Left - cmdMinimize.Width - 30, 45
End Sub

Private Sub Timer1_Timer()
        For i = 0 To 9
        If Form3.Label1(i).Caption = "" Then
        Label2(i).Caption = "예약된 정보가 없습니다"
        Else
        Label2(i).Caption = Form3.Label1(i).Caption
        End If
        Next
        Label1.Caption = "일정 관리 창"

End Sub
