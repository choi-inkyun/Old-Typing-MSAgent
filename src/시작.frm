VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   Icon            =   "����.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   5040
      Top             =   3000
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   187
      Left            =   3820
      TabIndex        =   5
      Top             =   8280
      Width           =   187
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   187
      Left            =   3820
      TabIndex        =   4
      Top             =   8520
      Width           =   187
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4080
      Top             =   3000
   End
   Begin VB.Timer ���� 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3600
      Top             =   3000
   End
   Begin VB.Timer �ö󰡶� 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4560
      Top             =   3000
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   230
      Left            =   6000
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "15"
      ToolTipText     =   "��簡 ��������� �ӵ��� �Է��մϴ�"
      Top             =   9000
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   230
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "15"
      ToolTipText     =   "��簡 ���ܳ����� �ӵ��� �Է��մϴ�"
      Top             =   8880
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   1440
      MaxLength       =   28
      TabIndex        =   0
      Text            =   "���α׷� ������ ���α� ���� ���α׷�"
      ToolTipText     =   "��縦 �Է��մϴ�"
      Top             =   8880
      Width           =   2535
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   9000
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   360
      MousePointer    =   2  '������
      TabIndex        =   10
      Top             =   7600
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   240
      MousePointer    =   2  '������
      TabIndex        =   9
      Top             =   7080
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   360
      MousePointer    =   2  '������
      TabIndex        =   8
      Top             =   6300
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   360
      MousePointer    =   2  '������
      TabIndex        =   7
      Top             =   5800
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   360
      MousePointer    =   2  '������
      TabIndex        =   6
      Top             =   5290
      Width           =   3855
   End
   Begin VB.Image Image4 
      Height          =   4350
      Left            =   45
      Picture         =   "����.frx":030A
      Top             =   4500
      Width           =   4260
   End
   Begin VB.Image Image3 
      Height          =   750
      Left            =   4320
      MousePointer    =   2  '������
      Picture         =   "����.frx":3C876
      Top             =   1200
      Width           =   3000
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   3720
      MousePointer    =   2  '������
      Picture         =   "����.frx":43DEA
      Top             =   360
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   8850
      Left            =   0
      Picture         =   "����.frx":4B35E
      Top             =   0
      Width           =   7500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   840
      X2              =   6240
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Height          =   180
      Left            =   3960
      TabIndex        =   3
      Top             =   7800
      Width           =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aaa As Integer
Dim ����(100)
Dim ����� As Boolean
Dim ���簳��
Dim ���
Dim aa, bb, cc, dd, ee, ff, gg, hh, jj
Dim a, b, c, d, e
Private Sub ��ξ�����()
Dim I As Integer
For I = 1 To ���簳��
Unload Label1(I)
Next I
Timer1.Enabled = True
End Sub
Private Sub ����_Timer()
Static a As Integer
Dim ǥ��
If a = Len(Text1.Text) Then ����.Enabled = False: a = 0: GoTo 1
a = a + 1
ǥ�� = Mid(���, a, 1)
Load Label1(a)
Label1(a).Move Label1(a - 1).Left + Label1(a - 1).Width, Label1(a - 1).Top
Label1(a).ForeColor = RGB(0, 0, 0)
Label1(a).Visible = True
Label1(a).Caption = ǥ��
Label1(a).ForeColor = RGB(0, 0, 0)
�ö󰡶�.Enabled = True
���簳�� = a
If a = 1 And b = 1 And c = 1 And d = 1 And e = 1 Then
    Check1.Value = 1
End If
1 End Sub

Private Sub �ö󰡶�_Timer()
Static �Ϸᰳ��
Static b As Integer
Dim I As Integer
If ����� = True Then GoTo 2 Else GoTo 1
1 For I = 0 To ���簳��
����(I) = ����(I) + Val(Text2.Text)
Label1(I).ForeColor = RGB(����(I), ����(I), ����(I))
If ����(I) >= 255 Then ����(I) = 255: �Ϸᰳ�� = I
Next I
GoTo 3
2 For I = 0 To ���簳��
����(I) = ����(I) - Val(Text3.Text)
If ����(I) <= 0 Then ����(I) = 0: �Ϸᰳ�� = I
Label1(I).ForeColor = RGB(����(I), ����(I), ����(I))
Next I
GoTo 4
3 If �Ϸᰳ�� = Len(���) Then ����� = True: �Ϸᰳ�� = 0: GoTo 5
4 If �Ϸᰳ�� = Len(���) Then �ö󰡶�.Enabled = False: �Ϸᰳ�� = 0: Text1.Enabled = True: Text2.Enabled = True: Text3.Enabled = True: ��ξ�����: ����� = False: ���簳�� = 0: GoTo 5
5 End Sub




Private Sub asAssisPopup8_Click()
End
End Sub

Private Sub Check1_Click()
     Open App.Path & "\üũ.txt" For Input As #1
          Input #1, aaa
     Close #1
Open App.Path & "\üũ.txt" For Output As #1
          If Check1.Value = 1 Then
          aaa = 1
          Write #1, aaa
          End If
          If Check1.Value = 0 Then
          aaa = 0
          Write #1, aaa
          End If
          
    Close #1
End Sub

Private Sub Form_Load()
     Open App.Path & "\üũ.txt" For Input As #1
          Input #1, aaa
          If aaa = 1 Then
        Call Shell(App.Path & "\�ð�����.exe", 1)
          End
          End If
     Close #1
     aa = 0
     bb = 0
     cc = 0
     dd = 0
     ee = 0
     ff = 0
     gg = 0
     hh = 0
     jj = 0
     a = 0
     b = 0
     c = 0
     d = 0
     e = 0
 Form2.Show
' Label2.Caption = "�ȳ��ϼ���.^^.���� ���� ���� ���α׷��� ��ġ �ɼ��Դϴ�." + Chr(13) _
'                + "��ġ �ؾ��Ҳ� ����?.�ʼ� �׸� �ִ°͵��� �� ��ġ�� �ֽð�. �ؿ� �����׸�" + Chr(13) _
'                + "�ִ°��� ��ġ�� ���� �����ŵ� �����Ͻð� �̿��Ͻ� �� ������." + Chr(13) _
'                + "�ٽ��� ������ºκ��� �ʵű� ������ �ǵ����̸����ġ���ּ������ϴ� �ٷ��Դϴ�" + Chr(13) _
'                + "�ؿ�üũ�ϴ°��� üũ�Ͻð�ٽ� �����Ͻø� �� ȭ���̴ٽô� �ʶߴ°��� �����ֽ��ϴ�. �ٽ� �߰� �ϰ�����ø� üũ.txt �� 0���� �ٲ��ּ���.�׷������ϰԾ��ñ�~"
Timer2.Enabled = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If jj = 0 Then
jj = 1
aa = 0
bb = 0
cc = 0
dd = 0
ee = 0
ff = 0
gg = 0
hh = 0
End If
Timer2.Enabled = True
End Sub

Private Sub Image2_Click()
Dim a As Integer
If Check2.Value = 1 Then
Kill App.Path & "\util\MSAgent.exe"
Kill App.Path & "\util\AgtX0412.exe"
Kill App.Path & "\util\Genie.exe"
Kill App.Path & "\util\lhttskok.exe"
Kill App.Path & "\util\spchapi.exe"
End If
a = MsgBox("��ġ�� ��� �ϼ̽��ϱ�?", vbYesNo, "��ġ")
If a = vbYes Then
Check1.Value = 1
     Open App.Path & "\üũ.txt" For Input As #1
          Input #1, aaa
     Close #1
     Open App.Path & "\üũ.txt" For Output As #1
          If Check1.Value = 1 Then
          aaa = 1
          Write #1, aaa
          End If
          If Check1.Value = 0 Then
          aaa = 0
          Write #1, aaa
          End If
          
    Close #1
    Call Shell(App.Path & "\�ð�����.exe", 1)
Unload Form2
End
End If



End Sub

Private Sub Image3_Click()
    a = MsgBox("�����Ͻðڽ��ϱ�?", vbCritical + vbYesNo, "����")
    
    If a = vbYes Then
    End
    End If
End Sub

Private Sub Label2_Click()
Call Shell(App.Path & "\util\MSAgent.exe", 1)
a = 1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If aa = 0 Then
jj = 0
ee = 0
aa = 1
bb = 0
jj = 0
cc = 0
dd = 0
End If
Timer2.Enabled = True

End Sub

Private Sub Label4_Click()
Call Shell(App.Path & "\util\AgtX0412.exe", 1)
b = 1
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer2.Enabled = True
If bb = 0 Then
ee = 0
jj = 0
aa = 0
bb = 1
cc = 0
dd = 0
End If
End Sub

Private Sub Label5_Click()
Call Shell(App.Path & "\util\Genie.exe", 1)
c = 1
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer2.Enabled = True
If cc = 0 Then
ee = 0
jj = 0
aa = 0
bb = 0
cc = 1
dd = 0
End If
End Sub

Private Sub Label6_Click()
Call Shell(App.Path & "\util\lhttskok.exe", 1)
d = 1
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If dd = 0 Then
dd = 1
ee = 0
aa = 0
bb = 0
jj = 0
cc = 0
End If
Timer2.Enabled = True

End Sub

Private Sub Label7_Click()
Call Shell(App.Path & "\util\spchapi.exe", 1)
e = 1
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ee = 0 Then
ee = 1
aa = 0
bb = 0
cc = 0
jj = 0
dd = 0
End If
Timer2.Enabled = True

End Sub

Private Sub Timer1_Timer()
  If Len(Text2.Text) = 0 Then MsgBox "�ӵ� 1�� ���� �־��ּ���.", vbExclamation, "����": GoTo 2
2 If Len(Text3.Text) = 0 Then MsgBox "�ӵ� 2�� ���� �־��ּ���.", vbExclamation, "����": GoTo 3

 ��� = Text1.Text
 ����.Enabled = True
 Text1.Enabled = False
 Text2.Enabled = False
 Text3.Enabled = False
Timer1.Enabled = False

3 End Sub

Private Sub Timer2_Timer()
If aa = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\a1.bmp")
aa = 3
End If
If bb = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\a2.bmp")
bb = 3
End If
If cc = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\a3.bmp")
cc = 3
End If
If dd = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\a4.bmp")
dd = 3
End If
If ee = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\a5.bmp")
ee = 3
End If
If aa = 0 And bb = 0 And cc = 0 And dd = 0 And ee = 0 And jj = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\2.bmp")
jj = 3
End If

Timer2.Enabled = False



End Sub
