VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  '����
   Caption         =   "�ڸ�����"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   Icon            =   "Ÿ�� ���� V1.0 6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox Text1 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   3840
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�ѱۿ���"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�����"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '����
      Height          =   615
      Left            =   9120
      MousePointer    =   2  '������
      TabIndex        =   10
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '����
      Height          =   735
      Left            =   9000
      MousePointer    =   2  '������
      TabIndex        =   9
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '����
      Height          =   375
      Left            =   9360
      MousePointer    =   2  '������
      TabIndex        =   7
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "�������� :"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '����
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '����
      Caption         =   "�ѱ��ڸ� ������ �Ҷ����� ����� �ϰ� �Է��� �ּ���^-^"
      Height          =   975
      Left            =   840
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Height          =   375
      Left            =   8640
      MousePointer    =   2  '������
      TabIndex        =   1
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   5760
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "Ÿ�� ���� V1.0 6.frx":030A
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim �ѱ��ڸ�(26), a, �����ڸ�(27)


Private Sub Form_Load()
Randomize
Label1.Visible = False
            If mumain.Option12 = True Then
               mumain.Agent1.Characters("Genie").Speak Label5.Caption
             End If

Label1.Visible = False
�ѱ��ڸ�(1) = "��"
�ѱ��ڸ�(2) = "��"
�ѱ��ڸ�(3) = "��"
�ѱ��ڸ�(4) = "��"
�ѱ��ڸ�(5) = "��"
�ѱ��ڸ�(6) = "��"
�ѱ��ڸ�(7) = "��"
�ѱ��ڸ�(8) = "��"
�ѱ��ڸ�(9) = "��"
�ѱ��ڸ�(10) = "��"
�ѱ��ڸ�(11) = "��"
�ѱ��ڸ�(12) = "��"
�ѱ��ڸ�(13) = "��"
�ѱ��ڸ�(26) = "��"
�ѱ��ڸ�(15) = "��"
�ѱ��ڸ�(16) = "��"
�ѱ��ڸ�(17) = "��"
�ѱ��ڸ�(18) = "��"
�ѱ��ڸ�(19) = "��"
�ѱ��ڸ�(20) = "��"
�ѱ��ڸ�(21) = "��"
�ѱ��ڸ�(22) = "��"
�ѱ��ڸ�(23) = "��"
�ѱ��ڸ�(24) = "��"
�ѱ��ڸ�(25) = "��"
�ѱ��ڸ�(26) = "��"
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Command3.Enabled = False
Command4.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
mumain.Agent1.Characters("Genie").Stop

End Sub

Private Sub Label7_Click()
mumain.Show
Unload Me
mumain.Agent1.Characters("Genie").Stop
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Label8_Click()
Label1.Visible = True
Label4.Caption = 0
Text1.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Command3.Enabled = False Then
If Label1.Caption = "��" Then
If KeyCode = vbKeyR Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyS Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyE Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyF Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyA Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyQ Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyT Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyD Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyW Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyC Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyX Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyZ Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyV Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyG Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "��" Then
If KeyCode = vbKeyY Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "��" Then
If KeyCode = vbKeyU Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyI Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyO Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyP Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyH Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyJ Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyK Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyL Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyB Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyN Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "��" Then
If KeyCode = vbKeyM Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If




End If

If Command4.Enabled = False Then
If Label1.Caption = "A" Then
If KeyCode = vbKeyA Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "B" Then
If KeyCode = vbKeyB Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "C" Then
If KeyCode = vbKeyC Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If


If Label1.Caption = "D" Then
If KeyCode = vbKeyD Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "E" Then
If KeyCode = vbKeyE Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "F" Then
If KeyCode = vbKeyF Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "G" Then
If KeyCode = vbKeyG Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "H" Then
If KeyCode = vbKeyH Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "I" Then
If KeyCode = vbKeyI Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "J" Then
If KeyCode = vbKeyJ Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "K" Then
If KeyCode = vbKeyK Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "L" Then
If KeyCode = vbKeyL Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "M" Then
If KeyCode = vbKeyM Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "N" Then
If KeyCode = vbKeyN Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "O" Then
If KeyCode = vbKeyO Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "P" Then
If KeyCode = vbKeyP Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "Q" Then
If KeyCode = vbKeyQ Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "R" Then
If KeyCode = vbKeyR Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "S" Then
If KeyCode = vbKeyS Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "T" Then
If KeyCode = vbKeyT Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "U" Then
If KeyCode = vbKeyU Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "V" Then
If KeyCode = vbKeyV Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "W" Then
If KeyCode = vbKeyW Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "X" Then
If KeyCode = vbKeyX Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
If Label1.Caption = "Y" Then
If KeyCode = vbKeyY Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If

If Label1.Caption = "Z" Then
If KeyCode = vbKeyZ Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1
Text1.Text = ""
Else
Text1.Text = ""
End If
End If
End If
If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If

If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If

If Command3.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �ѱ��ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Command4.Enabled = False Then
If Label1.Caption = "" Then
Label1.Caption = �����ڸ�(a)
Label4.Caption = Label4.Caption + 1

End If
End If
If Label1.Caption = "" Then
If Command3.Enabled = False Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
End If
If Command3.Enabled = False Then
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
End If
If Command4.Enabled = False Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
End If
If Command4.Enabled = False Then
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
End If
End If
End Sub

Private Sub Label2_Click()
Label1.Visible = False
�ѱ��ڸ�(1) = "��"
�ѱ��ڸ�(2) = "��"
�ѱ��ڸ�(3) = "��"
�ѱ��ڸ�(4) = "��"
�ѱ��ڸ�(5) = "��"
�ѱ��ڸ�(6) = "��"
�ѱ��ڸ�(7) = "��"
�ѱ��ڸ�(8) = "��"
�ѱ��ڸ�(9) = "��"
�ѱ��ڸ�(10) = "��"
�ѱ��ڸ�(11) = "��"
�ѱ��ڸ�(12) = "��"
�ѱ��ڸ�(13) = "��"
�ѱ��ڸ�(26) = "��"
�ѱ��ڸ�(15) = "��"
�ѱ��ڸ�(16) = "��"
�ѱ��ڸ�(17) = "��"
�ѱ��ڸ�(18) = "��"
�ѱ��ڸ�(19) = "��"
�ѱ��ڸ�(20) = "��"
�ѱ��ڸ�(21) = "��"
�ѱ��ڸ�(22) = "��"
�ѱ��ڸ�(23) = "��"
�ѱ��ڸ�(24) = "��"
�ѱ��ڸ�(25) = "��"
�ѱ��ڸ�(26) = "��"
a = Int(Rnd(1) * 26)
Label1.Caption = �ѱ��ڸ�(a)
Command3.Enabled = False
Command4.Enabled = True
End Sub

Private Sub Label6_Click()
Text1.IMEMode = vbIMEModeAlpha  '����
Label1.Visible = False
�����ڸ�(1) = "A"
�����ڸ�(2) = "B"
�����ڸ�(3) = "C"
�����ڸ�(4) = "D"
�����ڸ�(5) = "E"
�����ڸ�(6) = "F"
�����ڸ�(7) = "G"
�����ڸ�(8) = "H"
�����ڸ�(9) = "I"
�����ڸ�(10) = "J"
�����ڸ�(11) = "K"
�����ڸ�(12) = "L"
�����ڸ�(13) = "M"
�����ڸ�(14) = "N"
�����ڸ�(15) = "O"
�����ڸ�(16) = "P"
�����ڸ�(18) = "Q"
�����ڸ�(17) = "R"
�����ڸ�(19) = "S"
�����ڸ�(20) = "T"
�����ڸ�(21) = "U"
�����ڸ�(22) = "V"
�����ڸ�(23) = "W"
�����ڸ�(24) = "U"
�����ڸ�(25) = "X"
�����ڸ�(26) = "Y"
�����ڸ�(27) = "Z"
a = Int(Rnd(1) * 27)
Label1.Caption = �����ڸ�(a)
Command3.Enabled = True
Command4.Enabled = False
End Sub
