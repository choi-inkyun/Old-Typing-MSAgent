VERSION 5.00
Begin VB.Form mumain 
   BackColor       =   &H80000009&
   Caption         =   "Ÿ�� ���α׷�"
   ClientHeight    =   3765
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5085
   DrawStyle       =   3  '���-��
   LinkTopic       =   "Form2"
   ScaleHeight     =   3765
   ScaleWidth      =   5085
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.PictureBox Picture1 
      Height          =   8535
      Left            =   0
      Picture         =   "Ÿ�ڿ��� V1.0.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   0
      Width           =   10335
   End
   Begin VB.Menu mufile 
      Caption         =   "����(&f)"
      Begin VB.Menu muend 
         Caption         =   "������"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mugame 
      Caption         =   "����(&g)"
      Begin VB.Menu mu111 
         Caption         =   "�̽���"
         Shortcut        =   ^A
      End
      Begin VB.Menu musdsd 
         Caption         =   "������"
         Shortcut        =   ^D
      End
      Begin VB.Menu mueksdj 
         Caption         =   "�ܾ����"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu muxkwk 
      Caption         =   "Ÿ��(&s)"
      Begin VB.Menu murmf 
         Caption         =   "ª���� ����"
         Shortcut        =   ^I
      End
      Begin VB.Menu mulong 
         Caption         =   "��� ����"
         Shortcut        =   ^L
      End
      Begin VB.Menu murjawjd 
         Caption         =   "Ÿ�� ����"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu muhelp 
      Caption         =   "����(&h)"
      Begin VB.Menu mu111help 
         Caption         =   "�̽��� ����"
         Shortcut        =   ^H
      End
      Begin VB.Menu mume 
         Caption         =   "������"
      End
      Begin VB.Menu muhelp1 
         Caption         =   "����"
      End
      Begin VB.Menu muver 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "mumain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub mu111_Click()
Form1.Show
End Sub

Private Sub mu111help_Click()
MsgBox "< �̽��� ���� >" + Chr(13) + Chr(13) _
       + "�̽��� ���� ������ ���α׷� �Դϴ�." + Chr(13) + Chr(13) _
       + "ó���� �����Ҷ� �� �ؿ��ִ� �ѱ۰� �����߿��� ������ �ּ���." + Chr(13) + Chr(13) _
       + "���̵� ������ ���ּ���. �׸��� ���α׷��� ���� ������ �������ϴ�." + Chr(13) + Chr(13) _
       + "��������� ������ ������ ������. ������ ��� �������� �����Դϴ�." + Chr(13) + Chr(13) _
       + "���᰹���� 10�� �ö󰥶����� �ӵ��� �������ϴ�." + Chr(13) + Chr(13) _
       + "��̰� �ϼ���..^^" + Chr(13) + Chr(13)
       
       

End Sub

Private Sub mueksdj_Click()
Form8.Show
End Sub

Private Sub muend_Click()
    a = MsgBox("�����Ͻðڽ��ϱ�?", vbCritical + vbYesNo, "����")
    
    If a = vbYes Then
        End
    Else
        Exit Sub
    End If

End Sub

Private Sub muhelp1_Click()
MsgBox "�� ���α׷��� ������ ���� �̿��Ҽ� �ְ�" + Chr(13) + Chr(13) _
       + "��������ϴ�.^^ �׷��� ������ �����ϰڽ��ϴ�." + Chr(13) + Chr(13) _
       + "�ñ��Ͻ����� �����ø� ���Ϸ� �����ּ���."
End Sub

Private Sub mulong_Click()
Form4.Show
End Sub

Private Sub mume_Click()
MsgBox "�ȳ��ϼ���..^^" + Chr(13) + Chr(13) _
                  + "�� ���α׷��� �������� ���α� �̶�� �մϴ�." + Chr(13) + Chr(13) _
                  + "��� �� ��§ ���α׷� ������ �� ���ּ���." + Chr(13) + Chr(13) _
                  + "�� �Ұ��� �Ҳ���." + Chr(13) + Chr(13) _
                  + "86�� 2��22�� �¾���, �ҳ��Դϴ�." + Chr(13) + Chr(13) _
                  + "�׸��� ���� �ɰ����б� 3�г⿡ �������Դϴ�" + Chr(13) + Chr(13) _
                  + "�̸��� = dingpong@hitel.net" + Chr(13) + Chr(13) _
                  + "�׷� ����ְ� ��⼼��..^^" + Chr(13) + Chr(13)
End Sub

Private Sub murjawjd_Click()
Form5.Show
End Sub

Private Sub murmf_Click()
Form3.Show
End Sub

Private Sub musdsd_Click()
Form7.Show
End Sub

Private Sub muver_Click()
muloag.Show
End Sub

