VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "����"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "Ÿ�ڿ��� V1.0 2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Timer Timer1 
      Interval        =   2100
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu muhideshow 
         Caption         =   "���� - ���̱�"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu mumenu 
         Caption         =   "����"
         Shortcut        =   ^T
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
      Begin VB.Menu muend 
         Caption         =   "�����ϱ�"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ss As Integer
Private Sub Form_Load()
Timer1.Enabled = False
End Sub

Private Sub muend_Click()
mumain.Agent1.Characters("Genie").Stop
    
    ss = MsgBox("�����Ͻðڽ��ϱ�?", vbCritical + vbYesNo, "����")
    
    If ss = vbYes Then
    Timer1.Enabled = True
    mumain.Timer1.Enabled = True
    mumain.Agent1.Characters("Genie").Play "GetAttentionContinued"
        
    Else
        Exit Sub
    End If

End Sub

Private Sub muhideshow_Click()


    If mumain.Agent1.Characters("Genie").Visible Then
       mumain.Agent1.Characters("Genie").Hide
    Else
        mumain.Agent1.Characters("Genie").Show
 End If
End Sub

Private Sub mumenu_Click()

mumain.Agent1.Characters("Genie").Stop
mumain.Agent1.Characters("Genie").Speak "�������α׷��� �����մϴ�."
Timer1.Enabled = True
End Sub

Private Sub mutime_Click()
End Sub

Private Sub Timer1_Timer()
Call Shell(App.Path & "\�ð�����.exe", 1)
Timer1.Enabled = False
End
End Sub

Private Sub Timer2_Timer()
End Sub
