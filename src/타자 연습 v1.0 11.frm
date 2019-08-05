VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   0  '없음
   Caption         =   "Form11"
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   LinkTopic       =   "Form11"
   ScaleHeight     =   3765
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   2160
      MousePointer    =   2  '십자형
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   1440
      MousePointer    =   2  '십자형
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   2040
      MousePointer    =   2  '십자형
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   1800
      MousePointer    =   2  '십자형
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   1800
      MousePointer    =   2  '십자형
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   0
      Picture         =   "타자 연습 v1.0 11.frx":0000
      Top             =   0
      Width           =   5250
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

Form5.Text2.Visible = False
Form5.Text2.Enabled = False
긴글 = "약기운이 차츰 소진해 가는 마취 상태에서처럼 몽롱한 의식을 후두둑 털어 내며 그녀는 눈을 떴다.  희고 검은 빛깔의 물고기 형상을 하고 균일한 분포로 판박이된 천정의 사방 연속 무늬가 어슴푸레 공중에 걸려 있는 게 맨 먼저 시야에 들어왔다. 현관 바깥에 매달린 외등에서 가느다란 불빛이 유리창으로 새어 들어와 맞은편 벽면으로 날이 잘 다듬어진 비수처럼 음험한 그림자를 드리우고 있었다. 그녀는 메말라 껄끄러운 눈꺼풀을 몇 번인가 깜박거리며 눈의 초점을 맞추려 애를 썼다.  뚜걱, 뚜걱, 뚜거덕.  불현듯 온몸의 털구멍이 한꺼번에 바짝 아가리를 닫고 수축되어 버리는 듯한 긴장감. 그녀는 전신이 풀먹인 무명베처럼 빳빳하게 굳어 가는 느낌이 들었다. 발 소리는 역시 이츰에서 들려 오고 있었다. 두 뼘도 채 못 되는 청정의 콘크리트 두께를 뚫고 발소리는 분명히 그녀의 귀에까지 전달되고 있었다.  뚜걱, 뚜거덕, 뚜걱.  구두 밑창의 두꺼운 뒷축이 시멘트 바닥에 맞부딪쳐서 내는 둔중한 마찰음. 군화를 신고 있거나 굽 높은 투박한 등산화를 신었는지도 모른다. 발소리는 이 날따라 유난히 크도 대담하게 울리는 것 같았다.  아니, 그건 언제나 그랬다.  < 끝 >  "

For i = 0 To 3
Form5.lb(i).Caption = mid(긴글, k, 27)
k = k + 27
Next
Form5.Timer1.Enabled = False
Form5.Show
Unload Me

End Sub

Private Sub Label2_Click()
Form5.Text2.Visible = False
Form5.Text2.Enabled = False
긴글 = "쫓겨가는 한 마리 딱정벌레처럼 트럭은 저만큼 들판 가운데로 난 황토길을 따라 느릿느릿 기어가고 있었다. 고르지 못한 노면에서 바퀴가 투어 오를 때마다 덜컹거리는 쇳소리가 들려 왔고 꽁무니로 부옇게 마른 먼지가 피어 올랐다.  덮개 없는 트럭의 뒷간에 홀로 쭈그려 앉은 채 실려 가고 있는 녀석의 모습이 유난히도 자그맣게 오므라들어 있어 보였다. 뒷간에 적대된 알루미늄 식깡들이 이따금 섬뜩할이만큼 차가운 금속성의 광선을 되쏘곤 했다. 풀잎들이 저마다 윤기를 잃어 가고 있는 들녘과 차츰 잿빛으로 퇴색해 가기 시작하는 야산의 정지된 풍경 속에서 그것은 안간힘을 쓰며 집요하게 꿈틀거리고 있는 단 하나의 운동체였다.  '더럽게 운도 없는 녀석이군 전입해 온 지 보름 만에 초상을 치르다니.'  바지를 까내리고 오줌발을 내갈기며 오 일병이 뇌까렸다. 나는 말없이 마른 풀을 짓씹었다. 바로 조금 전에 우리는 그 트럭에서 내렸었다. 야영지를 출발한 지 얼마 되지 않아 차가 마을로 통하는 샛길 입구에 다다랐을 때 선임 탑승자는 차를 세워 우리 둘을 내려 주었던 것이다.  이제 트럭은 들판을 지나 산모퉁이를 마악 꺾어 돌아가려는 참이었다. < 끝 >"

For i = 0 To 3
Form5.lb(i).Caption = mid(긴글, k, 27)
k = k + 27
Next
Form5.Timer1.Enabled = False
Form5.Show
Unload Me

End Sub

Private Sub Label3_Click()
Form5.Text2.Visible = False
Form5.Text2.Enabled = False
긴글 = "막차는 좀처럼 오지 않았다.  별로 복잡한 내용이랄 것도 없는 장부를 마치 꼼꼼히 확인해 보고 나서야 늙은 역장은 돋보기 안경을 벗어 책상 위에 놓고 일어선다.  벌써 삼십 분이나 지났군.  출입문 위쪽에 붙은 낡은 벽시계가 여덟 시 십 오 분을 가리키고 있다. 하긴 뭐 벌써라는 말을 쓰는 것도 새삼스럽다고 그는 고쳐 생각한다. 이렇게 작은 산골 간이역에서 제 시간에 정확히 도착하는 완행 열차를 보기가 그리 쉬운 일은 아님을 익히 알고 있는 탓이다. 더구나 오늘은 눈까지 내리고 있지 않는가.  역장은 손바닥을 비비며 창가로 다가가더니 유리창 너머로 무심히 시선을 던진다. 건널목 옆 외눈박이 수은등이 껑충하게 서서 홀로 눈을 맞으며 희뿌연 얼굴로 땅바닥을 내려다보고 있다. 송이눈이다. 갓난아이의 주먹만한 눈송이들은 어둠저편에 까맣게 숨어 있다가 느닷없이 수은등의 불빛 속에 뛰어 들어오면서 뚱그렇게 놀란 표정을 채 지우지 못한 채 땅바닥으로 곤두박질치고 있다. 굉장한 눈이다. 바람도 그리 없는데 눈발이 비스듬히 비껴날리고 있다. 늙은 역장은 조금은 근심스런 기색으로 유리창에 얼굴을 바짝 대어 본다.     < 끝 >"

For i = 0 To 3
Form5.lb(i).Caption = mid(긴글, k, 27)
k = k + 27
Next
Form5.Timer1.Enabled = False
Form5.Show
Unload Me


End Sub

Private Sub Label4_Click()
Form5.Text2.Visible = False
Form5.Text2.Enabled = False
긴글 = "서울 시내의 제법 번화한 도로변이 대개 그러하듯이 지금 그가 높은 의자에 앉아있는 곳의 대기 속에도 온갖 소리들이 가득 차 있었다. 물론 그 곳은 길거리인지라 차량들의 엔진 소리와 차바퀴가 바닥에 마찰되는 소리, 소리라기보다 소음에 가까운 것들이 주종을 이루고 있긴 했지만 그래도 약간의 주의만 기울인다면 그 외의 더 많은 종류의 소리도 들을 수 있었다. 말하자면 길거리에 앉아 있어야 하는 직업을 가지고 있는 그는 이제 약간의 경력에 의해 시끄러운 자동차들의 소음 속에 파묻혀 있는, 그래서 간간이 단편적으로 드러나는 행인들의 말소리들, 그들의 몸이 내는 소리들, 하여튼 사람들이 만나고 헤어지고 살아가면서 내는 모든 소리들을 놓치지 않고 들을 수 있었다.  그러나 하루 종일 울리는 높운 데시벨의 소리들은 그의 청각 기관을 멍멍하게 했고 머리 속 더 깊이 파고들어서 가끔 두통을 일으켰다. 그러한 시달림 때문인지 아니면 실제로 어떤 이물질 때문인지 귓속이 근지러움을 느낀 그는 한 손에는 배차 시간표를 들고 오른쪽 새끼손가락으로 역시 오른쪽의 귓속을 후볐다. 그러나 귓속의 기묘한 고통은 그의 손가락 끝이 닿을 수 있는 곳 너머에 있었다. < 끝 >"

For i = 0 To 3
Form5.lb(i).Caption = mid(긴글, k, 27)
k = k + 27
Next
Form5.Timer1.Enabled = False
Form5.Show
Unload Me

End Sub

Private Sub Label5_Click()
Form5.Text2.Visible = False
Form5.Text2.Enabled = False
긴글 = "개를 먹는 행위는 일차적으로 소 돼지 양을 먹는 행위와 구별되고, 다른 한편으로 고양이 비둘기 원숭이를 먹는 행위와도 구별된다. 마찬가지로 개가 사람을 보고 짖어 대는 것도 그리 간단한 문제가 아니다. 그것은 소 돼지 양이 사람들에게 울어대는 것과 구별되고, 또한 고양이 원숭이 비둘기가 우는 것과도 구별된다. 이러한 관계가 미묘하다는 것을 동시에 시사하고 있는 것이다.  이런 생각들이 물 속에서 물방울이 보글보글 솟아오르듯이 하릴없이 그에게 떠올랐던 것은 그가 개라는 동물에 쏟은 관심의 덕분이었다. 사실, 생각해 봐야 별로 유쾌할 바 없는 이 동물에 대해 신경을 쓰게 된 데에는 매우 기구하고, 그럴 만한 사연이 있었다. 그것은 그가 개에 도움을 청해야 했던 채무자의 처지에 있었기 때문이었다. 말하자면 그는 개에 허락도 얻지 않고 그 이름을 도용하기 시작했던 것이다.  이 모두는 우선 그의 성격 탓이었다. 그로 말하자면 특징이라곤 담배를 자주 피운다거나, 원색의 옷을 입은 여자를 극도로 싫어한다든가, 걸을 때 고장난 장난감 인형 같은 몸짓이 조금씩 드러나는 것 정도였다. 그러던 그가 어느 날 개성의 필요성에 눈을 뜬 것이다. < 끝 >"
For i = 0 To 3
Form5.lb(i).Caption = mid(긴글, k, 27)
k = k + 27
Next
Form5.Timer1.Enabled = False
Form5.Show
Unload Me

End Sub
