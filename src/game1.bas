Attribute VB_Name = "Module1"
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public lastleft As Long
Public lasttop As Long
Public 오전오후(9), 메세지(9) As String
Public 시간(9), 분(9), 초(9) As Byte

