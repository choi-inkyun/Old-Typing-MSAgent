Attribute VB_Name = "Module1"
Public Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, _
        ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long

Public Const MB_OK = &H0&
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK

Public uOldTime As Integer
Public uTime As Integer
Public ±ä±Û, k, rr, gg
