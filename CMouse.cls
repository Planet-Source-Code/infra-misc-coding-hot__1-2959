VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Property Get X() As Single
    Dim tmpPnt As POINTAPI
    GetCursorPos tmpPnt
    X = CSng(tmpPnt.X)
End Property

Property Let X(ByVal New_X As Single)
    Dim tmpPnt As POINTAPI
    GetCursorPos tmpPnt
    SetCursorPos CLng(New_X), tmpPnt.Y
End Property

Property Get Y() As Single
    Dim tmpPnt As POINTAPI
    GetCursorPos tmpPnt
    Y = tmpPnt.Y
End Property

Property Let Y(ByVal New_Y As Single)
    Dim tmpPnt As POINTAPI
    GetCursorPos tmpPnt
    SetCursorPos tmpPnt.X, CLng(New_Y)
End Property

Function WindowOver(ByVal X As Single, ByVal Y As Single) As String
    Dim sBuffer As String * 255, sLen As Long
    sLen = GetWindowText(WindowFromPoint(CLng(X), CLng(Y)), sBuffer, 255)
    WindowOver = Left$(sBuffer, sLen)
End Function
