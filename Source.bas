Attribute VB_Name = "SourceMod"
'Source! Code
'By:  InfraRed
'Comments:  I hope you like my source code, if you
'notice anything that has been copied from other
'source code, then it must have been used in one
'of my applications which I copied all of this
'from directly.  This is all in sections plus with
'comments saying what the code does in every
'sub/function, for all of you newbies who want
'to learn lots of stuff fast.  Most of you who
'will use this source code probably will want to
'use it in some program you come up with.  Will
'you please give me a little credit if you do?
'I put in a lot of easy code, plus some harder
'source code.  Enjoy.
'Contacting Me:
'E-Mail:  InfraRed@flashmail.com
'ICQ:  17948286 (UIN)

'-------------------------------------------------------

'Sub Titles of all source code in Source.bas:

'Section 1 (Declarations):
'Global Declarations
'Other Declarations

'Section 2:
'FileSave
'FileOpen
'ListSave
'ListOpen

'Section 3:
'MakeDir
'DeleteDir
'DelFilesInDir

'Section 4:
'MoveFile
'CopyFile
'DeleteFile
'ExecuteFile

'Section 5:
'Encrypt
'Decrypt
'BitEncrypt
'BitDecrypt
'SuperEE (Private)

'Section 6:
'DisableCtrlAltDel
'EnableCtrlAltDel
'HideCtrlAltDel
'ShowCtrlAltDel

'Section 7:
'OpenCD
'CloseCD
'PrintBlankPage
'PrintText
'PrintPage (Private)
'PrintNewPage (Private)
'PrintEndOfLastPage (Private)

'Section 8:
'MakeStartupReg
'AddToStartupDir
'MakeRegFile (Private)

'Section 9:
'Ontop
'NotOntop
'InvisibleForm
'HoleInForm

'Section 10:
'ClipboardCopy
'ClipboardGet
'ClearClipboard

'Section 11:
'Ping
'ConvertIPAddressToLong (Private)

'Section 12:
'Code1
'Code2
'Decode1
'Decode2
'ReplaceC (Private)

'Section 13:
'Add
'Subtract
'Divide
'Multiply
'ToPower
'ToRoot
'FractionToDecimal
'DecimalToPercentage
'PercentageToDecimal
'AreaOfCircle
'Circumference
'AreaOfSquare
'PerimeterOfSquare
'PerimeterOfRectangle
'AreaOfRectangle
'AreaOfTriangle
'PerimeterOfTriangle
'PerimeterOf4SidedPolygon
'VolumeOfCube
'VolumeOfPrism
'VolumeOfSphere
'VolumeOfPyramid
'VolumeOfCone
'VolumeOfCylinder

'Section 14:
'FadeThreeColorHTML
'FadeTwoColorHTML
'FadeThreeColorYahoo
'FadeTwoColorYahoo
'FadeThreeColorANSI
'FadeTwoColorANSI

'Section 15:
'RestartWindows
'ExitWindows
'RebootComputer

'Section 16:
'AltCaps
'BackwardsText
'EliteType
'SpaceCharacters
'DoubleCharacters
'EchoText
'Scramble
'TwistText

'Section 17:
'GetAppVersion
'GetAppName
'GetAppPath
'GetAppDescription
'GetAppCopyRight
'GetAppComment
'GetAppTitle
'GetAppCompanyName
'GetAppProductName

'Section 18:
'MoveMouse
'MousePosition
'LeftClick
'LeftDown
'LeftUp
'MiddleClick
'MiddleDown
'MiddleUp
'RightClick
'RightDown
'RightUp

'Section 19:
'DrawSquareOnForm
'DrawLineOnForm
'DrawSquareOnPictureBox
'DrawLineOnPictureBox

'Section 20:
'ConvertRGBToHex
'RGBToHex (Private)
'ConvertHexToRGB
'HexToRGB (Private)
'WebPage
'RandomNumber
'MakeInputBox
'LengthOfString
'FindAsciiOfChr
'MakeChrFromAscii
'MakeRndChrString
'DoSendKeys
'GetTextFromListBox
'GetTextFromComboBox
'PasswordLock
'ChangeDefaultDir
'ChangeDefaultDrive
'MakeRegistrySetting


'# Of Subs:  127

'-------------------------------------------------------

'Section 1:  Declarations

'Global Declarations
Global MouseDown As Boolean
Global MouseOver As Boolean
Global Mouse As New CMouse
Global s(52) As String
Global pi As Long
Global NumLinesOnPageToPrint As Integer
Global FirstPageNum As Integer
Global NextPageNum As Integer
Global LineNum As Integer
Global CheckThisLineNum As Integer
Global NumLines As Integer
Global TotalPageCount As Integer

'Other Declarations
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10
Public Const MOUSEEVENTF_MOVE = &H1
Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Type POINTAPI
X As Long
Y As Long
End Type
Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Declare Function ExitWindows Lib "User" (ByVal dwReturnCode As Long, ByVal uReserved As Integer) As Integer
Global Const EW_REBOOTSYSTEM = &H43
Global Const EW_RESTARTWINDOWS = &H42
Global Const EW_EXITWINDOWS = 0
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SCREENSAVERRUNNING = 97
Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Boolean
End Type
Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Const IP_STATUS_BASE = 11000
Private Const IP_SUCCESS = 0
Private Const IP_BUF_TOO_SMALL = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Private Const IP_NO_RESOURCES = (11000 + 6)
Private Const IP_BAD_OPTION = (11000 + 7)
Private Const IP_HW_ERROR = (11000 + 8)
Private Const IP_PACKET_TOO_BIG = (11000 + 9)
Private Const IP_REQ_TIMED_OUT = (11000 + 10)
Private Const IP_BAD_REQ = (11000 + 11)
Private Const IP_BAD_ROUTE = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Private Const IP_PARAM_PROBLEM = (11000 + 15)
Private Const IP_SOURCE_QUENCH = (11000 + 16)
Private Const IP_OPTION_TOO_BIG = (11000 + 17)
Private Const IP_BAD_DESTINATION = (11000 + 18)
Private Const IP_ADDR_DELETED = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Private Const IP_MTU_CHANGE = (11000 + 21)
Private Const IP_UNLOAD = (11000 + 22)
Private Const IP_ADDR_ADDED = (11000 + 23)
Private Const IP_GENERAL_FAILURE = (11000 + 50)
Private Const MAX_IP_STATUS = 11000 + 50
Private Const IP_PENDING = (11000 + 255)
Private Type ip_option_information
Ttl             As Byte
Tos             As Byte
Flags           As Byte
OptionsSize     As Byte
OptionsData     As Long
End Type
Private Type icmp_echo_reply
Address         As Long
Status          As Long
RoundTripTime   As Long
DataSize        As Integer
Reserved        As Integer
DataPointer     As Long
Options         As ip_option_information
Data            As String * 250
End Type
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, _
                                                    ByVal DestinationAddress As Long, _
                                                    ByVal RequestData As String, _
                                                    ByVal RequestSize As Integer, _
                                                    RequestOptions As ip_option_information, _
                                                    ReplyBuffer As icmp_echo_reply, _
                                                    ByVal ReplySize As Long, _
                                                    ByVal Timeout As Long) As Long
Private Const PING_TIMEOUT = 200
Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYSSTATUS_LEN = 256
Private Const WSADESCRIPTION_LEN_1 = WSADESCRIPTION_LEN + 1
Private Const WSASYSSTATUS_LEN_1 = WSASYSSTATUS_LEN + 1
Private Const SOCKET_ERROR = -1
Private Type tagWSAData
wVersion            As Integer
wHighVersion        As Integer
szDescription       As String * WSADESCRIPTION_LEN_1
szSystemStatus      As String * WSASYSSTATUS_LEN_1
iMaxSockets         As Integer
iMaxUdpDg           As Integer
lpVendorInfo        As String * 200
End Type
Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequested As Integer, lpWSAData As tagWSAData) As Integer
Private Declare Function WSACleanup Lib "wsock32" () As Integer

'Section 2:  Saving/Opening Files

Public Sub FileSave(Text As String, FilePath As String)
'Save a text file
On Error GoTo error
Dim Directory As String
              Directory$ = FilePath
       On Error GoTo error
       Open Directory$ For Output As #1
           Print #1, Text
       Close #1
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Function FileOpen(FilePath As String)
'Open a text file
On Error GoTo error
Dim Directory As String
Directory$ = FilePath
    Dim MyString As String
       On Error GoTo error
       Open Directory$ For Input As #1
       While Not EOF(1)
           Input #1, FileOpen
           Wend
           Close #1
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Sub ListSave(List As ListBox, FilePath As String)
'Save all data in a list box
On Error GoTo error
Dim i As Integer
Dim Directory As String
              Directory$ = FilePath
       On Error GoTo error
       Open Directory$ For Output As #1
       For i = 0 To List.ListCount - 1
           Print #1, List.List(i)
       Next i
       Close #1
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub ListOpen(List As ListBox, FilePath As String)
'Open saved list box data
On Error GoTo error
Directory$ = FilePath
    Dim MyString As String
       On Error GoTo error
       Open Directory$ For Input As #1
       While Not EOF(1)
           Input #1, MyString$
           DoEvents
               List.AddItem MyString$
           Wend
           Close #1
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

'Section 3:  Deleting/Making Directories

Public Sub MakeDir(DirPath As String)
'Make a directory
On Error GoTo error
MkDir DirPath$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub DeleteDir(DirPath As String)
'Delete a directory
On Error GoTo error
RmDir DirPath$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub DelFilesInDir(DirPath As String, DelDir As Boolean)
'Delete all files in a directory and (optional) delete the directory too
On Error GoTo error
Kill DirPath$ & "*.*"
If DelDir = True Then
RmDir DirPath$
End If
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

'Section 4:  Copying/Moving/Executing/Deleting Files

Public Sub MoveFile(StartPath As String, EndPath As String)
'Move a file
On Error GoTo error
FileCopy StartPath$, EndPath$
Kill StartPath$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub CopyFile(StartPath As String, EndPath As String)
'Copy a file
On Error GoTo error
FileCopy StartPath$, EndPath$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub DeleteFile(FilePath As String)
'Delete a file
On Error GoTo error
Kill FilePath$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub ExecuteFile(FilePath As String)
'Execute a file
On Error GoTo error
ret = Shell("rundll32.exe url.dll,FileProtocolHandler " & (FilePath))
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

'Section 5:  Encryption/Decryption

Function Encrypt(Start As Integer, diff As Integer, beta As Integer, alpha As Integer, times As Integer, SuperEncrypt As Boolean, Text As String)
'Encrypt characters
On Error GoTo error
Dim i As Integer
Dim curkey As Long
Dim m As Long
Dim endstr As String
Dim Text2 As String
Dim lesser As Double
Dim larger As Double
Dim SuperE As Boolean
Dim a As Integer
SuperE = SuperEncrypt
If diff > 500 Then
diff = 500
ElseIf diff < 1 Then
diff = 1
End If
If times > 100 Then
times = 100
ElseIf times < 1 Then
times = 1
End If
If Start > 255 Then
Start = 255
ElseIf Start < 1 Then
Start = 1
End If
If beta > 5 Then
beta = 5
ElseIf beta < 1 Then
beta = 1
End If
If alpha > 5 Then
alpha = 5
ElseIf alpha < 1 Then
alpha = 1
End If
curkey = Start
curkey = (curkey * alpha) / beta
  If SuperE = True Then
    If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
    curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
    Else
    curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
    End If
  curkey = SuperEE(curkey, beta, alpha, beta)
  End If
  If curkey > 255 Then
  curkey = 255 - (curkey / 255)
  ElseIf curkey < 0 Then
  curkey = 0 - (curkey / 255)
  End If
For a = 1 To times
For i = 1 To Len(Text)
    If 255 - curkey > curkey Then
    larger = 255 - curkey
    lesser = curkey
    Else
    larger = curkey
    lesser = 255 - curkey
    End If
  If Asc(Mid$(Text, i, 1)) <= lesser Then
  m = Asc(Mid$(Text, i, 1)) + (larger - 1)
  endstr = endstr + Chr$(m)
  Else
  m = Asc(Mid$(Text, i, 1)) - lesser
  endstr = endstr + Chr$(m)
  End If
curkey = curkey + diff
  If curkey > 255 Then
  curkey = curkey - 255
  End If
curkey = (curkey * alpha) / beta
  If SuperE = True Then
    If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
    curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
    Else
    curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
    End If
  curkey = SuperEE(curkey, beta, alpha, beta)
  End If
beta = beta + (2 * diff)
alpha = alpha + diff
  If beta > 5 Then
  beta = 1
  End If
  If alpha > 5 Then
  alpha = 1
  End If
  If curkey > 255 Then
  curkey = 255 - (curkey / 255)
  ElseIf curkey < 0 Then
  curkey = 0 - (curkey / 255)
  End If
  If diff > 500 Then
  diff = 1
  Else
  diff = diff + diff
  End If
Next i
Text2 = ""
Text2 = endstr
endstr = ""
Next a
Encrypt = Text2
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Decrypt(Start As Integer, diff As Integer, beta As Integer, alpha As Integer, times As Integer, SuperEncrypt As Boolean, Text As String)
'Decrypt characters
On Error GoTo error
Dim i As Integer
Dim curkey As Long
Dim m As Long
Dim endstr As String
Dim Text2 As String
Dim lesser As Double
Dim larger As Double
Dim SuperE As Boolean
Dim a As Integer
SuperE = SuperEncrypt
If diff > 500 Then
diff = 500
ElseIf diff < 1 Then
diff = 1
End If
If times > 100 Then
times = 100
ElseIf times < 1 Then
times = 1
End If
If Start > 255 Then
Start = 255
ElseIf Start < 1 Then
Start = 1
End If
If beta > 5 Then
beta = 5
ElseIf beta < 1 Then
beta = 1
End If
If alpha > 5 Then
alpha = 5
ElseIf alpha < 1 Then
alpha = 1
End If
curkey = Start
curkey = (curkey * alpha) / beta
  If SuperE = True Then
    If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
    curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
    Else
    curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
    End If
  curkey = SuperEE(curkey, beta, alpha, beta)
  End If
  If curkey > 255 Then
  curkey = 255 - (curkey / 255)
  ElseIf curkey < 0 Then
  curkey = 0 - (curkey / 255)
  End If
For a = 1 To times
For i = 1 To Len(Text)
    If 255 - curkey > curkey Then
    larger = 255 - curkey
    lesser = curkey
    Else
    larger = curkey
    lesser = 255 - curkey
    End If
  If Asc(Mid$(Text, i, 1)) >= larger Then
  m = Asc(Mid$(Text, i, 1)) - (larger - 1)
  endstr = endstr + Chr$(m)
  Else
  m = Asc(Mid$(Text, i, 1)) + lesser
  endstr = endstr + Chr$(m)
  End If
curkey = curkey + diff
  If curkey > 255 Then
  curkey = curkey - 255
  End If
curkey = (curkey * alpha) / beta
  If SuperE = True Then
    If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
    curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
    Else
    curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
    End If
  curkey = SuperEE(curkey, beta, alpha, beta)
  End If
beta = beta + (2 * diff)
alpha = alpha + diff
  If beta > 5 Then
  beta = 1
  End If
  If alpha > 5 Then
  alpha = 1
  End If
  If curkey > 255 Then
  curkey = 255 - (curkey / 255)
  ElseIf curkey < 0 Then
  curkey = 0 - (curkey / 255)
  End If
  If diff > 500 Then
  diff = 1
  Else
  diff = diff + diff
  End If
Next i
Text2 = ""
Text2 = endstr
endstr = ""
Next a
Decrypt = Text2
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function BitEncrypt(Text As String, Key As String)
'This will encrypt a string, using the ascii character code of another string (Key$)
On Error GoTo error
Dim CurPos As Long
Dim i As Long
Dim endstr As String
Dim chrasc As Long
CurPos = 1
For i = 1 To Len(Text$)
chrasc = Asc(Mid$(Text$, i, 1)) + Asc(Mid$(Key$, CurPos, 1))
  If chrasc > 255 Then
  chrasc = chrasc - 255
  End If
endstr$ = endstr$ & Chr$(chrasc)
  If CurPos = Len(Key$) Then
  CurPos = 1
  Else
  CurPos = CurPos + 1
  End If
Graph2 Len(Text$), (i)
Next i
BitEncrypt = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function BitDecrypt(Text As String, Key As String)
'This will decrypt a string, using the ascii character code of another string (Key$)
On Error GoTo error
Dim CurPos As Long
Dim i As Long
Dim endstr As String
Dim chrasc As Long
CurPos = 1
For i = 1 To Len(Text$)
chrasc = Asc(Mid$(Text$, i, 1)) - Asc(Mid$(Key$, CurPos, 1))
  If chrasc < 1 Then
  chrasc = chrasc + 255
  End If
endstr$ = endstr$ & Chr$(chrasc)
  If CurPos = Len(Key$) Then
  CurPos = 1
  Else
  CurPos = CurPos + 1
  End If
Graph2 Len(Text$), (i)
Next i
RndBitD = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Private Function SuperEE(curkey As Long, beta As Integer, alpha As Integer, times As Integer)
'For encryption:  Change the current key around more
On Error GoTo error
curkey = (((curkey / times) - (beta + times)) * alpha) + ((beta / alpha) - times)
If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
Else
curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
End If
If beta - times = 0 Then
curkey = ((curkey * alpha) + (beta * times))
Else
curkey = ((curkey * (beta - times)) + (beta - times))
  If curkey < 0 Then
  curkey = curkey + (alpha + beta)
  ElseIf curkey = 0 Then
  curkey = curkey + (alpha + times)
  Else
  curkey = curkey + (beta + times)
  End If
End If
SuperEE = curkey
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

'Section 6:  Ctrl + Alt + Del Stuff

Public Sub DisableCtrlAltDel()
'Disable Ctrl + Alt + Del
On Error GoTo error
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub EnableCtrlAltDel()
'Enable Ctrl + Alt + Del
On Error GoTo error
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub HideCtrlAltDel()
'Hide this app from Ctrl + Alt + Del
On Error GoTo error
App.TaskVisible = False
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub ShowCtrlAltDel()
'Show this app in Ctrl + Alt + Del
On Error GoTo error
App.TaskVisible = True
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

'Section 7:  External Stuff (Printer/CD)

Public Sub OpenCD()
'Open the CD drive
On Error GoTo error
retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub CloseCD()
'Close the CD drive
On Error GoTo error
retvalue = mciSendString("set CDAudio door closed", returnstring, 127, 0)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub PrintBlankPage()
'Print a blank page out of a printer
On Error GoTo error
Printer.NewPage
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub PrintText(Text As String, MarginSize As Integer, AmountOfChrsInOneLine As Integer, JustUseDefault As Boolean)
'This will print the text out of the default printer
On Error Resume Next
Screen.MousePointer = 11
If JustUseDefault = True Then
MarginSize = 10
AmountOfChrsInOneLine = 65
End If
NumLinesOnPageToPrint = 60
If NextPageNum% > 0 Then NextPageNum% = 0
NextPageNum% = FirstPageNum% + NextPageNum% + 1
TotalPageCount% = 1
Call PrintPage(Text$, MarginSize, AmountOfChrsInOneLine)
PrintEndOfLastPage
Screen.MousePointer = 0
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub PrintPage(TextString, Margin As Integer, Length_ChrsInlineOfText As Integer)
'For Print Text:  This will print a page of the text out of the printer
On Error Resume Next
Dim ChrPosition
Dim AllChrsInThisLineOfText
Dim PlaceInLineOfText As Integer
ChrPosition = 1
Printer.FontSize = 18
Printer.Print Tab(MarginSize%);
LineNum% = 1
Do While ChrPosition < Len(TextString)
AllChrsInThisLineOfText = Mid$(TextString, ChrPosition, Length_ChrsInlineOfText%)
If ChrPosition + Len(AllChrsInThisLineOfText) < Len(TextString) Then
For PlaceInLineOfText% = Len(AllChrsInThisLineOfText) To 1 Step -1
If Mid$(AllChrsInThisLineOfText, PlaceInLineOfText%, 1) = Chr$(32) Then
CheckThisLineNum% = 1
PrintNewPage
If InStr(1, AllChrsInThisLineOfText, Chr$(10), 1) > 0 Then
CheckThisLineNum% = 1
PrintNewPage
PlaceInLineOfText% = InStr(1, AllChrsInThisLineOfText, Chr$(10), 1)
LineNum% = LineNum% + 1
End If
If Mid$(TextString, ChrPosition, PlaceInLineOfText%) <> Chr$(13) + Chr$(10) Then
Printer.Print Tab(MarginSize%);
Printer.Print Mid$(TextString, ChrPosition, PlaceInLineOfText%)
LineNum% = LineNum% + 1
Else
LineNum% = LineNum% - 1
End If
ChrPosition = ChrPosition + PlaceInLineOfText%
PlaceInLineOfText% = 0
End If
Next
Else
CheckThisLineNum% = 1
PrintNewPage
Printer.Print Tab(Margin%);
Printer.Print AllChrsInThisLineOfText
ChrPosition = Len(TextString)
LineNum% = LineNum% + 1
End If
Loop
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub PrintNewPage()
'For Print Text:  This will begin a new page to print the text out of the printer
On Error Resume Next
If LineNum% + CheckThisLineNum% >= NumLinesOnPageToPrint% Then
Printer.Print ""
Printer.Print Tab(MarginSize%);
Printer.Print "(continued on page " + CStr(NextPageNum%) + ")"
Printer.NewPage
TotalPageCount% = TotalPageCount% + 1
Printer.Print Tab(MarginSize%);
Printer.Print "Page " + CStr(NextPageNum%)
Printer.Print ""
Printer.Print ""
NextPageNum% = NextPageNum% + 1
LineNum% = 3
End If
CheckThisLineNum% = 0
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub PrintEndOfLastPage()
'For Print Text:  This will print the end of the last page out of the printer
On Error Resume Next
If LineNum% + 2 > NumLinesOnPageToPrint% Then
Printer.NewPage
TotalPageCount% = TotalPageCount% + 1
Printer.Print Tab(MarginSize%);
Printer.Print "Page " + CStr(NextPageNum%)
Printer.Print ""
Printer.Print ""
Printer.Print Tab(MarginSize%);
Else
Printer.Print ""
Printer.Print Tab(MarginSize%);
End If
Printer.EndDoc
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

'Section 8:  Startup

Public Sub MakeStartupReg(AppTitle As String)
'Add your application to windows startup registry
On Error GoTo error
a = MakeRegFile(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", AppTitle$, App.Path & "\" & App.EXEName & ".exe")
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub AddToStartupDir()
'Add your application to the windows startup folder
On Error GoTo error
FileCopy App.Path & "\" & App.EXEName & ".EXE", Mid$(App.Path, 1, 3) & "WINDOWS\START MENU\PROGRAMS\STARTUP\" & App.EXEName & ".EXE"
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Function MakeRegFile(ByVal hKey As Long, ByVal lpszSubKey As String, ByVal sSetValue As String, ByVal sValue As String) As Boolean
'For make startup and make registry setting:  Makes the registry setting
On Error GoTo error
Dim phkResult As Long
Dim lResult As Long
Dim SA As SECURITY_ATTRIBUTES
Dim lCreate As Long
RegCreateKeyEx hKey, lpszSubKey, 0, "", REG_OPTION_NON_VOLATILE, _
KEY_ALL_ACCESS, SA, phkResult, lCreate
lResult = RegSetValueEx(phkResult, sSetValue, 0, 1, sValue, _
CLng(Len(sValue) + 1))
RegCloseKey phkResult
MakeRegFile = (lResult = ERROR_SUCCESS)
Exit Function
error:
MakeRegFile = False
End Function

Public Sub ExecuteNewProgram()
'This will execute the program over again, creating two working copies
On Error GoTo error
ret = Shell("rundll32.exe url.dll,FileProtocolHandler " & App.Path & "\" & App.EXEName & ".EXE")
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

'Section 9:  Form Stuff

Public Sub Ontop(FormName As Form)
'Make a form always ontop of other windows
On Error GoTo error
Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub NotOntop(FormName As Form)
'Make a form not always ontop of other windows
On Error GoTo error
Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub InvisibleForm(Frm As Form)
'Make a form invisible
On Error GoTo error
Dim rctClient As RECT, rctFrame As RECT
Dim hClient As Long, hFrame As Long
GetWindowRect Frm.hWnd, rctFrame
GetClientRect Frm.hWnd, rctClient
Dim lpTL As POINTAPI, lpBR As POINTAPI
lpTL.X = rctFrame.Left
lpTL.Y = rctFrame.Top
lpBR.X = rctFrame.Right
lpBR.Y = rctFrame.Bottom
ScreenToClient Frm.hWnd, lpTL
ScreenToClient Frm.hWnd, lpBR
rctFrame.Left = lpTL.X
rctFrame.Top = lpTL.Y
rctFrame.Right = lpBR.X
rctFrame.Bottom = lpBR.Y
rctClient.Left = Abs(rctFrame.Left)
rctClient.Top = Abs(rctFrame.Top)
rctClient.Right = rctClient.Right + Abs(rctFrame.Left)
rctClient.Bottom = rctClient.Bottom + Abs(rctFrame.Top)
rctFrame.Right = rctFrame.Right + Abs(rctFrame.Left)
rctFrame.Bottom = rctFrame.Bottom + Abs(rctFrame.Top)
rctFrame.Top = 0
rctFrame.Left = 0
hClient = CreateRectRgn(rctClient.Left, rctClient.Top, rctClient.Right, rctClient.Bottom)
hFrame = CreateRectRgn(rctFrame.Left, rctFrame.Top, rctFrame.Right, rctFrame.Bottom)
CombineRgn hFrame, hClient, hFrame, RGN_XOR
SetWindowRgn Frm.hWnd, hFrame, True
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub HoleInForm(Rectangular As Boolean, HoleWidth As Single, HoleHeight As Single, HoleLeft As Single, HoleTop As Single, Frm As Form)
'This will put a hole in the form (you can see through the form with that hole)
On Error GoTo error
Const RGN_DIFF = 4
Dim outer_rgn As Long
Dim inner_rgn As Long
Dim combined_rgn As Long
Dim wid As Single
Dim hgt As Single
Dim border_width As Single
Dim title_height As Single
If Frm.WindowState = vbMinimized Then Exit Sub
wid = ScaleX(Frm.width, vbTwips, vbPixels)
hgt = ScaleY(Frm.height, vbTwips, vbPixels)
outer_rgn = CreateRectRgn(0, 0, wid, hgt)
border_width = (wid - ScaleWidth) / 2
title_height = hgt - border_width - ScaleHeight
If Rectangular = True Then
inner_rgn = CreateRectRgn(HoleLeft, HoleTop, HoleWidth, HoleHeight)
Else
inner_rgn = CreateEllipticRgn(HoleLeft, HoleTop, HoleWidth, HoleHeight)
End If
combined_rgn = CreateRectRgn(0, 0, 0, 0)
CombineRgn combined_rgn, outer_rgn, inner_rgn, RGN_DIFF
SetWindowRgn Frm.hWnd, combined_rgn, True
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

'Section 10:  Clipboard Stuff

Public Sub ClipboardCopy(Text As String)
'Copies text to the clipboard
On Error GoTo error
Clipboard.Clear
Clipboard.SetText Text$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Function ClipboardGet()
'Gets the copied text from the clipboard
On Error GoTo error
ClipboardGet = Clipboard.GetText
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Sub ClearClipboard()
'Clears the clipboard
On Error GoTo error
Clipboard.Clear
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

'Section 11:  Ping

Public Sub Ping(Message As String, IPAddress As String)
'Ping an IP Address
On Error GoTo error
    Dim hFile       As Long
    Dim lRet        As Long
    Dim lIPAddress  As Long
    Dim strMessage  As String
    Dim pOptions    As ip_option_information
    Dim pReturn     As icmp_echo_reply
    Dim iVal        As Integer
    Dim lPingRet    As Long
    Dim pWsaData    As tagWSAData
    strMessage = Message$
    iVal = WSAStartup(&H101, pWsaData)
    lIPAddress = ConvertIPAddressToLong(IPAddress$)
    hFile = IcmpCreateFile()
    pOptions.Ttl = 30
    pOptions.Tos = 12
    pWsaData.wVersion = 4
    lRet = IcmpSendEcho(hFile, _
                        lIPAddress, _
                        strMessage, _
                        Len(strMessage), _
                        pOptions, _
                        pReturn, _
                        Len(pReturn), _
                        PING_TIMEOUT)

    If lRet = 0 Then
    Else
        If pReturn.Status <> 0 Then
        Else
            lRet = IcmpCloseHandle(hFile)
            iVal = WSACleanup()
            Exit Sub
        End If
    End If
lRet = IcmpCloseHandle(hFile)
iVal = WSACleanup()
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Function ConvertIPAddressToLong(strAddress As String) As Long
'For Ping:  It changes the IP Address so it can be used to send the ping
On Error GoTo error
    Dim strTemp             As String
    Dim lAddress            As Long
    Dim iValCount           As Integer
    Dim lDotValues(1 To 4)  As String
    strTemp = strAddress
    iValCount = 0
    While InStr(strTemp, ".") > 0
        iValCount = iValCount + 1
        lDotValues(iValCount) = Mid(strTemp, 1, InStr(strTemp, ".") - 1)
        strTemp = Mid(strTemp, InStr(strTemp, ".") + 1)
        Wend
    iValCount = iValCount + 1
    lDotValues(iValCount) = strTemp
    If iValCount <> 4 Then
        ConvertIPAddressToLong = 0
        Exit Function
        End If
    lAddress = Val("&H" & Right("00" & Hex(lDotValues(4)), 2) & _
                Right("00" & Hex(lDotValues(3)), 2) & _
                Right("00" & Hex(lDotValues(2)), 2) & _
                Right("00" & Hex(lDotValues(1)), 2))
    ConvertIPAddressToLong = lAddress
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

'Section 12:  Code/Decode

Function Code1(Text As String)
'This codes text into different words and phrases!  Make like a secret agent..
On Error GoTo error
Dim i As Long
Dim RndN As Integer
Dim endstr As String
Randomize Timer
Text$ = ReplaceC(Text$, "A", "a")
Text$ = ReplaceC(Text$, "B", "b")
Text$ = ReplaceC(Text$, "C", "c")
Text$ = ReplaceC(Text$, "D", "d")
Text$ = ReplaceC(Text$, "E", "e")
Text$ = ReplaceC(Text$, "F", "f")
Text$ = ReplaceC(Text$, "G", "g")
Text$ = ReplaceC(Text$, "H", "h")
Text$ = ReplaceC(Text$, "I", "i")
Text$ = ReplaceC(Text$, "J", "j")
Text$ = ReplaceC(Text$, "K", "k")
Text$ = ReplaceC(Text$, "L", "l")
Text$ = ReplaceC(Text$, "M", "m")
Text$ = ReplaceC(Text$, "N", "n")
Text$ = ReplaceC(Text$, "O", "o")
Text$ = ReplaceC(Text$, "P", "p")
Text$ = ReplaceC(Text$, "Q", "q")
Text$ = ReplaceC(Text$, "R", "r")
Text$ = ReplaceC(Text$, "S", "s")
Text$ = ReplaceC(Text$, "T", "t")
Text$ = ReplaceC(Text$, "U", "u")
Text$ = ReplaceC(Text$, "V", "v")
Text$ = ReplaceC(Text$, "W", "w")
Text$ = ReplaceC(Text$, "X", "x")
Text$ = ReplaceC(Text$, "Y", "y")
Text$ = ReplaceC(Text$, "Z", "z")
Text$ = ReplaceC(Text$, "  ", ";")
Text$ = ReplaceC(Text$, " ", ",")
For i = 1 To Len(Text$)
RndN = Int((3 - 0 + 1) * Rnd + 0)
If Mid$(Text$, i, 1) = "a" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " somewhere"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " did you"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " flowers"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " eat food"
  End If
ElseIf Mid$(Text$, i, 1) = "b" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " light candle"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " mirror"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " cold soup"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " video tape"
  End If
ElseIf Mid$(Text$, i, 1) = "c" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " the murder"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " read book"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " the show"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " paper"
  End If
ElseIf Mid$(Text$, i, 1) = "d" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " beautiful"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " do not"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " bring"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " that"
  End If
ElseIf Mid$(Text$, i, 1) = "e" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " star"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " itself"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " in a"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " by"
  End If
ElseIf Mid$(Text$, i, 1) = "f" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " it is"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " sea"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " myself"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " powerful"
  End If
ElseIf Mid$(Text$, i, 1) = "g" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " aren't"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " nail filer"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " everlasting"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " magic"
  End If
ElseIf Mid$(Text$, i, 1) = "h" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " tomorrow"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " tree"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " it will"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " fat"
  End If
ElseIf Mid$(Text$, i, 1) = "i" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " isn't"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " explosion"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " at school"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " apples"
  End If
ElseIf Mid$(Text$, i, 1) = "j" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " when"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " onions"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " night"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " about it"
  End If
ElseIf Mid$(Text$, i, 1) = "k" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " days"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " right"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " please"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " oranges"
  End If
ElseIf Mid$(Text$, i, 1) = "l" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " wrong"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " yesterday"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " has"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " money"
  End If
ElseIf Mid$(Text$, i, 1) = "m" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " today"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " dad"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " mother"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " his"
  End If
ElseIf Mid$(Text$, i, 1) = "n" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " french"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " hurt"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " ham"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " milk"
  End If
ElseIf Mid$(Text$, i, 1) = "o" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " not"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " see you"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " rot"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " five"
  End If
ElseIf Mid$(Text$, i, 1) = "p" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " see me"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " hard"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " mask"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " ants"
  End If
ElseIf Mid$(Text$, i, 1) = "q" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " yes"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " soft"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " four"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " in flour"
  End If
ElseIf Mid$(Text$, i, 1) = "r" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " no"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " fast"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " three"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " cat"
  End If
ElseIf Mid$(Text$, i, 1) = "s" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " slow"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " super"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " two"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " over the"
  End If
ElseIf Mid$(Text$, i, 1) = "t" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " medium"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " hit"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " one"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " rainbow"
  End If
ElseIf Mid$(Text$, i, 1) = "u" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " zero"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " fire"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " ice"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " malt"
  End If
ElseIf Mid$(Text$, i, 1) = "v" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " six"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " hair"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " light switch"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " metal"
  End If
ElseIf Mid$(Text$, i, 1) = "w" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " computer"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " comb"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " bomb"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " writing"
  End If
ElseIf Mid$(Text$, i, 1) = "x" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " eight ball"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " smear"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " letter"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " cups"
  End If
ElseIf Mid$(Text$, i, 1) = "y" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " nine"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " table"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " basket"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " open door"
  End If
ElseIf Mid$(Text$, i, 1) = "z" Then
  If RndN = 0 Then
  endstr$ = endstr$ + " ten"
  ElseIf RndN = 1 Then
  endstr$ = endstr$ + " to car"
  ElseIf RndN = 2 Then
  endstr$ = endstr$ + " hallway"
  ElseIf RndN = 3 Then
  endstr$ = endstr$ + " in house"
  End If
Else
endstr$ = endstr$ + Mid$(Text$, i, 1)
End If
Next i
endstr$ = Mid$(endstr$, 2, Len(endstr$) - 1)
Code1 = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Code2(Text As String)
'This is a simpler (and smaller) coding system than code 1
On Error GoTo error
Text$ = ReplaceC(Text$, "  ", ";")
Text$ = ReplaceC(Text$, " ", ",")
Text$ = ReplaceC(Text$, "A", "a")
Text$ = ReplaceC(Text$, "B", "b")
Text$ = ReplaceC(Text$, "C", "c")
Text$ = ReplaceC(Text$, "D", "d")
Text$ = ReplaceC(Text$, "E", "e")
Text$ = ReplaceC(Text$, "F", "f")
Text$ = ReplaceC(Text$, "G", "g")
Text$ = ReplaceC(Text$, "H", "h")
Text$ = ReplaceC(Text$, "I", "i")
Text$ = ReplaceC(Text$, "J", "j")
Text$ = ReplaceC(Text$, "K", "k")
Text$ = ReplaceC(Text$, "L", "l")
Text$ = ReplaceC(Text$, "M", "m")
Text$ = ReplaceC(Text$, "N", "n")
Text$ = ReplaceC(Text$, "O", "o")
Text$ = ReplaceC(Text$, "P", "p")
Text$ = ReplaceC(Text$, "Q", "q")
Text$ = ReplaceC(Text$, "R", "r")
Text$ = ReplaceC(Text$, "S", "s")
Text$ = ReplaceC(Text$, "T", "t")
Text$ = ReplaceC(Text$, "U", "u")
Text$ = ReplaceC(Text$, "V", "v")
Text$ = ReplaceC(Text$, "W", "w")
Text$ = ReplaceC(Text$, "X", "x")
Text$ = ReplaceC(Text$, "Y", "y")
Text$ = ReplaceC(Text$, "Z", "z")
Text$ = ReplaceC(Text$, "a", " IT")
Text$ = ReplaceC(Text$, "b", " AE")
Text$ = ReplaceC(Text$, "c", " TA")
Text$ = ReplaceC(Text$, "d", " EA")
Text$ = ReplaceC(Text$, "e", " NA")
Text$ = ReplaceC(Text$, "f", " NT")
Text$ = ReplaceC(Text$, "g", " IE")
Text$ = ReplaceC(Text$, "h", " NN")
Text$ = ReplaceC(Text$, "i", " TE")
Text$ = ReplaceC(Text$, "j", " EI")
Text$ = ReplaceC(Text$, "k", " TI")
Text$ = ReplaceC(Text$, "l", " II")
Text$ = ReplaceC(Text$, "m", " NE")
Text$ = ReplaceC(Text$, "n", " AI")
Text$ = ReplaceC(Text$, "o", " TN")
Text$ = ReplaceC(Text$, "p", " AA")
Text$ = ReplaceC(Text$, "q", " EN")
Text$ = ReplaceC(Text$, "r", " IN")
Text$ = ReplaceC(Text$, "s", " AT")
Text$ = ReplaceC(Text$, "t", " AN")
Text$ = ReplaceC(Text$, "u", " NI")
Text$ = ReplaceC(Text$, "v", " EE")
Text$ = ReplaceC(Text$, "w", " TT")
Text$ = ReplaceC(Text$, "x", " XX")
Text$ = ReplaceC(Text$, "y", " ET")
Text$ = ReplaceC(Text$, "z", " IA")
Text$ = Mid$(Text$, 2, Len(Text$) - 1)
Code2 = Text$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Decode1(Text As String)
'This decodes text coded by code 1
On Error GoTo error
Text$ = " " & Text$
Text$ = ReplaceC(Text$, " somewhere", "a")
Text$ = ReplaceC(Text$, " did you", "a")
Text$ = ReplaceC(Text$, " flowers", "a")
Text$ = ReplaceC(Text$, " eat food", "a")
Text$ = ReplaceC(Text$, " light candle", "b")
Text$ = ReplaceC(Text$, " mirror", "b")
Text$ = ReplaceC(Text$, " cold soup", "b")
Text$ = ReplaceC(Text$, " video tape", "b")
Text$ = ReplaceC(Text$, " the murder", "c")
Text$ = ReplaceC(Text$, " read book", "c")
Text$ = ReplaceC(Text$, " the show", "c")
Text$ = ReplaceC(Text$, " paper", "c")
Text$ = ReplaceC(Text$, " beautiful", "d")
Text$ = ReplaceC(Text$, " do not", "d")
Text$ = ReplaceC(Text$, " bring", "d")
Text$ = ReplaceC(Text$, " that", "d")
Text$ = ReplaceC(Text$, " star", "e")
Text$ = ReplaceC(Text$, " itself", "e")
Text$ = ReplaceC(Text$, " in a", "e")
Text$ = ReplaceC(Text$, " by", "e")
Text$ = ReplaceC(Text$, " it is", "f")
Text$ = ReplaceC(Text$, " sea", "f")
Text$ = ReplaceC(Text$, " myself", "f")
Text$ = ReplaceC(Text$, " powerful", "f")
Text$ = ReplaceC(Text$, " aren't", "g")
Text$ = ReplaceC(Text$, " nail filer", "g")
Text$ = ReplaceC(Text$, " everlasting", "g")
Text$ = ReplaceC(Text$, " magic", "g")
Text$ = ReplaceC(Text$, " tomorrow", "h")
Text$ = ReplaceC(Text$, " tree", "h")
Text$ = ReplaceC(Text$, " it will", "h")
Text$ = ReplaceC(Text$, " fat", "h")
Text$ = ReplaceC(Text$, " isn't", "i")
Text$ = ReplaceC(Text$, " explosion", "i")
Text$ = ReplaceC(Text$, " at school", "i")
Text$ = ReplaceC(Text$, " apples", "i")
Text$ = ReplaceC(Text$, " when", "j")
Text$ = ReplaceC(Text$, " onions", "j")
Text$ = ReplaceC(Text$, " night", "j")
Text$ = ReplaceC(Text$, " about it", "j")
Text$ = ReplaceC(Text$, " days", "k")
Text$ = ReplaceC(Text$, " right", "k")
Text$ = ReplaceC(Text$, " please", "k")
Text$ = ReplaceC(Text$, " oranges", "k")
Text$ = ReplaceC(Text$, " wrong", "l")
Text$ = ReplaceC(Text$, " yesterday", "l")
Text$ = ReplaceC(Text$, " has", "l")
Text$ = ReplaceC(Text$, " money", "l")
Text$ = ReplaceC(Text$, " today", "m")
Text$ = ReplaceC(Text$, " had", "m")
Text$ = ReplaceC(Text$, " mother", "m")
Text$ = ReplaceC(Text$, " his", "m")
Text$ = ReplaceC(Text$, " french", "n")
Text$ = ReplaceC(Text$, " hurt", "n")
Text$ = ReplaceC(Text$, " ham", "n")
Text$ = ReplaceC(Text$, " milk", "n")
Text$ = ReplaceC(Text$, " not", "o")
Text$ = ReplaceC(Text$, " see you", "o")
Text$ = ReplaceC(Text$, " rot", "o")
Text$ = ReplaceC(Text$, " five", "o")
Text$ = ReplaceC(Text$, " see me", "p")
Text$ = ReplaceC(Text$, " hard", "p")
Text$ = ReplaceC(Text$, " mask", "p")
Text$ = ReplaceC(Text$, " ants", "p")
Text$ = ReplaceC(Text$, " yes", "q")
Text$ = ReplaceC(Text$, " soft", "q")
Text$ = ReplaceC(Text$, " four", "q")
Text$ = ReplaceC(Text$, " in flour", "q")
Text$ = ReplaceC(Text$, " no", "r")
Text$ = ReplaceC(Text$, " fast", "r")
Text$ = ReplaceC(Text$, " three", "r")
Text$ = ReplaceC(Text$, " cat", "r")
Text$ = ReplaceC(Text$, " slow", "s")
Text$ = ReplaceC(Text$, " super", "s")
Text$ = ReplaceC(Text$, " two", "s")
Text$ = ReplaceC(Text$, " over the", "s")
Text$ = ReplaceC(Text$, " medium", "t")
Text$ = ReplaceC(Text$, " hit", "t")
Text$ = ReplaceC(Text$, " one", "t")
Text$ = ReplaceC(Text$, " rainbow", "t")
Text$ = ReplaceC(Text$, " zero", "u")
Text$ = ReplaceC(Text$, " fire", "u")
Text$ = ReplaceC(Text$, " ice", "u")
Text$ = ReplaceC(Text$, " malt", "u")
Text$ = ReplaceC(Text$, " six", "v")
Text$ = ReplaceC(Text$, " hair", "v")
Text$ = ReplaceC(Text$, " light switch", "v")
Text$ = ReplaceC(Text$, " metal", "v")
Text$ = ReplaceC(Text$, " computer", "w")
Text$ = ReplaceC(Text$, " comb", "w")
Text$ = ReplaceC(Text$, " bomb", "w")
Text$ = ReplaceC(Text$, " writing", "w")
Text$ = ReplaceC(Text$, " eight ball", "x")
Text$ = ReplaceC(Text$, " smear", "x")
Text$ = ReplaceC(Text$, " letter", "x")
Text$ = ReplaceC(Text$, " cups", "x")
Text$ = ReplaceC(Text$, " nine", "y")
Text$ = ReplaceC(Text$, " table", "y")
Text$ = ReplaceC(Text$, " basket", "y")
Text$ = ReplaceC(Text$, " open door", "y")
Text$ = ReplaceC(Text$, " ten", "z")
Text$ = ReplaceC(Text$, " to car", "z")
Text$ = ReplaceC(Text$, " hallway", "z")
Text$ = ReplaceC(Text$, " in house", "z")
Text$ = ReplaceC(Text$, ";", "  ")
Text$ = ReplaceC(Text$, ",", " ")
Decode1 = Text$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Decode2(Text As String)
'This decodes text coded by code 2
On Error GoTo error
Text$ = " " & Text$
Text$ = ReplaceC(Text$, " IT", "a")
Text$ = ReplaceC(Text$, " AE", "b")
Text$ = ReplaceC(Text$, " TA", "c")
Text$ = ReplaceC(Text$, " EA", "d")
Text$ = ReplaceC(Text$, " NA", "e")
Text$ = ReplaceC(Text$, " NT", "f")
Text$ = ReplaceC(Text$, " IE", "g")
Text$ = ReplaceC(Text$, " NN", "h")
Text$ = ReplaceC(Text$, " TE", "i")
Text$ = ReplaceC(Text$, " EI", "j")
Text$ = ReplaceC(Text$, " TI", "k")
Text$ = ReplaceC(Text$, " II", "l")
Text$ = ReplaceC(Text$, " NE", "m")
Text$ = ReplaceC(Text$, " AI", "n")
Text$ = ReplaceC(Text$, " TN", "o")
Text$ = ReplaceC(Text$, " AA", "p")
Text$ = ReplaceC(Text$, " EN", "q")
Text$ = ReplaceC(Text$, " IN", "r")
Text$ = ReplaceC(Text$, " AT", "s")
Text$ = ReplaceC(Text$, " AN", "t")
Text$ = ReplaceC(Text$, " NI", "u")
Text$ = ReplaceC(Text$, " EE", "v")
Text$ = ReplaceC(Text$, " TT", "w")
Text$ = ReplaceC(Text$, " XX", "x")
Text$ = ReplaceC(Text$, " ET", "y")
Text$ = ReplaceC(Text$, " IA", "z")
Text$ = ReplaceC(Text$, ";", "  ")
Text$ = ReplaceC(Text$, ",", " ")
Decode2 = Text$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Private Function ReplaceC(MainStr As String, OldStr As String, NewStr As String) As String
'For Section 12 (Code/Decode):  Replaces one string with another
On Error GoTo error
ReplaceC = ""
Dim NewStrString As String
Dim i As Integer
For i = 1 To Len(MainStr)
  If Mid(MainStr, i, Len(OldStr)) = OldStr Then
  NewStrString = NewStrString & NewStr
  i = i + Len(OldStr) - 1
  Else
  NewStrString = NewStrString & Mid(MainStr, i, 1)
  End If
Next i
ReplaceC = NewStrString
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

'Section 13:  Math

Function Add(num1 As Long, num2 As Long) As Long
'Add two numbers
On Error GoTo error
Add = Val(num1) + Val(num2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Subtract(num1 As Long, num2 As Long) As Long
'Subtract two numbers
On Error GoTo error
Subtract = Val(num1) - Val(num2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Divide(num1 As Long, num2 As Long) As Long
'Divide two numbers
On Error GoTo error
Divide = Val(num1) / Val(num2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Multiply(num1 As Long, num2 As Long) As Long
'Multiply two numbers
On Error GoTo error
Multiply = Val(num1) * Val(num2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function ToPower(num1 As Long, num2 As Long) As Long
'Bring num1 to the power (exponent) of num2
On Error GoTo error
ToPower = Val(num1) ^ Val(num2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function ToRoot(num1 As Long, num2 As Long) As Long
'Bring num1 to the root of num2
On Error GoTo error
ToRoot = Val(num1) ^ (1 / Val(num2))
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function FractionToDecimal(numerator As Integer, denominator As Integer) As Long
'Turns a fraction into a decimal
On Error GoTo error
FractionToDecimal = numerator / denominator
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function DecimalToPercentage(DecimalNum As Long) As String
'Turns a decimal into a percentage
On Error GoTo error
DecimalToPercentage = (DecimalNum * 100) & "%"
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function PercentageToDeciaml(PercentNum As String) As Long
'Turns a percentage into a decimal
On Error GoTo error
If Mid$(PercentNum$, Len(PercentNum$), 1) = "%" Then
PercentNum$ = Mid$(PercentNum$, 2, Len(PercentNum$) - 1)
End If
PercentageToDecimal = Val(PercentNum$) / 100
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function AreaOfCircle(radius As Long)
'Gets the area of a circle
On Error GoTo error
pi = 3.141592654
AreaOfCircle = pi * (radius ^ 2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Circumference(radius As Long)
'Gets the circumference of a circle
On Error GoTo error
pi = 3.141592654
Circumference = pi * 2 * radius
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function AreaOfSquare(side As Long)
'Gets the area of a square
On Error GoTo error
AreaOfSquare = side ^ 2
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function PerimeterOfSquare(side As Long)
'Gets the perimeter of a square
On Error GoTo error
PerimeterOfSquare = 4 * side
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function PerimeterOfRectangle(Length As Long, width As Long)
'Gets the perimeter of a rectangle
On Error GoTo error
PerimeterOfRectangle = (2 * Length) + (2 * width)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function AreaOfRectangle(Length As Long, width As Long)
'Gets the area of a rectangle
On Error GoTo error
AreaOfRectangle = Length * width
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function AreaOfTriangle(base As Long, height As Long)
'Gets the area of a triangle
On Error GoTo error
AreaOfTriangle = (1 / 2) * base * height
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function PerimeterOfTriangle(side1 As Long, side2 As Long, side3 As Long)
'Gets the perimeter of a triangle
On Error GoTo error
PerimeterOfTriangle = side1 + side2 + side3
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function PerimeterOf4SidedPolygon(side1 As Long, side2 As Long, side3 As Long, side4 As Long)
'Gets the perimeter of any 4 sided polygon
On Error GoTo error
PerimeterOf4SidedPolygon = side1 + side2 + side3 + side4
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfCube(edge As Long)
'Gets the volume of a cube
On Error GoTo error
VolumeOfCube = edge ^ 3
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfPrism(base As Long, height As Long)
'Gets the volume of a prism
On Error GoTo error
VolumeOfPrism = base * height
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfSphere(radius As Long)
'Gets the volume of a sphere
On Error GoTo error
pi = 3.141592654
VolumeOfSphere = (4 / 3) * pi * (radius ^ 3)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfPyramid(base As Long, height As Long)
'Gets the volume of a pyramid
On Error GoTo error
VolumeOfPyramid = (1 / 3) * base * height
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfCone(radius As Long, height As Long)
'Gets the volume of a cone
On Error GoTo error
pi = 3.141592654
VolumeOfCone = (1 / 3) * pi * (radius ^ 2) * height
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function VolumeOfCylinder(radius As Long, height As Long)
'Gets the volume of a cylinder
On Error GoTo error
pi = 3.141592654
VolumeOfCylinder = pi * height * (radius ^ 2)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

'Section 14:  Color Fading

Function FadeThreeColorHTML(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$)
'This will fade three colors in HTML color coding
On Error GoTo error
textlen% = Len(TheText)
fstlen% = (Int(textlen%) / 2)
part1$ = Left(TheText, fstlen%)
part2$ = Right(TheText, textlen% - fstlen%)
textlen% = Len(part1$)
For i = 1 To textlen%
TextDone$ = Left(part1$, i)
LastChr$ = Right(TextDone$, 1)
ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
colorx2 = RGBToHex(ColorX)
Faded1$ = Faded1$ + "<FONT COLOR=" & colorx2 & ">" + LastChr$ + "</FONT>"
Next i
textlen% = Len(part2$)
For i = 1 To textlen%
TextDone$ = Left(part2$, i)
LastChr$ = Right(TextDone$, 1)
ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
colorx2 = RGBToHex(ColorX)
Faded2$ = Faded2$ + "<FONT COLOR=" & colorx2 & ">" + LastChr$ + "</FONT>"
Next i
FadeThreeColorHTML = Faded1$ + Faded2$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Private Function FadeTwoColorHTML(R1%, G1%, B1%, R2%, G2%, B2%, TheText$)
'This will fade two colors in HTML color coding
On Error GoTo error
textlen$ = Len(TheText)
For i = 1 To textlen$
TextDone$ = Left(TheText, i)
LastChr$ = Right(TextDone$, 1)
ColorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
colorx2 = RGBToHex(ColorX)
Faded$ = Faded$ + "<FONT COLOR=" & colorx2 & ">" + LastChr$ + "</FONT>"
Next i
FadeTwoColorHTML = Faded$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function FadeThreeColorYahoo(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$)
'This will fade three colors in Yahoo color coding
On Error GoTo error
textlen% = Len(TheText)
fstlen% = (Int(textlen%) / 2)
part1$ = Left(TheText, fstlen%)
part2$ = Right(TheText, textlen% - fstlen%)
textlen% = Len(part1$)
For i = 1 To textlen%
TextDone$ = Left(part1$, i)
LastChr$ = Right(TextDone$, 1)
ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
colorx2 = RGBToHex(ColorX)
Faded1$ = Faded1$ + "<#" & colorx2 & ">" + LastChr$
Next i
textlen% = Len(part2$)
For i = 1 To textlen%
TextDone$ = Left(part2$, i)
LastChr$ = Right(TextDone$, 1)
ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
colorx2 = RGBToHex(ColorX)
Faded2$ = Faded2$ + "<#" & colorx2 & ">" + LastChr$
Next i
FadeThreeColorYahoo = Faded1$ + Faded2$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Private Function FadeTwoColorYahoo(R1%, G1%, B1%, R2%, G2%, B2%, TheText$)
'This will fade two colors in Yahoo color coding
On Error GoTo error
textlen$ = Len(TheText)
For i = 1 To textlen$
TextDone$ = Left(TheText, i)
LastChr$ = Right(TextDone$, 1)
ColorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
colorx2 = RGBToHex(ColorX)
Faded$ = Faded$ + "<#" & colorx2 & ">" + LastChr$
Next i
FadeTwoColorYahoo = Faded$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function FadeThreeColorANSI(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$)
'This will fade three colors in ANSI color coding
On Error GoTo error
textlen% = Len(TheText)
fstlen% = (Int(textlen%) / 2)
part1$ = Left(TheText, fstlen%)
part2$ = Right(TheText, textlen% - fstlen%)
textlen% = Len(part1$)
For i = 1 To textlen%
TextDone$ = Left(part1$, i)
LastChr$ = Right(TextDone$, 1)
ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
colorx2 = RGBToHex(ColorX)
Faded1$ = Faded1$ + "[#" & colorx2 & LastChr$
Next i
textlen% = Len(part2$)
For i = 1 To textlen%
TextDone$ = Left(part2$, i)
LastChr$ = Right(TextDone$, 1)
ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
colorx2 = RGBToHex(ColorX)
Faded2$ = Faded2$ + "[#" & colorx2 & LastChr$
Next i
FadeThreeColorANSI = Faded1$ + Faded2$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Private Function FadeTwoColorANSI(R1%, G1%, B1%, R2%, G2%, B2%, TheText$)
'This will fade two colors in ANSI color coding
On Error GoTo error
textlen$ = Len(TheText)
For i = 1 To textlen$
TextDone$ = Left(TheText, i)
LastChr$ = Right(TextDone$, 1)
ColorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
colorx2 = RGBToHex(ColorX)
Faded$ = Faded$ + "[#" & colorx2 & LastChr$
Next i
FadeTwoColorANSI = Faded$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

'Section 15:  Exit/Restart/Reboot Computer

Function RestartWindows()
'This will restart windows
On Error GoTo error
Dim RetVal As Integer
RetVal = ExitWindows(EW_RESTARTWINDOWS, 0)
RestartWindows = RetVal
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function DoExitWindows()
'This will exit windows
On Error GoTo error
Dim RetVal As Integer
RetVal = ExitWindows(EW_EXITWINDOWS, 0)
ExitWindows = RetVal
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function RebootComputer()
'This will reboot the computer
On Error GoTo error
Dim RetVal As Integer
RetVal = ExitWindows(EW_REBOOTSYSTEM, 0)
RebootComputer = RetVal
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

'Section 16:  Text$ "Changers"

Function AltCaps(Text As String)
'This will make the caps in text go on and off for each letter, like this:  cOoL
On Error GoTo error
Dim i As Integer
Dim s As String
s = ""
For i = 1 To Len(Text$)
  keyval = Asc(Mid$(Text$, i, 1))
  If (keyval >= 96 And keyval < 96 + 26) Or (keyval >= 64 And keyval < 64 + 26) Then
    If (i And 1) = 1 Then
      If keyval < 96 Then
        s = s + Chr$(96 + keyval - 64)
      Else
        s = s + Chr$(keyval)
      End If
    Else
      If keyval >= 96 Then
        s = s + Chr$(64 + keyval - 96)
      Else
        s = s + Chr$(keyval)
      End If
    End If
  Else
    s = s + Chr$(keyval)
  End If
Next i
Text$ = s
AltCaps = Text$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function BackwardsText(Text As String)
'This will make text go backwards, like this:  looC (Cool)
On Error GoTo error
For i% = 1 To Len(Text$)
stringy$ = Mid$(Text$, i%, 1)
final$ = stringy$ + final$
Next i%
BackwardsText = final$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function EliteType(Text As String)
'This will change characters to make them "elite", example:  ©00|_
On Error GoTo error
s(0) = "æ"
s(1) = "å"
s(2) = "b"
s(3) = "<"
s(4) = "c|"
s(5) = "ë"
s(6) = "f"
s(7) = "9"
s(8) = "h"
s(9) = "ï"
s(10) = "j"
s(11) = "|<"
s(12) = "|_"
s(13) = "/x\"
s(14) = "|\|"
s(15) = "0"
s(16) = "p"
s(17) = "q"
s(18) = "r"
s(19) = "_/¯"
s(20) = "-|-"
s(21) = "µ"
s(22) = "\/"
s(23) = "\/\/"
s(24) = "×"
s(25) = "ÿ"
s(26) = "¯/_"
s(27) = "Ä"
s(28) = "ß"
s(29) = "©"
s(30) = "|}"
s(31) = "È"
s(32) = "F"
s(33) = "G"
s(34) = "|-|"
s(35) = "I"
s(36) = "J"
s(37) = "]<"
s(38) = "]_"
s(39) = "/\/\"
s(40) = "|\|"
s(41) = "{}"
s(42) = "P"
s(43) = "¶"
s(44) = "|2"
s(45) = "§"
s(46) = "¯|¯"
s(47) = "|_|"
s(48) = "\/"
s(49) = "\x/"
s(50) = "><"
s(51) = "¥"
s(52) = "¯/_"
Text$ = ReplaceC(Text$, "a", s(1))
Text$ = ReplaceC(Text$, "b", s(2))
Text$ = ReplaceC(Text$, "c", s(3))
Text$ = ReplaceC(Text$, "d", s(4))
Text$ = ReplaceC(Text$, "e", s(5))
Text$ = ReplaceC(Text$, "f", s(6))
Text$ = ReplaceC(Text$, "g", s(7))
Text$ = ReplaceC(Text$, "h", s(8))
Text$ = ReplaceC(Text$, "i", s(9))
Text$ = ReplaceC(Text$, "j", s(10))
Text$ = ReplaceC(Text$, "k", s(11))
Text$ = ReplaceC(Text$, "l", s(12))
Text$ = ReplaceC(Text$, "m", s(13))
Text$ = ReplaceC(Text$, "n", s(14))
Text$ = ReplaceC(Text$, "o", s(15))
Text$ = ReplaceC(Text$, "p", s(16))
Text$ = ReplaceC(Text$, "q", s(17))
Text$ = ReplaceC(Text$, "r", s(18))
Text$ = ReplaceC(Text$, "s", s(19))
Text$ = ReplaceC(Text$, "t", s(20))
Text$ = ReplaceC(Text$, "u", s(21))
Text$ = ReplaceC(Text$, "v", s(22))
Text$ = ReplaceC(Text$, "w", s(23))
Text$ = ReplaceC(Text$, "x", s(24))
Text$ = ReplaceC(Text$, "y", s(25))
Text$ = ReplaceC(Text$, "z", s(26))
Text$ = ReplaceC(Text$, "A", s(27))
Text$ = ReplaceC(Text$, "B", s(28))
Text$ = ReplaceC(Text$, "C", s(29))
Text$ = ReplaceC(Text$, "D", s(30))
Text$ = ReplaceC(Text$, "E", s(31))
Text$ = ReplaceC(Text$, "F", s(32))
Text$ = ReplaceC(Text$, "G", s(33))
Text$ = ReplaceC(Text$, "H", s(34))
Text$ = ReplaceC(Text$, "I", s(35))
Text$ = ReplaceC(Text$, "J", s(36))
Text$ = ReplaceC(Text$, "K", s(37))
Text$ = ReplaceC(Text$, "L", s(38))
Text$ = ReplaceC(Text$, "M", s(39))
Text$ = ReplaceC(Text$, "N", s(40))
Text$ = ReplaceC(Text$, "O", s(41))
Text$ = ReplaceC(Text$, "P", s(42))
Text$ = ReplaceC(Text$, "Q", s(43))
Text$ = ReplaceC(Text$, "R", s(44))
Text$ = ReplaceC(Text$, "S", s(45))
Text$ = ReplaceC(Text$, "T", s(46))
Text$ = ReplaceC(Text$, "U", s(47))
Text$ = ReplaceC(Text$, "V", s(48))
Text$ = ReplaceC(Text$, "W", s(49))
Text$ = ReplaceC(Text$, "X", s(50))
Text$ = ReplaceC(Text$, "Y", s(51))
Text$ = ReplaceC(Text$, "Z", s(52))
Text$ = ReplaceC(Text$, "ae", s(0))
EliteType = Text$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function SpaceCharacters(Text As String, AmountOfSpaces As Integer)
'This will put a space between every character in the text, like this:  C o o l
On Error GoTo error
Dim i As Long
Dim SpaceStr As String
If AmountOfSpaces > 100 Then
AmountOfSpaces = 100
ElseIf AmountOfSpaces < 1 Then
AmountOfSpaces = 1
End If
For i = 1 To AmountOfSpaces
SpaceStr$ = SpaceStr$ + " "
Next i
Dim endstr As String
For i = 1 To Len(Text$)
endstr$ = endstr$ & Mid$(Text$, i, 1) & SpaceStr$
Next i
endstr$ = Mid$(endstr$, 1, Len(endstr$) - 1)
SpaceCharacters = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function DoubleCharacters(Text As String, AmountOfExtras As Integer)
'This will double every character in the text, like this:  CCooooll
On Error GoTo error
Dim i As Long
Dim i2 As Long
Dim endstr As String
If AmountOfExtras > 100 Then
AmountOfExtras = 100
ElseIf AmountOfExtras < 1 Then
AmountOfExtras = 1
End If
For i = 1 To Len(Text$)
  For i2 = 1 To AmountOfExtras
  endstr$ = endstr$ & Mid$(Text$, i, 1)
  Next i2
Next i
DoubleCharacters = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function EchoText(Text As String, Reverse As Boolean)
'This will "echo" the text, like this:  Cool ool ol l
On Error GoTo error
Dim i As Long
Dim endstr As String
For i = 1 To Len(Text$)
  If Reverse = True Then
  endstr$ = Mid$(Text$, i, Len(Text$) - (i - 1)) & " " & endstr$
  Else
  endstr$ = endstr$ & Mid$(Text$, i, Len(Text$) - (i - 1)) & " "
  End If
Next i
endstr$ = Mid$(endstr$, 1, Len(endstr$) - 1)
EchoText = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Scramble(Text As String, Key As Integer)
'This will scramble text up, example:  oCol
On Error GoTo error
Dim RndNum As Long
Dim i As Long
Dim endstr As String
Dim ListN(10000) As Long
Dim CurPos As Long
Randomize Key
CurPos = 0
Text$ = Mid$(Text$, 1, 10000)
Start:
RndNum = Int((Len(Text$) - 1 + 1) * Rnd + 1)
For i = 0 To CurPos
  If RndNum = ListN(i) Then
  GoTo Start
  End If
Next i
ListN(CurPos) = RndNum
CurPos = CurPos + 1
If Not CurPos = Len(Text$) Then
GoTo Start
End If
For i = 0 To CurPos - 1
endstr$ = endstr$ & Mid$(Text$, ListN(i), 1)
Next i
Scramble = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function TwistText(Text As String)
'This will "twist" text, it is kind of like scramble, example:  oClo
Dim CurPos As Long
Dim endstr As String
CurPos = 1
Start:
endstr$ = endstr$ & Mid$(Text$, CurPos + 1, 1) & Mid$(Text$, CurPos, 1)
CurPos = CurPos + 2
Graph2 Len(Text$), CurPos
If Len(Text$) > CurPos Then
GoTo Start
ElseIf Len(Text$) = CurPos Then
endstr$ = endstr$ & Mid$(Text$, Len(Text$), 1)
End If
TwistText = endstr$
End Function

'Section 17:  Current Application Info

Function GetAppVersion()
'This will retrieve the current version of your application
On Error GoTo error
AppVersion = App.Major & "." & App.Minor & "." & App.Revision
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppName(ShowEXE As Boolean)
'This will get the application's .exe name
On Error GoTo error
GetAppName = App.EXEName
If ShowEXE = True Then
GetAppName = GetAppName & ".exe"
End If
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppPath()
'This will get the application's current path
On Error GoTo error
GetAppPath = App.Path
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppDescription()
'This will get the application's file description
On Error GoTo error
GetAppDescription = App.FileDescription
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppCopyRight()
'This will get the application's copyright
On Error GoTo error
GetAppCopyRight = App.LegalCopyright
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppComment()
'This will get the application's comment
On Error GoTo error
GetAppComment = App.Comments
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppTitle()
'This will get the application's title
On Error GoTo error
GetAppTitle = App.Title
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppCompanyName()
'This will get the application's company name
On Error GoTo error
GetAppCompanyName = App.CompanyName
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppProductName()
'This will get the application's product name
On Error GoTo error
GetAppProductName = App.ProductName
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

'Section 18:  Mouse Stuff

Public Sub MoveMouse(X As Integer, Y As Integer)
'Move the mouse
On Error GoTo error
Mouse.X = CLng(CDbl(X))
Mouse.Y = CLng(CDbl(Y))
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Function MousePosition()
'Get the mouse's current position
On Error GoTo error
If Index = 0 Then
MousePosition = Mid$(Str$(Mouse.X), 2, Len(Str$(Mouse.X)) - 1)
MousePosition = MousePosition + "," + Str$(Mouse.Y)
End If
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Sub LeftClick()
'Mouse left click
On Error GoTo error
LeftDown
LeftUp
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub LeftDown()
'Mouse left down
On Error GoTo error
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub LeftUp()
'Mouse left up
On Error GoTo error
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub MiddleClick()
'Mouse middle click
On Error GoTo error
MiddleDown
MiddleUp
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub MiddleDown()
'Mouse middle down
On Error GoTo error
mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub MiddleUp()
'Mouse middle up
On Error GoTo error
mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub RightClick()
'Mouse right click
On Error GoTo error
RightDown
RightUp
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub RightDown()
'Mouse right down
On Error GoTo error
mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub RightUp()
'Mouse right up
On Error GoTo error
mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub HideMouse()
'Hide mouse cursor
On Error GoTo error
ShowCursor (bShow = False)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub ShowMouse()
'Show mouse cursor
On Error GoTo error
ShowCursor (bShow = True)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

'Section 19:  Draw

Public Sub DrawSquareOnForm(Frm As Form, X1 As Single, X2 As Single, Y1 As Single, Y2 As Single, Red As Integer, Green As Integer, Blue As Integer, Solid As Boolean)
'This will draw a square on a form
On Error GoTo error
If Solid = True Then
Frm.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), BF
Else
Frm.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), B
End If
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub DrawLineOnForm(Frm As Form, X1 As Single, X2 As Single, Y1 As Single, Y2 As Single, Red As Integer, Green As Integer, Blue As Integer)
'This will draw a line on a form
On Error GoTo error
Frm.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub DrawSquareOnPictureBox(Picture As PictureBox, X1 As Single, X2 As Single, Y1 As Single, Y2 As Single, Red As Integer, Green As Integer, Blue As Integer, Solid As Boolean)
'This will draw a square on a form
On Error GoTo error
If Solid = True Then
Picture.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), BF
Else
Picture.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), B
End If
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub DrawLineOnPictureBox(Picture As PictureBox, X1 As Single, X2 As Single, Y1 As Single, Y2 As Single, Red As Integer, Green As Integer, Blue As Integer)
'This will draw a line on a form
On Error GoTo error
Picture.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

'Section 20:  Misc

Function ConvertRGBToHex(Red As Double, Green As Double, Blue As Double)
'Convert RGB color coding to Hexidecimal color coding
On Error GoTo error
ConvertRGBToHex = RGBToHex(RGB(Blue, Green, Red))
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Private Function RGBToHex(RGB)
'For Convert RGB to Hexidecimal (and HTML fader):  Converts RGB to Hexidecimal
On Error GoTo error
Dim a As String
Dim B As Integer
a$ = Hex(RGB)
    B% = Len(a$)
    If B% = 5 Then a$ = "0" & a$
    If B% = 4 Then a$ = "00" & a$
    If B% = 3 Then a$ = "000" & a$
    If B% = 2 Then a$ = "0000" & a$
    If B% = 1 Then a$ = "00000" & a$
    RGBToHex = a$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function ConvertHexToRGB(HexCode As String)
'This will convert Hexidecimal color coding to RGB color coding
On Error GoTo error
HexCode$ = Mid$(HexCode$, 1, 6)
ConvertHexToRGB = HexToRGB(HexCode$)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Private Function HexToRGB(H As String) As Currency
'For Convert Hexidecimal to RGB:  Converts Hexidecimal to RGB
On Error GoTo error
Dim Tmp$
Dim lo1 As Integer, lo2 As Integer
Dim hi1 As Long, hi2 As Long
Const Hx = "&H"
Const BigShift = 65536
Const LilShift = 256, Two = 2
Tmp = H
If UCase(Left$(H, 2)) = "&H" Then Tmp = Mid$(H, 3)
Tmp = Right$("0000000" & Tmp, 8)
If IsNumeric(Hx & Tmp) Then
lo1 = CInt(Hx & Right$(Tmp, Two))
hi1 = CLng(Hx & Mid$(Tmp, 5, Two))
lo2 = CInt(Hx & Mid$(Tmp, 3, Two))
hi2 = CLng(Hx & Left$(Tmp, Two))
HexToRGB = CCur(hi2 * LilShift + lo2) * BigShift + (hi1 * LilShift) + lo1
End If
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Sub WebPage(Address As String)
'Open a webpage in the default browser
On Error GoTo error
ret = Shell("Start.exe " & Address, 0)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Function RandomNumber(Max As Double, Min As Double)
'Create a random number
On Error GoTo error
Randomize Timer
RandomNumber = Int((Max - Min + 1) * Rnd + Min)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function MakeInputBox(DefaultText As String, Question As String, Title As String)
'This creates an input box
On Error GoTo error
MakeInputBox = InputBox(Question$, Title$, DefaultText$)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function LengthOfString(Text As String) As Long
'This will tell you how many characters are in a string
On Error GoTo error
LengthOfString = Len(Text$)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function FindAsciiOfChr(Chr As String)
'This will tell you the ascii of ONE CHARACTER (first one in the string)
On Error GoTo error
FindAsciiOfChr = Asc(Mid$(Chr$, 1, 1))
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function MakeChrFromAscii(Ascii As Integer)
'This will make a character out of ascii
On Error GoTo error
MakeChrFromAscii = Chr$(Ascii)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function MakeRndChrString(Length As Integer, Numbers As Boolean, Letters As Boolean, Symbols As Boolean, Other As Boolean) As String
'This will make a random string (good for passwords)
On Error GoTo error
Dim chrasc As Integer
Dim i As Integer
Dim endstr As String
Randomize Timer
If Length > 100 Then
Length = 100
ElseIf Length < 1 Then
Length = 1
End If
For i = 1 To Length
Start:
chrasc = Int((255 - 1 + 1) * Rnd + 1)
  If chrasc < 34 Then
    If Other = False Then
    GoTo Start
    End If
  ElseIf chrasc > 33 And chrasc < 49 Then
    If Symbols = False Then
    GoTo Start
    End If
  ElseIf chrasc > 48 And chrasc < 58 Then
    If Numbers = False Then
    GoTo Start
    End If
  ElseIf chrasc > 57 And chrasc < 65 Then
    If Symbols = False Then
    GoTo Start
    End If
  ElseIf chrasc > 64 And chrasc < 91 Then
    If Letters = False Then
    GoTo Start
    End If
  ElseIf chrasc > 90 And chrasc < 97 Then
    If Symbols = False Then
    GoTo Start
    End If
  ElseIf chrasc > 96 And chrasc < 123 Then
    If Letters = False Then
    GoTo Start
    End If
  ElseIf chrasc > 122 And chrasc < 127 Then
    If Symbols = False Then
    GoTo Start
    End If
  Else
    If Other = False Then
    GoTo Start
    End If
  End If
endstr$ = endstr$ & Chr$(chrasc)
Next i
MakeRndChrString = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Sub DoSendKeys(AppToActivate As String, AppActivateDelay As Integer, TextToSend As String, SendKeysDelay As Integer)
'This will use SendKeys to send text to an outside application
On Error GoTo error
AppActivate AppToActivate$, AppActivateDelay
SendKeys TextToSend$, SendKeysDelay
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Function GetTextFromListBox(ListB As ListBox, ListIndex As Long) As String
'This will get text from a listbox
On Error GoTo error
GetTextFromListBox = ListB.List(ListIndex)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetTextFromComboBox(ComboB As ComboBox, ListIndex As Long) As String
'This will get text from a combobox
On Error GoTo error
GetTextFromComboBox = ComboB.List(ListIndex)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function PasswordLock(Password As String)
'This will create an input box to create a simple password protection
On Error GoTo error
Dim xtra As String
Start:
xtra$ = InputBox("Please enter the password.", "Password Lock")
If xtra$ = Password$ Then
MsgBox "Correct Password!", vbExclamation, "Password Lock"
Else
  If MsgBox("Incorrect Password!  Would you like to try again?", 48 + vbYesNo, "Password Lock") = vbYes Then
  GoTo Start
  Else
  End
  End If
End If
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Sub ChangeDefaultDir(NewDirPath As String)
'This will change the default directory on a computer
On Error GoTo error
ChDir NewDirPath$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub ChangeDefaultDrive(NewDrive As String)
'This will change the default drive on a computer
On Error GoTo error
ChDrive NewDrive$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub MakeRegistrySetting(RegPath As String, Title As String, Data As String)
'This will make a registry setting
On Error GoTo error
a = MakeRegFile(&H80000002, RegPath$, Title$, Data$)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
