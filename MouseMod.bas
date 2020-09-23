Attribute VB_Name = "MouseMod"
'**************************************
'Windows API/Global Declarations
'     mouse module, FINALLY!!! Move,click, ect
'
'**************************************
'**************************************
'
' Description:This module has the following functions (pretty self explanitory):
'
'     GetX, GetY, LeftClick, LeftDown, LeftUp,
'     RightClick, RightUp, RightDown,
'     MiddleClick, MiddleDown, MiddleUp,
'     MoveMouse, SetMousePos
'
'**************************************
' - Appened Oct 21 , 2000
' - By []Privare[]
'
'**************************************
Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Public Const MOUSEEVENTF_LEFTDOWN = &H2
    Public Const MOUSEEVENTF_LEFTUP = &H4
    Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
    Public Const MOUSEEVENTF_MIDDLEUP = &H40
    Public Const MOUSEEVENTF_RIGHTDOWN = &H8
    Public Const MOUSEEVENTF_RIGHTUP = &H10
    Public Const MOUSEEVENTF_MOVE = &H1


Public Type POINTAPI
    X As Long
    Y As Long
End Type
Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Const SW_SHOWNORMAL = 1
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWMAXIMIZED = 3




'**************************************
' + []Private
'**************************************
Public Steps As Integer
Public Recording As Boolean
Public Playing As Boolean
Public StartX As Long
Public StartY As Long
Public MaxX As Integer
Public MaxY As Integer
'**************************************
' -[]Pivate
'**************************************
Public Function FormsOnTop(frmForm As Form, fOnTop As Boolean)
'USAGE: ONTOP ME,TRUE   -ONTOP MOST
'       ONTOP ME,FALSE  -NOT TOP MOST
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Dim lState As Long
Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer
With frmForm
iLeft = .Left / Screen.TwipsPerPixelX
iTop = .Top / Screen.TwipsPerPixelY
iWidth = .Width / Screen.TwipsPerPixelX
iHeight = .Height / Screen.TwipsPerPixelY
End With
If fOnTop Then
lState = HWND_TOPMOST
Else
lState = HWND_NOTOPMOST
End If
Call SetWindowPos(frmForm.hwnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Function





Public Function GetX() As Long


    Dim n As POINTAPI
    GetCursorPos n
    GetX = n.X
End Function




Public Function GetY() As Long


    Dim n As POINTAPI
    GetCursorPos n
    GetY = n.Y
End Function




Public Sub LeftClick()


    LeftDown
    LeftUp
End Sub




Public Sub LeftDown()


    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub




Public Sub LeftUp()


    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub




Public Sub MiddleClick()


    MiddleDown
    MiddleUp
End Sub




Public Sub MiddleDown()


    mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
End Sub




Public Sub MiddleUp()


    mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
End Sub




Public Sub MoveMouse(xMove As Long, yMove As Long)


    mouse_event MOUSEEVENTF_MOVE, xMove, yMove, 0, 0
End Sub




Public Sub RightClick()


    RightDown
    RightUp
End Sub




Public Sub RightDown()


    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
End Sub




Public Sub RightUp()


    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub




Public Sub SetMousePos(xPos As Long, yPos As Long)


    SetCursorPos xPos, yPos
End Sub



Public Function GetWindowFromClass(toFind As String) As Boolean

Dim hwnd As Long  ' receives handle to the found window
Dim retval As Long  ' generic return value
GetWindowFromClass = False
'  Note how the CLng function
' must be used to force 0 as a Long data type.
hwnd = FindWindow(toFind, CLng(0))  ' look for the window
If hwnd = 0 Then  ' could not find the window
  MsgBox "!!!!  No Web browser currently running.!!!!", vbCritical ' if it were, there'd be a window!
Else
  retval = ShowWindow(hwnd, SW_SHOWMAXIMIZED)
  GetWindowFromClass = True
End If
End Function

