Attribute VB_Name = "Module_ScreenSaver"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SPI_GETACCESSTIMEOUT = 60
Public Const SPI_GETANIMATION = 72
Public Const SPI_GETBEEP = 1
Public Const SPI_GETBORDER = 5
Public Const SPI_GETDEFAULTINPUTLANG = 89
Public Const SPI_GETDRAGFULLWINDOWS = 38
Public Const SPI_GETFASTTASKSWITCH = 35
Public Const SPI_GETFILTERKEYS = 50
Public Const SPI_SCREENSAVERRUNNING = 97

Sub Main()
 Load frmProperties
 Select Case Left$(LCase(Command), 2)
  Case "/s"
   ShowCursor False    'hide the cursor
   frmScr.Show vbModal 'show the screen saver (vbSystemModal lets the forms keypress event system wide)
  Case "/c"
   frmProperties.Show vbModal
  Case "/p"
   hW1& = FindWindow(vbNullString, "Display Properties")
   hW2& = FindWindowEx(hW1, 0, vbNullString, "Screen Saver")
   hW3& = FindWindowEx(hW2, 0, "Static", vbNullString)
   hW4& = FindWindowEx(hW3, 0, "SSDemoParent", vbNullString)
   If hW4& > 0 Then
    frmPreview.Show
    frmPreview.Top = 0
    frmPreview.Left = 0
    SetParent frmPreview.hwnd, hW4
   End If
 End Select
End Sub
