Attribute VB_Name = "Module1"
Public Const HWND_TOPMOST As Integer = -1
Public Const HWND_NOTOPMOST As Integer = -2
Public Const SWP_NOMOVE As Integer = &H2
Public Const SWP_NOSIZE As Integer = &H1
Public Const HTCAPTION As Integer = 2
Public Const WM_NCLBUTTONDOWN As Integer = &HA1
Public Const LWA_COLORKEY As Integer = &H1
Public Const LWA_ALPHA As Integer = &H2
Public Const GWL_EXSTYLE As Integer = (-20)
Public Const WS_EX_LAYERED As Long = &H80000

Public Declare Function apiSetWindowPos Lib "User32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Background_Index As Integer
Public Arrows_Index As Integer
Public Col As Long
Public On_top_value As Integer


