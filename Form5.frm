VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   360
   ClientLeft      =   -30
   ClientTop       =   -420
   ClientWidth     =   435
   ControlBox      =   0   'False
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form5.frx":164A
   MousePointer    =   99  'Custom
   Picture         =   "Form5.frx":1954
   ScaleHeight     =   360
   ScaleWidth      =   435
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picNotifier 
      Height          =   630
      Left            =   1200
      ScaleHeight     =   570
      ScaleWidth      =   690
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   750
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Open 
         Caption         =   "Display"
         Begin VB.Menu Display_Name 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu Opacit 
         Caption         =   "Opacity"
         Begin VB.Menu Background 
            Caption         =   "Background"
            Begin VB.Menu Background_opaque 
               Caption         =   "100 %"
               Index           =   0
            End
            Begin VB.Menu Background_opaque 
               Caption         =   "  75 %"
               Index           =   1
            End
            Begin VB.Menu Background_opaque 
               Caption         =   "  50 %"
               Index           =   2
            End
            Begin VB.Menu Background_opaque 
               Caption         =   "  25 %"
               Index           =   3
            End
         End
         Begin VB.Menu Arrows 
            Caption         =   "Arrows"
            Begin VB.Menu Arrows_opaque 
               Caption         =   "100 %"
               Index           =   0
            End
            Begin VB.Menu Arrows_opaque 
               Caption         =   "  75 %"
               Index           =   1
            End
            Begin VB.Menu Arrows_opaque 
               Caption         =   "  50 %"
               Index           =   2
            End
            Begin VB.Menu Arrows_opaque 
               Caption         =   "  25 %"
               Index           =   3
            End
         End
      End
      Begin VB.Menu On_top 
         Caption         =   "On Top"
         Checked         =   -1  'True
      End
      Begin VB.Menu Visibilidade 
         Caption         =   "Hide"
      End
      Begin VB.Menu About 
         Caption         =   "About"
         Begin VB.Menu Info 
            Caption         =   "Agustin Rodriguez"
            Index           =   0
         End
         Begin VB.Menu Info 
            Caption         =   "Send E-mail"
            Index           =   1
         End
         Begin VB.Menu Info 
            Caption         =   "Goto my Home Page"
            Index           =   2
         End
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1

' Declaration of the Shell_NotifyIcon which we use
' to acomplish the task.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

' Constants
Private Const NIM_ADD As Integer = &H0
Private Const NIM_MODIFY As Integer = &H1
Private Const NIM_DELETE As Integer = &H2
Private Const WM_MOUSEMOVE As Integer = &H200
Private Const NIF_MESSAGE As Integer = &H1
Private Const NIF_ICON As Integer = &H2
Private Const NIF_TIP As Integer = &H4

' More Constants
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205

' The NOTIFYICONDATA Type Stuff
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

' We'll make the theForm dimmed as NOTIFYICONDATA
Private theForm As NOTIFYICONDATA

Private Qt As Integer

Public Sub Arrows_opaque_Click(Index As Integer)

  Dim I As Integer
  Dim Arrows_opacidade As Integer

    For I = 0 To 3
        If Arrows_opaque(I).Checked Then
            Arrows_opaque(I).Checked = False
        End If
    Next I

    Arrows_opaque(Index).Checked = True
    Arrows_opacidade = Int((100 - Index * 25) * 255 / 100)
    
    SetLayeredWindowAttributes Form2.hwnd, Col, Arrows_opacidade, LWA_COLORKEY Or LWA_ALPHA
    SetLayeredWindowAttributes Form3.hwnd, Col, Arrows_opacidade, LWA_COLORKEY Or LWA_ALPHA
    SetLayeredWindowAttributes Form4.hwnd, Col, Arrows_opacidade, LWA_COLORKEY Or LWA_ALPHA

    Arrows_Index = Index

End Sub

Public Sub Background_opaque_Click(Index As Integer)

  Dim I As Integer
  Dim Background_opacidade As Integer

    For I = 0 To 3
        If Background_opaque(I).Checked Then
            Background_opaque(I).Checked = False
        End If
    Next I

    Background_opaque(Index).Checked = True
    Background_opacidade = Int((100 - Index * 25) * 255 / 100)

    SetLayeredWindowAttributes Form1.hwnd, Col, Background_opacidade, LWA_COLORKEY Or LWA_ALPHA
    SetLayeredWindowAttributes Form5.hwnd, Col, Background_opacidade, LWA_COLORKEY Or LWA_ALPHA
    SetLayeredWindowAttributes Form6.hwnd, Col, Background_opacidade, LWA_COLORKEY Or LWA_ALPHA

    Background_Index = Index

End Sub

Public Sub Display_Name_Click(Index As Integer)

  Static I As Integer

    For I = 0 To Qt - 1
        If Display_Name(I).Checked Then
            Display_Name(I).Checked = False
        End If
    Next I
    Display_Name(Index).Checked = True
    Form1.Picture = LoadPicture(App.Path & "\" & Display_Name(Index).Caption & ".Display")
    SaveSetting "Relogio Virtual", "Picture", "Index", Index

End Sub

Private Sub Exit_Click()

  ' We also need to remove it when the program
  ' is ended.
        
  ' Change theForm's cbSize to theForm's length.

    theForm.cbSize = Len(theForm)
    
    ' Change theForm's hWnd to picNotifier's hWnd.
    theForm.hwnd = picNotifier.hwnd
    
    ' Change theForm's uId to 1&.
    theForm.uId = 1&
    
    ' Remove it from the TaskBar.
    Shell_NotifyIcon NIM_DELETE, theForm
    
    Unload Form1
    Unload Form2
    Unload Form3
    Unload Form4
    Unload Form6
    Unload Me
 
End Sub

Private Sub Form_Click()

    Exit_Click

End Sub

Private Sub Form_Load()

  Static x As String

    x = Dir$(App.Path & "\*.Display")
    Do While x <> ""
        If Qt > 0 Then
            Load Display_Name(Qt)
        End If
            
        Display_Name(Qt).Caption = Left$(x, Len(x) - 8)
        x = Dir
        Qt = Qt + 1
    Loop
    
    ' Center the Main Form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    ' Change theForm's cbSize to theForm's Length
    theForm.cbSize = Len(theForm)
    ' Use the picNotifier's hWnd as theForm's
    theForm.hwnd = picNotifier.hwnd
    ' Change the uId to 1&
    theForm.uId = 1&
    ' Use the respective Flags that should be used,
    ' so it works properly, just like any other.
    theForm.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    ' When there's WM_MOUSEMOVE, we'll need to see
    ' if there's mouse clicking, or whatever.
    theForm.ucallbackMessage = WM_MOUSEMOVE
    ' Use the Main Form's Icon for the process.
    theForm.hIcon = Me.Icon
    ' Use the Tip "Visual Basic Island" as our tooltip
    theForm.szTip = "Virtual Clock" & Chr$(0)
    ' Now, we actually add theForm to the Taskbar.
    Shell_NotifyIcon NIM_ADD, theForm
    ' Hide the Main Form.
    Me.Hide
    ' Make App.TaskVisible False.
    App.TaskVisible = False
    
End Sub

Private Sub Info_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 1
    ShellExecute hwnd, "open", "mailto:virtual_guitar_1@hotmail.com", vbNullString, vbNullString, conSwNormal
Case 2
    ShellExecute hwnd, "open", "www.foreverbahia.com.br/agustin", vbNullString, vbNullString, conSwNormal
End Select

End Sub

Public Sub On_top_Click()

    On_top.Checked = On_top.Checked Xor -1
    Select Case On_top.Checked
      Case False
        Put_no_On_top
      Case True
        Put_On_top
    End Select

    On_top_value = On_top.Checked

End Sub

Public Sub Put_no_On_top()

    apiSetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    apiSetWindowPos Form2.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    apiSetWindowPos Form3.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    apiSetWindowPos Form4.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    apiSetWindowPos Form5.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    apiSetWindowPos Form6.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE

End Sub

Public Sub Put_On_top()

    apiSetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    apiSetWindowPos Form2.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    apiSetWindowPos Form3.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    apiSetWindowPos Form4.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    apiSetWindowPos Form5.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    apiSetWindowPos Form6.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE

End Sub

Private Sub picNotifier_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  ' We'll use this sub to determine if the icon
  ' was doubleclicked, rightclicked or just a
  ' normal click.
    
  Static Rec As Boolean, Msg As Long

    ' Msg is the current X divided by the Screen's
    ' X in TwipsPerPixel Measurement's, so it's the
    ' same as the picNotifier.
    Msg = x / Screen.TwipsPerPixelX
    
    ' If Rec is False
    If Rec = False Then
        ' Make Rec True.
        Rec = True
        ' Determine what Msg really was:
        Select Case Msg
            ' If DoubleClick
          Case WM_LBUTTONDBLCLK:
            'Me.Show
            ' If Button is Down
          Case WM_LBUTTONDOWN:
            Form1.Show
            Form2.Show
            Form3.Show
            Form4.Show
            Form5.Show
            Form6.Show
            
            'If Button is Up
          Case WM_LBUTTONUP:
            PopupMenu Menu
            
            'If the RightButton is clicked
          Case WM_RBUTTONDBLCLK:
            
            'If the RightBurron is Down
          Case WM_RBUTTONDOWN:
            
            'If RightButton is Up
          Case WM_RBUTTONUP:
            PopupMenu Menu
            'End Determination
        End Select
        
        'Change Rec Back to False.
        Rec = False
    End If

End Sub

Private Sub Visibilidade_Click()

    If Form1.Visible Then
        Form1.Hide
        Form2.Hide
        Form3.Hide
        Form4.Hide
        Form5.Hide
        Form6.Hide
    End If

End Sub


