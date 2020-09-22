VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   LinkTopic       =   "Form6"
   MouseIcon       =   "Form6.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Form6.frx":030A
   ScaleHeight     =   315
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   600
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RelX(1 To 5) As Long
Private RelY(1 To 5) As Long

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim Returnval As Long
    
    RelX(1) = Left - Form1.Left
    RelY(1) = Top - Form1.Top
    RelX(2) = Left - Form2.Left
    RelY(2) = Top - Form2.Top
    RelX(3) = Left - Form3.Left
    RelY(3) = Top - Form3.Top
    RelX(4) = Left - Form4.Left
    RelY(4) = Top - Form4.Top
    RelX(5) = Left - Form5.Left
    RelY(5) = Top - Form5.Top

    Me.Visible = False
    Form5.Visible = False
    Timer1.Enabled = True

    x = ReleaseCapture()
    Returnval = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    Timer1.Enabled = False
    Me.Visible = True
    Form5.Visible = True
    Call Timer1_Timer

End Sub

Private Sub Timer1_Timer()

    Form1.Move Me.Left - RelX(1), Top - RelY(1)
    Form2.Move Me.Left - RelX(2), Top - RelY(2)
    Form3.Move Me.Left - RelX(3), Top - RelY(3)
    Form4.Move Me.Left - RelX(4), Top - RelY(4)
    Form5.Move Me.Left - RelX(5), Top - RelY(5)

End Sub


