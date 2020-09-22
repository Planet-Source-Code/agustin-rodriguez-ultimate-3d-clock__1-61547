VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9465
   Enabled         =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8520
      Top             =   6360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Run only in Windows XP.

'Use the icone that will appear in sys tray for configurations.

Option Explicit

Private Sub Form_Activate()

    Form2.ZOrder 0
    Form3.ZOrder 0
    Form4.ZOrder 0

End Sub

Private Sub Form_Load()

  Dim Ret As Long
  Dim PosX As Long
  Dim PosY As Long
  
    PosX = Val(GetSetting("Relogio Virtual", "Coordenadas", "LEFT", "0"))
    PosY = Val(GetSetting("Relogio Virtual", "Coordenadas", "TOP", "0"))

    Move PosX, PosY
    Form2.Move PosX, PosY
    Form3.Move PosX, PosY
    Form4.Move PosX, PosY

    Form2.Picture1.Picture = LoadPicture(App.Path & "\PH.bmp")

    Form3.Picture1(0).Picture = LoadPicture(App.Path & "\PM1.bmp")
    Form3.Picture1(1).Picture = LoadPicture(App.Path & "\PM2.bmp")
    Form3.Picture1(2).Picture = LoadPicture(App.Path & "\PM3.bmp")
    Form3.Picture1(3).Picture = LoadPicture(App.Path & "\PM4.bmp")

    Form4.Picture1(0).Picture = LoadPicture(App.Path & "\PS1.bmp")
    Form4.Picture1(1).Picture = LoadPicture(App.Path & "\PS2.bmp")
    Form4.Picture1(2).Picture = LoadPicture(App.Path & "\PS3.bmp")
    Form4.Picture1(3).Picture = LoadPicture(App.Path & "\PS4.bmp")
    
    Form5.Display_Name_Click (Val(GetSetting("Relogio Virtual", "Picture", "Index", "0")))
      
    Ret = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Form1.hwnd, GWL_EXSTYLE, Ret
  
    Ret = GetWindowLong(Form2.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Form2.hwnd, GWL_EXSTYLE, Ret
    
    Ret = GetWindowLong(Form3.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Form3.hwnd, GWL_EXSTYLE, Ret
    
    Ret = GetWindowLong(Form4.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Form4.hwnd, GWL_EXSTYLE, Ret
    
    Ret = GetWindowLong(Form5.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Form5.hwnd, GWL_EXSTYLE, Ret
     
    Ret = GetWindowLong(Form6.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Form6.hwnd, GWL_EXSTYLE, Ret
    
    Col = RGB(0, 0, 0)
    
    Background_Index = Val(GetSetting("Relogio Virtual", "Opacidade", "Background", 0))
    Arrows_Index = Val(GetSetting("Relogio Virtual", "Opacidade", "Arrows", 0))
    
    Form5.Background_opaque_Click (Background_Index)
    Form5.Arrows_opaque_Click (Arrows_Index)
    
    On_top_value = Val(GetSetting("Relogio Virtual", "On Top", "Valor", 0))
    Form5.On_top.Checked = On_top_value
    
    If On_top_value Then
        Form5.Put_On_top
    End If
    
End Sub

Private Sub Form_Resize()

    Form2.Show
    Form3.Show
    Form4.Show
    Form5.Show
    Form6.Show
    Form5.Move Left + 5280, Top + 1200
    Form6.Move Left + 5210, Top + 1490
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SaveSetting "Relogio Virtual", "Coordenadas", "LEFT", Str(Left)
    SaveSetting "Relogio Virtual", "Coordenadas", "TOP", Str(Top)
    SaveSetting "Relogio Virtual", "Opacidade", "Background", Str(Background_Index)
    SaveSetting "Relogio Virtual", "Opacidade", "Arrows", Str(Arrows_Index)
    SaveSetting "Relogio Virtual", "On Top", "Valor", Str(On_top_value)

End Sub

Private Sub Timer1_Timer()

  Static hora As Integer, minutos As Integer, segundos As Integer, x As String, Frame As Long

    'x = Right$(Now, 8)
    x = Format(Now, "hh:mm:ss")
    hora = Val(Mid$(x, 1, 2)) Mod 12
    minutos = Val(Mid$(x, 4, 2))
    segundos = Val(Mid$(x, 7, 2))
    
    With Form2
        Frame = ((hora Mod 12) * 2 - (minutos > 30)) * ((.Picture1.Height / 24))
        .Picture2.PaintPicture .Picture1, 0, 0, .Picture1.Width, .Picture1.Height / 24, 0, Frame, .Picture1.Width, .Picture1.Height / 24
    End With

    If minutos < 15 Then
        With Form3
            Frame = minutos * (.Picture1(0).Height / 15)
            .Picture2.PaintPicture .Picture1(0), 0, 0, .Picture1(0).Width, .Picture1(0).Height / 15, 0, Frame, .Picture1(0).Width, .Picture1(0).Height / 15
        End With
        GoTo siga
    End If
    
    If minutos > 14 And minutos < 30 Then
        minutos = minutos - 15
        With Form3
            Frame = minutos * (Form3.Picture1(1).Height / 15)
            .Picture2.PaintPicture Form3.Picture1(1), 0, 0, Form3.Picture1(1).Width, Form3.Picture1(1).Height / 15, 0, Frame, Form3.Picture1(1).Width, Form3.Picture1(1).Height / 15
        End With
        GoTo siga
    End If
    
    If minutos > 29 And minutos < 45 Then
        minutos = minutos - 30
        With Form3
            Frame = minutos * (.Picture1(2).Height / 15)
            .Picture2.PaintPicture .Picture1(2), 0, 0, .Picture1(2).Width, .Picture1(2).Height / 15, 0, Frame, .Picture1(2).Width, .Picture1(2).Height / 15
        End With
        GoTo siga
    End If
    
    If minutos > 44 Then
        minutos = minutos - 45
        With Form3
            Frame = minutos * (.Picture1(3).Height / 15)
            .Picture2.PaintPicture .Picture1(3), 0, 0, .Picture1(3).Width, .Picture1(3).Height / 15, 0, Frame, .Picture1(3).Width, .Picture1(3).Height / 15
        End With
    End If
  
siga:
    
    If segundos < 15 Then
        With Form4
            Frame = segundos * (.Picture1(0).Height / 15)
            .Picture2.PaintPicture .Picture1(0), 0, 0, .Picture1(0).Width, .Picture1(0).Height / 15, 0, Frame, .Picture1(0).Width, .Picture1(0).Height / 15
        End With
        GoTo fim
    End If
    
    If segundos > 14 And segundos < 30 Then
        segundos = segundos - 15
        With Form4
            Frame = segundos * (.Picture1(1).Height / 15)
            .Picture2.PaintPicture .Picture1(1), 0, 0, .Picture1(1).Width, .Picture1(1).Height / 15, 0, Frame, .Picture1(1).Width, .Picture1(1).Height / 15
        End With
        GoTo fim
    End If
    
    If segundos > 29 And segundos < 45 Then
        segundos = segundos - 30
        With Form4
            Frame = segundos * (.Picture1(2).Height / 15)
            .Picture2.PaintPicture .Picture1(2), 0, 0, .Picture1(2).Width, .Picture1(2).Height / 15, 0, Frame, .Picture1(2).Width, .Picture1(2).Height / 15
        End With
        GoTo fim
    End If
    
    If segundos > 44 Then
        segundos = segundos - 45
        With Form4
            Frame = segundos * (.Picture1(3).Height / 15)
            .Picture2.PaintPicture .Picture1(3), 0, 0, .Picture1(3).Width, .Picture1(3).Height / 15, 0, Frame, .Picture1(3).Width, .Picture1(3).Height / 15
        End With
    End If

fim:

End Sub


