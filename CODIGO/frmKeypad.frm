VERSION 5.00
Begin VB.Form frmKeypad 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3930
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   262
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   405
      Width           =   5025
   End
   Begin VB.Image imgCerrar 
      Height          =   135
      Left            =   7080
      Top             =   120
      Width           =   135
   End
   Begin VB.Image imgEspacio 
      Height          =   405
      Left            =   2160
      Top             =   3435
      Width           =   3000
   End
   Begin VB.Image imgMin 
      Height          =   480
      Left            =   6270
      Top             =   3345
      Width           =   960
   End
   Begin VB.Image imgMay 
      Height          =   480
      Left            =   120
      Top             =   3345
      Width           =   960
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   45
      Left            =   6030
      Top             =   2760
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   44
      Left            =   5490
      Top             =   2805
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   43
      Left            =   4935
      Top             =   2805
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   42
      Left            =   4350
      Top             =   2790
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   41
      Left            =   3780
      Top             =   2820
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   40
      Left            =   3240
      Top             =   2820
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   39
      Left            =   2700
      Top             =   2820
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   38
      Left            =   2160
      Top             =   2835
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   37
      Left            =   1590
      Top             =   2835
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   36
      Left            =   1035
      Top             =   2820
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   35
      Left            =   6315
      Top             =   2175
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   34
      Left            =   5760
      Top             =   2235
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   33
      Left            =   5205
      Top             =   2205
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   32
      Left            =   4680
      Top             =   2205
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   31
      Left            =   4125
      Top             =   2220
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   30
      Left            =   3555
      Top             =   2175
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   29
      Left            =   3000
      Top             =   2220
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   28
      Left            =   2430
      Top             =   2205
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   27
      Left            =   1890
      Top             =   2205
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   26
      Left            =   1335
      Top             =   2235
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   25
      Left            =   780
      Top             =   2235
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   24
      Left            =   6600
      Top             =   1605
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   23
      Left            =   6060
      Top             =   1605
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   22
      Left            =   5475
      Top             =   1620
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   21
      Left            =   4905
      Top             =   1620
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   20
      Left            =   4380
      Top             =   1620
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   19
      Left            =   3825
      Top             =   1650
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   18
      Left            =   3270
      Top             =   1650
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   17
      Left            =   2730
      Top             =   1620
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   16
      Left            =   2175
      Top             =   1650
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   15
      Left            =   1635
      Top             =   1650
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   14
      Left            =   1065
      Top             =   1650
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   13
      Left            =   510
      Top             =   1650
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   12
      Left            =   6825
      Top             =   960
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   11
      Left            =   6285
      Top             =   975
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   10
      Left            =   5730
      Top             =   1020
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   9
      Left            =   5190
      Top             =   960
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   8
      Left            =   4635
      Top             =   960
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   7
      Left            =   4080
      Top             =   960
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   6
      Left            =   3525
      Top             =   960
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   5
      Left            =   2955
      Top             =   960
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   4
      Left            =   2415
      Top             =   960
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   3
      Left            =   1860
      Top             =   960
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   2
      Left            =   1305
      Top             =   975
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   1
      Left            =   750
      Top             =   990
      Width           =   570
   End
   Begin VB.Image imgKeyPad 
      Height          =   555
      Index           =   0
      Left            =   165
      Top             =   975
      Width           =   570
   End
End
Attribute VB_Name = "frmKeypad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmKepad.frm
'
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

Private clsFormulario As clsFormMovementManager

Private Enum e_modo_keypad
    MINUSCULA
    MAYUSCULA
End Enum

Private MinMayBack(1) As Picture
Private cBotonMin As clsGraphicalButton
Private cBotonMay As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Const MinIndex = "1234567890-=\qwertyuiop[]asdfghjkl;'zxcvbnm,./"
Private Const MayIndex = "!@#$%^&*()_+|QWERTYUIOP{}ASDFGHJKL:""ZXCVBNM<>?"
Private Modo As e_modo_keypad

Private Sub Form_Activate()
Dim i As Integer
Dim j As Integer
    i = RandomNumber(-2000, 2000)
    j = RandomNumber(-350, 350)
    Me.Top = Me.Top + j
    Me.Left = Me.Left + i

End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Dim GrhPath As String
    GrhPath = DirGraficos
    Set MinMayBack(0) = LoadPicture(GrhPath & "TecladoMinuscula.jpg")
    Set MinMayBack(1) = LoadPicture(GrhPath & "TecladoMayuscula.jpg")
    
    Call LoadButtons
    
    Me.Picture = MinMayBack(e_modo_keypad.MINUSCULA)
    
    Modo = MINUSCULA
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonMin = New clsGraphicalButton
    Set cBotonMay = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonMin.Initialize(imgMay, GrhPath & "BotonMay.jpg", _
                                    GrhPath & "BotonMayRollover.jpg", _
                                    GrhPath & "BotonMayClick.jpg", Me)

    Call cBotonMay.Initialize(imgMin, GrhPath & "BotonMin.jpg", _
                                    GrhPath & "BotonMinRollover.jpg", _
                                    GrhPath & "BotonMinClick.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgEspacio_Click()
    Call Audio.PlayWave(SND_CLICK)
    Me.txtPassword.Text = Me.txtPassword.Text & " "
    Me.txtPassword.SetFocus
End Sub

Private Sub imgKeyPad_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)
    
    If Modo = MAYUSCULA Then
        Me.txtPassword.Text = Me.txtPassword.Text & mid$(MayIndex, Index + 1, 1)
    Else
        Me.txtPassword.Text = Me.txtPassword.Text & mid$(MinIndex, Index + 1, 1)
    End If
    
    Me.txtPassword.SetFocus
End Sub

Private Sub imgKeyPad_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgMay_Click()
    If Modo = MAYUSCULA Then Exit Sub
    
    'Call Audio.PlayWave(SND_CLICK)
    Me.Picture = MinMayBack(e_modo_keypad.MAYUSCULA)  'LoadPicture(App.path & "\graficos\KeyPadMay.bmp")
    Modo = MAYUSCULA
    Me.txtPassword.SetFocus
End Sub

Private Sub imgMin_Click()
    If Modo = MINUSCULA Then Exit Sub
    
    'Call Audio.PlayWave(SND_CLICK)
    Me.Picture = MinMayBack(e_modo_keypad.MINUSCULA) 'LoadPicture(App.path & "\graficos\KeyPadMin.bmp")
    Modo = MINUSCULA
    Me.txtPassword.SetFocus
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmConnect.txtPasswd.Text = Me.txtPassword.Text
        Unload Me
    Else
        Me.txtPassword.Text = vbNullString
    End If
End Sub
