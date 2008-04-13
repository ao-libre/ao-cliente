VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3105
      TabIndex        =   8
      Text            =   "1"
      Top             =   6690
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   435
      ScaleHeight     =   540
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   750
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6780
      Width           =   465
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Index           =   1
      Left            =   3855
      TabIndex        =   1
      Top             =   1800
      Width           =   2490
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Index           =   0
      Left            =   615
      TabIndex        =   0
      Top             =   1800
      Width           =   2490
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2265
      TabIndex        =   9
      Top             =   6750
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   1
      Left            =   3855
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6165
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   0
      Left            =   615
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6150
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3990
      TabIndex        =   7
      Top             =   975
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   3990
      TabIndex        =   6
      Top             =   630
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2730
      TabIndex        =   5
      Top             =   1170
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   1125
      TabIndex        =   4
      Top             =   450
      Width           =   45
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->

Public LasActionBuy As Boolean
Public LastIndex1 As Integer
Public LastIndex2 As Integer




Private Sub cantidad_Change()

If Val(cantidad.Text) < 1 Then
    cantidad.Text = 1
End If

If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
    cantidad.Text = 1
End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Command2_Click()
    Call WriteBankEnd
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub


Private Sub Form_Load()
'Cargamos la interfase
Me.Picture = LoadPicture(App.path & "\Graficos\comerciar.jpg")
Image1(0).Picture = LoadPicture(App.path & "\Graficos\BotónComprar.jpg")
Image1(1).Picture = LoadPicture(App.path & "\Graficos\Botónvender.jpg")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.path & "\Graficos\BotónComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.path & "\Graficos\Botónvender.jpg")
    Image1(1).Tag = 1
End If
End Sub

Private Sub Image1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

If List1(index).List(List1(index).listIndex) = "" Or _
   List1(index).listIndex < 0 Then Exit Sub

If Not IsNumeric(cantidad.Text) Then Exit Sub

Select Case index
    Case 0
        frmBancoObj.List1(0).SetFocus
        LastIndex1 = List1(0).listIndex
        LasActionBuy = True
        Call WriteBankExtractItem(List1(0).listIndex + 1, cantidad.Text)
        
   Case 1
        LastIndex2 = List1(1).listIndex
        LasActionBuy = False
        Call WriteBankDeposit(List1(1).listIndex + 1, cantidad.Text)
End Select

End Sub

Private Sub Image1_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case index
    Case 0
        If Image1(0).Tag = 1 Then
                Image1(0).Picture = LoadPicture(App.path & "\Graficos\BotónComprarApretado.jpg")
                Image1(0).Tag = 0
                Image1(1).Picture = LoadPicture(App.path & "\Graficos\Botónvender.jpg")
                Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
                Image1(1).Picture = LoadPicture(App.path & "\Graficos\Botónvenderapretado.jpg")
                Image1(1).Tag = 0
                Image1(0).Picture = LoadPicture(App.path & "\Graficos\BotónComprar.jpg")
                Image1(0).Tag = 1
        End If
        
End Select
End Sub

Private Sub list1_Click(index As Integer)
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

Select Case index
    Case 0
        Label1(0).Caption = UserBancoInventory(List1(0).listIndex + 1).Name
        Label1(2).Caption = UserBancoInventory(List1(0).listIndex + 1).Amount
        Select Case UserBancoInventory(List1(0).listIndex + 1).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe:" & UserBancoInventory(List1(0).listIndex + 1).MaxHit
                Label1(4).Caption = "Min Golpe:" & UserBancoInventory(List1(0).listIndex + 1).MinHit
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & UserBancoInventory(List1(0).listIndex + 1).Def
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        
        If UserBancoInventory(List1(0).listIndex + 1).Amount <> 0 Then _
            Call DrawGrhtoHdc(Picture1.hdc, UserBancoInventory(List1(0).listIndex + 1).GrhIndex, SR, DR)
    Case 1
        Label1(0).Caption = Inventario.ItemName(List1(1).listIndex + 1)
        Label1(2).Caption = Inventario.Amount(List1(1).listIndex + 1)
        Select Case Inventario.OBJType(List1(1).listIndex + 1)
            Case 2
                Label1(3).Caption = "Max Golpe:" & Inventario.MaxHit(List1(1).listIndex + 1)
                Label1(4).Caption = "Min Golpe:" & Inventario.MinHit(List1(1).listIndex + 1)
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & Inventario.Def(List1(1).listIndex + 1)
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        
        If Inventario.Amount(List1(1).listIndex + 1) <> 0 Then _
            Call DrawGrhtoHdc(Picture1.hdc, Inventario.GrhIndex(List1(1).listIndex + 1), SR, DR)
End Select

If Label1(2).Caption = 0 Then ' 27/08/2006 - GS > No mostrar imagen ni nada, cuando no ahi nada que mostrar.
    Label1(3).Visible = False
    Label1(4).Visible = False
    Picture1.Visible = False
Else
    Picture1.Visible = True
    Picture1.Refresh
End If

End Sub
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Private Sub List1_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.path & "\Graficos\BotónComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.path & "\Graficos\Botónvender.jpg")
    Image1(1).Tag = 1
End If
End Sub
