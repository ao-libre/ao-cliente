VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   560
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox CantidadOro 
      Height          =   465
      Left            =   3045
      TabIndex        =   7
      Text            =   "1"
      Top             =   1320
      Width           =   840
   End
   Begin VB.PictureBox PicBancoInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   420
      ScaleHeight     =   4005
      ScaleWidth      =   2490
      TabIndex        =   6
      Top             =   2700
      Width           =   2520
   End
   Begin VB.PictureBox PicInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   3975
      ScaleHeight     =   17.676
      ScaleMode       =   0  'User
      ScaleWidth      =   856.774
      TabIndex        =   5
      Top             =   2700
      Width           =   2520
   End
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
      Height          =   465
      Left            =   3045
      TabIndex        =   4
      Text            =   "1"
      Top             =   4050
      Width           =   840
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   600
      ScaleHeight     =   540
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   7200
      Width           =   555
   End
   Begin VB.Label lblUserGld 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2040
      TabIndex        =   8
      Top             =   945
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   930
      Index           =   1
      Left            =   4755
      Tag             =   "0"
      Top             =   1155
      Width           =   1050
   End
   Begin VB.Image Image2 
      Height          =   930
      Index           =   0
      Left            =   1380
      Tag             =   "0"
      Top             =   1155
      Width           =   570
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   6480
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   3210
      Top             =   4740
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   3210
      Top             =   3480
      Width           =   495
   End
   Begin VB.Image CmdMoverBov 
      Height          =   375
      Index           =   1
      Left            =   0
      Top             =   4200
      Width           =   570
   End
   Begin VB.Image CmdMoverBov 
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   4560
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   7470
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   7725
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   7230
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
Public NoPuedeMover As Boolean

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

Private Sub CantidadOro_Change()
If Val(CantidadOro.Text) < 1 Then
    cantidad.Text = 1
End If

If Val(CantidadOro.Text) > &H7FFFFFFF Then
    CantidadOro.Text = &H7FFFFFFF
End If
End Sub

Private Sub CantidadOro_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub CmdMoverBov_Click(index As Integer)
'If InvBanco(0).SelectedItem = 0 Then Exit Sub

'If NoPuedeMover Then Exit Sub

'Select Case index
'    Case 1 'subir
'        If InvBanco(0).SelectedItem <= 1 Then
'            With FontTypes(FontTypeNames.FONTTYPE_INFO)
'                Call ShowConsoleMsg("No puedes mover el objeto en esa dirección.", .red, .green, .blue, .bold, .italic)
'            End With
'            Exit Sub
'        End If
'        LastIndex1 = InvBanco(0).SelectedItem - 1
'    Case 0 'bajar
'        If InvBanco(0).SelectedItem >= MAX_BANCOINVENTORY_SLOTS Then
'            With FontTypes(FontTypeNames.FONTTYPE_INFO)
'                Call ShowConsoleMsg("No puedes mover el objeto en esa dirección.", .red, .green, .blue, .bold, .italic)
'            End With
'            Exit Sub
'        End If
'        LastIndex1 = InvBanco(0).SelectedItem - 1
'End Select

'NoPuedeMover = True
'LasActionBuy = True
'LastIndex2 = InvBanco(1).SelectedItem
'Call WriteMoveBank(index, InvBanco(0).SelectedItem)
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub


Private Sub Form_Load()
'Cargamos la interfase
Me.Picture = LoadPicture(App.path & "\Graficos\Boveda.bmp")
'Image1(0).Picture = LoadPicture(App.path & "\Graficos\BotónComprar.jpg")
'Image1(1).Picture = LoadPicture(App.path & "\Graficos\Botónvender.jpg")

'CmdMoverBov(1).Picture = LoadPicture(App.path & "\Graficos\FlechaSubirObjeto.jpg")
'CmdMoverBov(0).Picture = LoadPicture(App.path & "\Graficos\FlechaBajarObjeto.jpg")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image2(0).Tag = 0 Then
    Image2(0).Picture = LoadPicture()
    Image2(0).Tag = 1
End If
If Image2(1).Tag = 0 Then
    Image2(1).Picture = LoadPicture()
    Image2(1).Tag = 1
End If
End Sub

Private Sub Image1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

If InvBanco(index).SelectedItem = 0 Then Exit Sub

If Not IsNumeric(cantidad.Text) Then Exit Sub

Select Case index
    Case 0
        LastIndex1 = InvBanco(0).SelectedItem
        LasActionBuy = True
        Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text)
        
   Case 1
        LastIndex2 = InvBanco(1).SelectedItem
        LasActionBuy = False
        Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text)
End Select

End Sub

Private Sub Image2_Click(index As Integer)

Select Case index
    Case 0
        Call WriteBankDepositGold(Val(CantidadOro.Text))
    Case 1
        Call WriteBankExtractGold(Val(CantidadOro.Text))
End Select
End Sub

Private Sub Image2_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case index
    Case 0
        If Image2(0).Tag = 1 Then
            Image2(0).Picture = LoadPicture(App.path & "\Graficos\BotónDepositaOroApretado.bmp")
            Image2(0).Tag = 0
            Image2(1).Picture = LoadPicture()
            Image2(1).Tag = 1
        End If
        
    Case 1
        If Image2(1).Tag = 1 Then
            Image2(1).Picture = LoadPicture(App.path & "\Graficos\BotónRetirarOroApretado.bmp")
            Image2(1).Tag = 0
            Image2(0).Picture = LoadPicture()
            Image2(0).Tag = 1
        End If
        
End Select
End Sub

Private Sub PicBancoInv_Click()
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

If InvBanco(0).SelectedItem <> 0 Then
    With UserBancoInventory(InvBanco(0).SelectedItem)
        Label1(0).Caption = .Name
        
        Select Case .OBJType
            Case 2, 32
                Label1(1).Caption = "Max Golpe:" & .MaxHit
                Label1(2).Caption = "Min Golpe:" & .MinHit
                Label1(1).Visible = True
                Label1(2).Visible = True
                
            Case 3, 16, 17
                Label1(1).Caption = "Defensa:" & .Def
                Label1(1).Visible = True
                Label1(2).Visible = False
                
            Case Else
                Label1(1).Visible = False
                Label1(2).Visible = False
                
        End Select
        
        If .Amount <> 0 Then _
            Call DrawGrhtoHdc(Picture1.hdc, .GrhIndex, SR, DR)
            
        Picture1.Visible = True
        Picture1.Refresh
    End With
    
Else
    Label1(0).Caption = ""
    Label1(3).Visible = False
    Label1(4).Visible = False
    Picture1.BackColor = vbBlack
End If

End Sub

Private Sub PicBancoInv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image2(0).Tag = 0 Then
    Image2(0).Picture = LoadPicture()
    Image2(0).Tag = 1
End If
If Image2(1).Tag = 0 Then
    Image2(1).Picture = LoadPicture()
    Image2(1).Tag = 1
End If
End Sub

Private Sub PicInv_Click()
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

If InvBanco(1).SelectedItem <> 0 Then
    With Inventario
        Label1(0).Caption = .ItemName(InvBanco(1).SelectedItem)
        
        Select Case .OBJType(InvBanco(1).SelectedItem)
            Case 2, 32
                Label1(1).Caption = "Max Golpe:" & .MaxHit(InvBanco(1).SelectedItem)
                Label1(2).Caption = "Min Golpe:" & .MinHit(InvBanco(1).SelectedItem)
                Label1(1).Visible = True
                Label1(2).Visible = True
                
            Case 3, 16, 17
                Label1(1).Caption = "Defensa:" & .Def(InvBanco(1).SelectedItem)
                Label1(1).Visible = True
                Label1(2).Visible = False
                
            Case Else
                Label1(1).Visible = False
                Label1(2).Visible = False
                
        End Select
        
        If .Amount(InvBanco(1).SelectedItem) <> 0 Then _
            Call DrawGrhtoHdc(Picture1.hdc, .GrhIndex(InvBanco(1).SelectedItem), SR, DR)
            
        Picture1.Visible = True
        Picture1.Refresh
    End With
Else
    Label1(0).Caption = ""
    Label1(3).Visible = False
    Label1(4).Visible = False
    Picture1.BackColor = vbBlack
End If
End Sub

Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image2(0).Tag = 0 Then
    Image2(0).Picture = LoadPicture()
    Image2(0).Tag = 1
End If
If Image2(1).Tag = 0 Then
    Image2(1).Picture = LoadPicture()
    Image2(1).Tag = 1
End If
End Sub

Private Sub Image3_Click()
    Call WriteBankEnd
    NoPuedeMover = False
End Sub
