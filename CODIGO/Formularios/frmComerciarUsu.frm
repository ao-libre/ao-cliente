VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmComerciarUsu 
   BorderStyle     =   0  'None
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   590
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picInvOroProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3450
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   930
      Width           =   960
   End
   Begin VB.TextBox txtAgregar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4500
      TabIndex        =   6
      Top             =   2295
      Width           =   1035
   End
   Begin VB.PictureBox picInvOroOfertaOtro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5610
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   5040
      Width           =   960
   End
   Begin VB.PictureBox picInvOfertaOtro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   6975
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   4
      Top             =   5040
      Width           =   2400
   End
   Begin VB.PictureBox picInvOfertaProp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   6960
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   3
      Top             =   930
      Width           =   2400
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   495
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   7965
      Width           =   6060
   End
   Begin VB.PictureBox picInvComercio 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   630
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   1
      Top             =   945
      Width           =   2400
   End
   Begin VB.PictureBox picInvOroOfertaProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5610
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   930
      Width           =   960
   End
   Begin RichTextLib.RichTextBox CommerceConsole 
      Height          =   1620
      Left            =   495
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   6030
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   2858
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmComerciarUsu.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton imgCancelar 
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   4680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Cancelar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciarUsu.frx":007E
      PICF            =   "frmComerciarUsu.frx":009A
      PICH            =   "frmComerciarUsu.frx":00B6
      PICV            =   "frmComerciarUsu.frx":00D2
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton imgConfirmar 
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Confirmar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciarUsu.frx":00EE
      PICF            =   "frmComerciarUsu.frx":010A
      PICH            =   "frmComerciarUsu.frx":0126
      PICV            =   "frmComerciarUsu.frx":0142
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton imgAceptar 
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   8160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Aceptar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciarUsu.frx":015E
      PICF            =   "frmComerciarUsu.frx":017A
      PICH            =   "frmComerciarUsu.frx":0196
      PICV            =   "frmComerciarUsu.frx":01B2
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton imgRechazar 
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   8160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Rechazar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciarUsu.frx":01CE
      PICF            =   "frmComerciarUsu.frx":01EA
      PICH            =   "frmComerciarUsu.frx":0206
      PICV            =   "frmComerciarUsu.frx":0222
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblQuitar 
      BackStyle       =   0  'Transparent
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   4680
      TabIndex        =   18
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblAgregar 
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblChat 
      BackStyle       =   0  'Transparent
      Caption         =   "Chat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblSuOferta 
      BackStyle       =   0  'Transparent
      Caption         =   "Su Oferta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblTuOferta 
      BackStyle       =   0  'Transparent
      Caption         =   "Tu Oferta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblInventario 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image imgAgregar 
      Height          =   255
      Left            =   4920
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image imgQuitar 
      Height          =   255
      Left            =   4920
      Top             =   2760
      Width           =   255
   End
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmComerciarUsu.frm
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


Private Const GOLD_OFFER_SLOT As Byte = INV_OFFER_SLOTS + 1

Private sCommerceChat As String

Private Sub imgAceptar_Click()
    If Not imgAceptar.Enabled Then Exit Sub  ' Deshabilitado
    
    Call WriteUserCommerceOk
    HabilitarAceptarRechazar False
    
End Sub

Private Sub imgAgregar_Click()
   
    ' No tiene seleccionado ningun item
    If InvComUsu.SelectedItem = 0 Then
        Call PrintCommerceMsg(JsonLanguage.item("MENSAJE_NO_SELECCIONASTE_NADA").item("TEXTO"), FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    ' Numero invalido
    If Not IsNumeric(txtAgregar.Text) Then Exit Sub
    
    HabilitarConfirmar True
    
    Dim OfferSlot As Byte
    Dim Amount As Long
    Dim InvSlot As Byte
    
    Dim MENSAJE_COMM_AGREGA As String
        MENSAJE_COMM_AGREGA = JsonLanguage.item("MENSAJE_COMM_AGREGA").item("TEXTO")
    
    With InvComUsu
        If .SelectedItem = FLAGORO Then
            If Val(txtAgregar.Text) > InvOroComUsu(0).Amount(1) Then
                Call PrintCommerceMsg(JsonLanguage.item("MENSAJE_SIN_CANTIDAD_SUFICIENTE").item("TEXTO"), FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            Amount = InvOroComUsu(1).Amount(1) + Val(txtAgregar.Text)
    
            ' Le aviso al otro de mi cambio de oferta
            Call WriteUserCommerceOffer(FLAGORO, Val(txtAgregar.Text), GOLD_OFFER_SLOT)
            
            ' Actualizo los inventarios
            Call InvOroComUsu(0).ChangeSlotItemAmount(1, InvOroComUsu(0).Amount(1) - Val(txtAgregar.Text))
            Call InvOroComUsu(1).ChangeSlotItemAmount(1, Amount)
            
            MENSAJE_COMM_AGREGA = Replace$(MENSAJE_COMM_AGREGA, "VAR_CANTIDAD_AGREGA", Val(txtAgregar.Text))
            MENSAJE_COMM_AGREGA = Replace$(MENSAJE_COMM_AGREGA, "VAR_QUE_AGREGA", "monedas de oro")
            If Val(txtAgregar.Text) = 1 Then
                MENSAJE_COMM_AGREGA = Replace$(MENSAJE_COMM_AGREGA, "monedas", "moneda")
            End If
            
            Call PrintCommerceMsg(MENSAJE_COMM_AGREGA, FontTypeNames.FONTTYPE_GUILD)
            
        ElseIf .SelectedItem > 0 Then
             If Val(txtAgregar.Text) > .Amount(.SelectedItem) Then
                Call PrintCommerceMsg(JsonLanguage.item("MENSAJE_SIN_CANTIDAD_SUFICIENTE").item("TEXTO"), FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
             
            OfferSlot = CheckAvailableSlot(.SelectedItem, Val(txtAgregar.Text))
            
            ' Hay espacio o lugar donde sumarlo?
            If OfferSlot > 0 Then
            
                MENSAJE_COMM_AGREGA = Replace$(MENSAJE_COMM_AGREGA, "VAR_CANTIDAD_AGREGA", Val(txtAgregar.Text))
                MENSAJE_COMM_AGREGA = Replace$(MENSAJE_COMM_AGREGA, "VAR_QUE_AGREGA", .ItemName(.SelectedItem))
                
                Call PrintCommerceMsg(MENSAJE_COMM_AGREGA, FontTypeNames.FONTTYPE_GUILD)
                            
                ' Le aviso al otro de mi cambio de oferta
                Call WriteUserCommerceOffer(.SelectedItem, Val(txtAgregar.Text), OfferSlot)
                
                ' Actualizo el inventario general de comercio
                Call .ChangeSlotItemAmount(.SelectedItem, .Amount(.SelectedItem) - Val(txtAgregar.Text))
                
                Amount = InvOfferComUsu(0).Amount(OfferSlot) + Val(txtAgregar.Text)
                
                ' Actualizo los inventarios
                If InvOfferComUsu(0).ObjIndex(OfferSlot) > 0 Then
                    ' Si ya esta el item, solo actualizo su cantidad en el invenatario
                    Call InvOfferComUsu(0).ChangeSlotItemAmount(OfferSlot, Amount)
                Else
                    InvSlot = .SelectedItem
                    ' Si no agrego todo
                    Call InvOfferComUsu(0).SetItem(OfferSlot, .ObjIndex(InvSlot), _
                                                    Amount, 0, .GrhIndex(InvSlot), .OBJType(InvSlot), _
                                                    .MaxHit(InvSlot), .MinHit(InvSlot), .MaxDef(InvSlot), .MinDef(InvSlot), _
                                                    .Valor(InvSlot), .ItemName(InvSlot), .Incompatible(InvSlot))
                End If
            End If
        End If
    End With
End Sub

Private Sub imgCancelar_Click()
    Call WriteUserCommerceEnd
End Sub

Private Sub imgConfirmar_Click()
    If Not imgConfirmar.Enabled Then Exit Sub  ' Deshabilitado
    
    HabilitarConfirmar False
    imgAgregar.Visible = False
    imgQuitar.Visible = False
    txtAgregar.Enabled = False
    
    Call PrintCommerceMsg(JsonLanguage.item("MENSAJE_COMM_OFERTA_COMFIRMADA").item("TEXTO"), FontTypeNames.FONTTYPE_CONSE)
    Call WriteUserCommerceConfirm
End Sub

Private Sub imgQuitar_Click()
    Dim Amount As Long
    
    Dim MENSAJE_COMM_SACA As String
    MENSAJE_COMM_SACA = JsonLanguage.item("MENSAJE_COMM_SACA").item("TEXTO")

    ' No tiene seleccionado ningun item
    If InvOfferComUsu(0).SelectedItem = 0 Then
        Call PrintCommerceMsg(JsonLanguage.item("MENSAJE_NO_SELECCIONASTE_NADA").item("TEXTO"), FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    ' Numero invalido
    If Not IsNumeric(txtAgregar.Text) Then Exit Sub

    ' Comparar con el inventario para distribuir los items
    If InvOfferComUsu(0).SelectedItem = FLAGORO Then
        Amount = IIf(Val(txtAgregar.Text) > InvOroComUsu(1).Amount(1), InvOroComUsu(1).Amount(1), Val(txtAgregar.Text))
        ' Estoy quitando, paso un valor negativo
        Amount = Amount * (-1)
        
        ' No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then
            ' Le aviso al otro de mi cambio de oferta
            Call WriteUserCommerceOffer(FLAGORO, Amount, GOLD_OFFER_SLOT)
            
            ' Actualizo los inventarios
            Call InvOroComUsu(0).ChangeSlotItemAmount(1, InvOroComUsu(0).Amount(1) - Amount)
            Call InvOroComUsu(1).ChangeSlotItemAmount(1, InvOroComUsu(1).Amount(1) + Amount)
            
            MENSAJE_COMM_SACA = Replace$(MENSAJE_COMM_SACA, "VAR_CANTIDAD_SACA", Amount * (-1))
            MENSAJE_COMM_SACA = Replace$(MENSAJE_COMM_SACA, "VAR_QUE_SACA", "monedas de oro")
            If Val(txtAgregar.Text) = 1 Then
                MENSAJE_COMM_SACA = Replace$(MENSAJE_COMM_SACA, "monedas", "moneda")
            End If
            
            Call PrintCommerceMsg(MENSAJE_COMM_SACA, FontTypeNames.FONTTYPE_GUILD)
        End If
    Else
        Amount = IIf(Val(txtAgregar.Text) > InvOfferComUsu(0).Amount(InvOfferComUsu(0).SelectedItem), _
                    InvOfferComUsu(0).Amount(InvOfferComUsu(0).SelectedItem), Val(txtAgregar.Text))
        ' Estoy quitando, paso un valor negativo
        Amount = Amount * (-1)
        
        ' No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then
            With InvOfferComUsu(0)
                
                MENSAJE_COMM_SACA = Replace$(MENSAJE_COMM_SACA, "VAR_CANTIDAD_SACA", Amount * (-1))
                MENSAJE_COMM_SACA = Replace$(MENSAJE_COMM_SACA, "VAR_QUE_SACA", .ItemName(.SelectedItem))
                
                Call PrintCommerceMsg(MENSAJE_COMM_SACA, FontTypeNames.FONTTYPE_GUILD)
    
                ' Le aviso al otro de mi cambio de oferta
                Call WriteUserCommerceOffer(0, Amount, .SelectedItem)
            
                ' Actualizo el inventario general
                Call UpdateInvCom(.ObjIndex(.SelectedItem), Abs(Amount))
                 
                 ' Actualizo el inventario de oferta
                 If .Amount(.SelectedItem) + Amount = 0 Then
                     ' Borro el item
                     Call .SetItem(.SelectedItem, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
                 Else
                     ' Le resto la cantidad deseada
                     Call .ChangeSlotItemAmount(.SelectedItem, .Amount(.SelectedItem) + Amount)
                 End If
            End With
        End If
    End If
    
    ' Si quito todos los items de la oferta, no puede confirmarla
    If Not HasAnyItem(InvOfferComUsu(0)) And _
       Not HasAnyItem(InvOroComUsu(1)) Then HabilitarConfirmar (False)
End Sub

Private Sub imgRechazar_Click()
    If Not imgRechazar.Enabled Then Exit Sub  ' Deshabilitado
    
    Call WriteUserCommerceReject
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaComercioUsuario.jpg")
    
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)
    
    Call PrintCommerceMsg("> " & JsonLanguage.item("MENSAJE_COMM_INFO").item("TEXTO").item(1), FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> " & JsonLanguage.item("MENSAJE_COMM_INFO").item("TEXTO").item(2), FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> " & JsonLanguage.item("MENSAJE_COMM_INFO").item("TEXTO").item(3), FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> " & JsonLanguage.item("MENSAJE_COMM_INFO").item("TEXTO").item(4), FontTypeNames.FONTTYPE_GUILDMSG)
    
End Sub

Private Sub Form_Activate()
On Error Resume Next

    InvComUsu.DrawInventory
    InvOfferComUsu(0).DrawInventory
    InvOfferComUsu(1).DrawInventory
    InvOroComUsu(0).DrawInventory
    InvOroComUsu(1).DrawInventory
    InvOroComUsu(2).DrawInventory

End Sub

Private Sub Form_GotFocus()
On Error Resume Next

    InvComUsu.DrawInventory
    InvOfferComUsu(0).DrawInventory
    InvOfferComUsu(1).DrawInventory
    InvOroComUsu(0).DrawInventory
    InvOroComUsu(1).DrawInventory
    InvOroComUsu(2).DrawInventory

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    InvComUsu.DrawInventory
    InvOfferComUsu(0).DrawInventory
    InvOfferComUsu(1).DrawInventory
    InvOroComUsu(0).DrawInventory
    InvOroComUsu(1).DrawInventory
    InvOroComUsu(2).DrawInventory

End Sub

Private Sub LoadTextsForm()
    lblInventario.Caption = JsonLanguage.item("FRM_COMERCIARUSU_LBLINVENTARIO").item("TEXTO")
    lblTuOferta.Caption = JsonLanguage.item("FRM_COMERCIARUSU_LBLTUOFERTA").item("TEXTO")
    lblAgregar.Caption = JsonLanguage.item("FRM_COMERCIARUSU_AGREGAR").item("TEXTO")
    lblQuitar.Caption = JsonLanguage.item("FRM_COMERCIARUSU_QUITAR").item("TEXTO")
    lblChat.Caption = JsonLanguage.item("FRM_COMERCIARUSU_CHAT").item("TEXTO")
    lblSuOferta.Caption = JsonLanguage.item("FRM_COMERCIARUSU_LBLSUOFERTA").item("TEXTO")
    imgRechazar.Caption = JsonLanguage.item("FRM_COMERCIARUSU_RECHAZAR").item("TEXTO")
    imgAceptar.Caption = JsonLanguage.item("FRM_COMERCIARUSU_ACEPTAR").item("TEXTO")
    imgConfirmar.Caption = JsonLanguage.item("FRM_COMERCIARUSU_CONFIRMAR").item("TEXTO")
    imgCancelar.Caption = JsonLanguage.item("FRM_COMERCIARUSU_CANCELAR").item("TEXTO")
End Sub

Private Sub Form_LostFocus()
    Me.SetFocus
End Sub

Private Sub SubtxtAgregar_Change()
    If Val(txtAgregar.Text) < 1 Then txtAgregar.Text = "1"

    If Val(txtAgregar.Text) > 2147483647 Then txtAgregar.Text = "2147483647"
End Sub

Private Sub picInvComercio_Click()
    Call InvOroComUsu(0).DeselectItem
End Sub

Private Sub picInvOfertaProp_Click()
    InvOroComUsu(1).DeselectItem
End Sub

Private Sub picInvOroOfertaOtro_Click()
    ' No se puede seleccionar el oro que oferta el otro :P
    InvOroComUsu(2).DeselectItem
End Sub

Private Sub picInvOroOfertaProp_Click()
    InvOfferComUsu(0).SelectGold
End Sub

Private Sub picInvOroProp_Click()
    InvComUsu.SelectGold
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 03/10/2009
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        sCommerceChat = JsonLanguage.item("MENSAJE_SOY_CHEATER")
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        sCommerceChat = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(sCommerceChat) <> 0 Then Call WriteCommerceChat(sCommerceChat)
        
        sCommerceChat = vbNullString
        SendTxt.Text = vbNullString
        KeyCode = 0
    End If
End Sub

Private Sub txtAgregar_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 03/10/2009
'**************************************************************
    'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
    Dim i As Long
    Dim tempstr As String
    Dim CharAscii As Integer
    
    For i = 1 To Len(txtAgregar.Text)
        CharAscii = Asc(mid$(txtAgregar.Text, i, 1))
        
        If CharAscii >= 48 And CharAscii <= 57 Then
            tempstr = tempstr & Chr$(CharAscii)
        End If
    Next i
    
    If tempstr <> txtAgregar.Text Then
        'We only set it if it's different, otherwise the event will be raised
        'constantly and the client will crush
        txtAgregar.Text = tempstr
    End If
End Sub

Private Sub txtAgregar_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or _
            KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
        KeyCode = 0
    End If

End Sub

Private Sub txtAgregar_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or _
            KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
        'txtCant = KeyCode
        KeyAscii = 0
    End If

End Sub

Private Function CheckAvailableSlot(ByVal InvSlot As Byte, ByVal Amount As Long) As Byte
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Search for an available slot to put an item. If found returns the slot, else returns 0.
'***************************************************
    Dim slot As Long
On Error GoTo Err
    ' Primero chequeo si puedo sumar esa cantidad en algun slot que ya tenga ese item
    For slot = 1 To INV_OFFER_SLOTS
        If InvComUsu.ObjIndex(InvSlot) = InvOfferComUsu(0).ObjIndex(slot) Then
            If InvOfferComUsu(0).Amount(slot) + Amount <= MAX_INVENTORY_OBJS Then
                ' Puedo sumarlo aca
                CheckAvailableSlot = slot
                Exit Function
            End If
        End If
    Next slot
    
    ' No lo puedo sumar, me fijo si hay alguno vacio
    For slot = 1 To INV_OFFER_SLOTS
        If InvOfferComUsu(0).ObjIndex(slot) = 0 Then
            ' Esta vacio, lo dejo aca
            CheckAvailableSlot = slot
            Exit Function
        End If
    Next slot
    Exit Function
Err:
    Debug.Print "Slot: " & slot
End Function

Public Sub UpdateInvCom(ByVal ObjIndex As Integer, ByVal Amount As Long)
    Dim slot As Byte
    Dim RemainingAmount As Long
    Dim DifAmount As Long
    
    RemainingAmount = Amount
    
    For slot = 1 To MAX_INVENTORY_SLOTS
        
        If InvComUsu.ObjIndex(slot) = ObjIndex Then
            DifAmount = Inventario.Amount(slot) - InvComUsu.Amount(slot)
            If DifAmount > 0 Then
                If RemainingAmount > DifAmount Then
                    RemainingAmount = RemainingAmount - DifAmount
                    Call InvComUsu.ChangeSlotItemAmount(slot, Inventario.Amount(slot))
                Else
                    Call InvComUsu.ChangeSlotItemAmount(slot, InvComUsu.Amount(slot) + RemainingAmount)
                    Exit Sub
                End If
            End If
        End If
    Next slot
End Sub

Public Sub PrintCommerceMsg(ByRef msg As String, ByVal FontIndex As Integer)
    
    With FontTypes(FontIndex)
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, msg, .Red, .Green, .Blue, .bold, .italic)
    End With
    
End Sub

Public Function HasAnyItem(ByRef Inventory As clsGraphicalInventory) As Boolean

    Dim slot As Long
    
    For slot = 1 To Inventory.MaxObjs
        If Inventory.Amount(slot) > 0 Then HasAnyItem = True: Exit Function
    Next slot
    
End Function

Public Sub HabilitarConfirmar(ByVal Habilitar As Boolean)
    imgConfirmar.Enabled = Habilitar
End Sub

Public Sub HabilitarAceptarRechazar(ByVal Habilitar As Boolean)
    imgAceptar.Enabled = Habilitar
    imgRechazar.Enabled = Habilitar
End Sub
