VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
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
   Picture         =   "frmComerciarUsu.frx":0000
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
      TextRTF         =   $"frmComerciarUsu.frx":4B687
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
   Begin VB.Image imgCancelar 
      Height          =   360
      Left            =   480
      Picture         =   "frmComerciarUsu.frx":4B705
      Tag             =   "1"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Image imgRechazar 
      Height          =   360
      Left            =   8220
      Picture         =   "frmComerciarUsu.frx":4F5F6
      Tag             =   "2"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image imgConfirmar 
      Height          =   360
      Left            =   7440
      Picture         =   "frmComerciarUsu.frx":536B7
      Tag             =   "2"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Image imgAceptar 
      Height          =   360
      Left            =   6750
      Picture         =   "frmComerciarUsu.frx":5797B
      Tag             =   "2"
      Top             =   8160
      Width           =   1455
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

Private clsFormulario         As clsFormMovementManager

Private cBotonAceptar         As clsGraphicalButton
Private cBotonCancelar        As clsGraphicalButton
Private cBotonRechazar        As clsGraphicalButton
Private cBotonConfirmar       As clsGraphicalButton
Public LastButtonPressed      As clsGraphicalButton

Private Const GOLD_OFFER_SLOT As Byte = INV_OFFER_SLOTS + 1

Private sCommerceChat         As String

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgAceptar_Click()
    
    On Error GoTo imgAceptar_Click_Err
    

    If Not cBotonAceptar.IsEnabled Then Exit Sub  ' Deshabilitado
    
    Call WriteUserCommerceOk
    HabilitarAceptarRechazar False
    
    
    Exit Sub

imgAceptar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "imgAceptar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgAgregar_Click()
    
    On Error GoTo imgAgregar_Click_Err
    
   
    ' No tiene seleccionado ningun item
    If InvComUsu.SelectedItem = 0 Then
        Call PrintCommerceMsg("�No tienes ning�n item seleccionado!", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub

    End If
    
    ' Numero invalido
    If Not IsNumeric(txtAgregar.Text) Then Exit Sub
    
    HabilitarConfirmar True
    
    Dim OfferSlot As Byte
    Dim Amount    As Long
    Dim InvSlot   As Byte
        
    With InvComUsu

        If .SelectedItem = FLAGORO Then
            If Val(txtAgregar.Text) > InvOroComUsu(0).Amount(1) Then
                Call PrintCommerceMsg("�No tienes esa cantidad!", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            Amount = InvOroComUsu(1).Amount(1) + Val(txtAgregar.Text)
    
            ' Le aviso al otro de mi cambio de oferta
            Call WriteUserCommerceOffer(FLAGORO, Val(txtAgregar.Text), GOLD_OFFER_SLOT)
            
            ' Actualizo los inventarios
            Call InvOroComUsu(0).ChangeSlotItemAmount(1, InvOroComUsu(0).Amount(1) - Val(txtAgregar.Text))
            Call InvOroComUsu(1).ChangeSlotItemAmount(1, Amount)
            
            Call PrintCommerceMsg("�Agregaste " & Val(txtAgregar.Text) & " moneda" & IIf(Val(txtAgregar.Text) = 1, "", "s") & " de oro a tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
            
        ElseIf .SelectedItem > 0 Then

            If Val(txtAgregar.Text) > .Amount(.SelectedItem) Then
                Call PrintCommerceMsg("�No tienes esa cantidad!", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
             
            OfferSlot = CheckAvailableSlot(.SelectedItem, Val(txtAgregar.Text))
            
            ' Hay espacio o lugar donde sumarlo?
            If OfferSlot > 0 Then
            
                Call PrintCommerceMsg("�Agregaste " & Val(txtAgregar.Text) & " " & .ItemName(.SelectedItem) & " a tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
                
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
                    Call InvOfferComUsu(0).SetItem(OfferSlot, .ObjIndex(InvSlot), Amount, 0, .GrhIndex(InvSlot), .OBJType(InvSlot), .MaxHit(InvSlot), .MinHit(InvSlot), .MaxDef(InvSlot), .MinDef(InvSlot), .Valor(InvSlot), .ItemName(InvSlot))

                End If

            End If

        End If

    End With

    
    Exit Sub

imgAgregar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "imgAgregar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgCancelar_Click()
    
    On Error GoTo imgCancelar_Click_Err
    
    Call WriteUserCommerceEnd

    
    Exit Sub

imgCancelar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "imgCancelar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgConfirmar_Click()
    
    On Error GoTo imgConfirmar_Click_Err
    

    If Not cBotonConfirmar.IsEnabled Then Exit Sub  ' Deshabilitado
    
    HabilitarConfirmar False
    imgAgregar.Visible = False
    imgQuitar.Visible = False
    txtAgregar.Enabled = False
    
    Call PrintCommerceMsg("�Has confirmado tu oferta! Ya no puedes cambiarla.", FontTypeNames.FONTTYPE_CONSE)
    Call WriteUserCommerceConfirm

    
    Exit Sub

imgConfirmar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "imgConfirmar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgQuitar_Click()
    
    On Error GoTo imgQuitar_Click_Err
    
    Dim Amount     As Long
    Dim InvComSlot As Byte

    ' No tiene seleccionado ningun item
    If InvOfferComUsu(0).SelectedItem = 0 Then
        Call PrintCommerceMsg("�No tienes ning�n �tem seleccionado!", FontTypeNames.FONTTYPE_FIGHT)
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
        
            Call PrintCommerceMsg("��Quitaste " & Amount * (-1) & " moneda" & IIf(Val(txtAgregar.Text) = 1, "", "s") & " de oro de tu oferta!!", FontTypeNames.FONTTYPE_GUILD)

        End If

    Else
        Amount = IIf(Val(txtAgregar.Text) > InvOfferComUsu(0).Amount(InvOfferComUsu(0).SelectedItem), InvOfferComUsu(0).Amount(InvOfferComUsu(0).SelectedItem), Val(txtAgregar.Text))
        ' Estoy quitando, paso un valor negativo
        Amount = Amount * (-1)
        
        ' No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then

            With InvOfferComUsu(0)
                
                Call PrintCommerceMsg("��Quitaste " & Amount * (-1) & " " & .ItemName(.SelectedItem) & " de tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
    
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
    If Not HasAnyItem(InvOfferComUsu(0)) And Not HasAnyItem(InvOroComUsu(1)) Then HabilitarConfirmar (False)

    
    Exit Sub

imgQuitar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "imgQuitar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgRechazar_Click()
    
    On Error GoTo imgRechazar_Click_Err
    

    If Not cBotonRechazar.IsEnabled Then Exit Sub  ' Deshabilitado
    
    Call WriteUserCommerceReject

    
    Exit Sub

imgRechazar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "imgRechazar_Click"
    End If
Resume Next
    
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    
    On Error GoTo Form_Load_Err
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Me.Picture = LoadPicture(DirGraficos & "VentanaComercioUsuario.jpg")
    
    LoadButtons
    
    Call PrintCommerceMsg("> Una vez termines de formar tu oferta, debes presionar en ""Confirmar"", tras lo cual ya no podr�s modificarla.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Luego que el otro usuario confirme su oferta, podr�s aceptarla o rechazarla. Si la rechazas, se terminar� el comercio.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Cuando ambos acepten la oferta del otro, se realizar� el intercambio.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Si se intercambian m�s �tems de los que pueden entrar en tu inventario, es probable que caigan al suelo, as� que presta mucha atenci�n a esto.", FontTypeNames.FONTTYPE_GUILDMSG)
    
    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    

    Dim GrhPath As String
    GrhPath = DirGraficos
    
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonConfirmar = New clsGraphicalButton
    Set cBotonRechazar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotonAceptarComUsu.jpg", GrhPath & "BotonAceptarRolloverComUsu.jpg", GrhPath & "BotonAceptarClickComUsu.jpg", Me, GrhPath & "BotonAceptarGrisComUsu.jpg", True)
                                    
    Call cBotonConfirmar.Initialize(imgConfirmar, GrhPath & "BotonConfirmarComUsu.jpg", GrhPath & "BotonConfirmarRolloverComUsu.jpg", GrhPath & "BotonConfirmarClickComUsu.jpg", Me, GrhPath & "BotonConfirmarGrisComUsu.jpg", True)
                                        
    Call cBotonRechazar.Initialize(imgRechazar, GrhPath & "BotonRechazarComUsu.jpg", GrhPath & "BotonRechazarRolloverComUsu.jpg", GrhPath & "BotonRechazarClickComUsu.jpg", Me, GrhPath & "BotonRechazarGrisComUsu.jpg", True)
                                        
    Call cBotonCancelar.Initialize(imgCancelar, GrhPath & "BotonCancelarComUsu.jpg", GrhPath & "BotonCancelarRolloverComUsu.jpg", GrhPath & "BotonCancelarClickComUsu.jpg", Me)
    
    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub Form_LostFocus()
    
    On Error GoTo Form_LostFocus_Err
    
    Me.SetFocus

    
    Exit Sub

Form_LostFocus_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "Form_LostFocus"
    End If
Resume Next
    
End Sub

Private Sub SubtxtAgregar_Change()
    
    On Error GoTo SubtxtAgregar_Change_Err
    

    If Val(txtAgregar.Text) < 1 Then txtAgregar.Text = "1"

    If Val(txtAgregar.Text) > 2147483647 Then txtAgregar.Text = "2147483647"

    
    Exit Sub

SubtxtAgregar_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "SubtxtAgregar_Change"
    End If
Resume Next
    
End Sub

Private Sub picInvComercio_Click()
    
    On Error GoTo picInvComercio_Click_Err
    
    Call InvOroComUsu(0).DeselectItem

    
    Exit Sub

picInvComercio_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "picInvComercio_Click"
    End If
Resume Next
    
End Sub

Private Sub picInvComercio_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)
    
    On Error GoTo picInvComercio_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

picInvComercio_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "picInvComercio_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub picInvOfertaOtro_MouseMove(Button As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    
    On Error GoTo picInvOfertaOtro_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

picInvOfertaOtro_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "picInvOfertaOtro_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub picInvOfertaProp_Click()
    
    On Error GoTo picInvOfertaProp_Click_Err
    
    InvOroComUsu(1).DeselectItem

    
    Exit Sub

picInvOfertaProp_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "picInvOfertaProp_Click"
    End If
Resume Next
    
End Sub

Private Sub picInvOfertaProp_MouseMove(Button As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    
    On Error GoTo picInvOfertaProp_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

picInvOfertaProp_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "picInvOfertaProp_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub picInvOroOfertaOtro_Click()
    ' No se puede seleccionar el oro que oferta el otro :P
    
    On Error GoTo picInvOroOfertaOtro_Click_Err
    
    InvOroComUsu(2).DeselectItem

    
    Exit Sub

picInvOroOfertaOtro_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "picInvOroOfertaOtro_Click"
    End If
Resume Next
    
End Sub

Private Sub picInvOroOfertaProp_Click()
    
    On Error GoTo picInvOroOfertaProp_Click_Err
    
    InvOfferComUsu(0).SelectGold

    
    Exit Sub

picInvOroOfertaProp_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "picInvOroOfertaProp_Click"
    End If
Resume Next
    
End Sub

Private Sub picInvOroProp_Click()
    
    On Error GoTo picInvOroProp_Click_Err
    
    InvComUsu.SelectGold

    
    Exit Sub

picInvOroProp_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "picInvOroProp_Click"
    End If
Resume Next
    
End Sub

Private Sub SendTxt_Change()
    
    On Error GoTo SendTxt_Change_Err
    

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 03/10/2009
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        sCommerceChat = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i         As Long
        Dim tempstr   As String
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

    
    Exit Sub

SendTxt_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "SendTxt_Change"
    End If
Resume Next
    
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    
    On Error GoTo SendTxt_KeyPress_Err
    

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

    
    Exit Sub

SendTxt_KeyPress_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "SendTxt_KeyPress"
    End If
Resume Next
    
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo SendTxt_KeyUp_Err
    

    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(sCommerceChat) <> 0 Then Call WriteCommerceChat(sCommerceChat)
        
        sCommerceChat = vbNullString
        SendTxt.Text = vbNullString
        KeyCode = 0

    End If

    
    Exit Sub

SendTxt_KeyUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "SendTxt_KeyUp"
    End If
Resume Next
    
End Sub

Private Sub txtAgregar_Change()
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 03/10/2009
    '**************************************************************
    'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
    
    On Error GoTo txtAgregar_Change_Err
    
    Dim i         As Long
    Dim tempstr   As String
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

    
    Exit Sub

txtAgregar_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "txtAgregar_Change"
    End If
Resume Next
    
End Sub

Private Sub txtAgregar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo txtAgregar_KeyDown_Err
    

    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
        KeyCode = 0

    End If

    
    Exit Sub

txtAgregar_KeyDown_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "txtAgregar_KeyDown"
    End If
Resume Next
    
End Sub

Private Sub txtAgregar_KeyPress(KeyAscii As Integer)
    
    On Error GoTo txtAgregar_KeyPress_Err
    

    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
        'txtCant = KeyCode
        KeyAscii = 0

    End If

    
    Exit Sub

txtAgregar_KeyPress_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "txtAgregar_KeyPress"
    End If
Resume Next
    
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
    
    On Error GoTo UpdateInvCom_Err
    
    Dim slot            As Byte
    Dim RemainingAmount As Long
    Dim DifAmount       As Long
    
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

    
    Exit Sub

UpdateInvCom_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "UpdateInvCom"
    End If
Resume Next
    
End Sub

Public Sub PrintCommerceMsg(ByRef msg As String, ByVal FontIndex As Integer)
    
    On Error GoTo PrintCommerceMsg_Err
    
    
    With FontTypes(FontIndex)
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, msg, .Red, .Green, .Blue, .bold, .italic)

    End With
    
    
    Exit Sub

PrintCommerceMsg_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "PrintCommerceMsg"
    End If
Resume Next
    
End Sub

Public Function HasAnyItem(ByRef Inventory As clsGrapchicalInventory) As Boolean
    
    On Error GoTo HasAnyItem_Err
    

    Dim slot As Long
    
    For slot = 1 To Inventory.MaxObjs

        If Inventory.Amount(slot) > 0 Then HasAnyItem = True: Exit Function
    Next slot
    
    
    Exit Function

HasAnyItem_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "HasAnyItem"
    End If
Resume Next
    
End Function

Public Sub HabilitarConfirmar(ByVal Habilitar As Boolean)
    
    On Error GoTo HabilitarConfirmar_Err
    
    Call cBotonConfirmar.EnableButton(Habilitar)

    
    Exit Sub

HabilitarConfirmar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "HabilitarConfirmar"
    End If
Resume Next
    
End Sub

Public Sub HabilitarAceptarRechazar(ByVal Habilitar As Boolean)
    
    On Error GoTo HabilitarAceptarRechazar_Err
    
    Call cBotonAceptar.EnableButton(Habilitar)
    Call cBotonRechazar.EnableButton(Habilitar)

    
    Exit Sub

HabilitarAceptarRechazar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmComerciarUsu" & "->" & "HabilitarAceptarRechazar"
    End If
Resume Next
    
End Sub
