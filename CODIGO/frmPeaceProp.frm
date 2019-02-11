VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   0  'None
   Caption         =   "Ofertas de paz"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   1785
      ItemData        =   "frmPeaceProp.frx":0000
      Left            =   240
      List            =   "frmPeaceProp.frx":0002
      TabIndex        =   0
      Top             =   600
      Width           =   4620
   End
   Begin VB.Image imgRechazar 
      Height          =   480
      Left            =   3840
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgAceptar 
      Height          =   480
      Left            =   2640
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgDetalle 
      Height          =   480
      Left            =   1440
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgCerrar 
      Height          =   480
      Left            =   240
      Top             =   2520
      Width           =   960
   End
End
Attribute VB_Name = "frmPeaceProp"
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

Private clsFormulario    As clsFormMovementManager

Private cBotonAceptar    As clsGraphicalButton
Private cBotonCerrar     As clsGraphicalButton
Private cBotonDetalles   As clsGraphicalButton
Private cBotonRechazar   As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private TipoProp         As TIPO_PROPUESTA

Public Enum TIPO_PROPUESTA

    ALIANZA = 1
    PAZ = 2

End Enum

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    
    On Error GoTo Form_Load_Err
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Call LoadBackGround
    Call LoadButtons

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPeaceProp" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonDetalles = New clsGraphicalButton
    Set cBotonRechazar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotonAceptarOferta.jpg", GrhPath & "BotonAceptarRolloverOferta.jpg", GrhPath & "BotonAceptarClickOferta.jpg", Me)

    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarOferta.jpg", GrhPath & "BotonCerrarRolloverOferta.jpg", GrhPath & "BotonCerrarClickOferta.jpg", Me)

    Call cBotonDetalles.Initialize(imgDetalle, GrhPath & "BotonDetallesOferta.jpg", GrhPath & "BotonDetallesRolloverOferta.jpg", GrhPath & "BotonDetallesClickOferta.jpg", Me)

    Call cBotonRechazar.Initialize(imgRechazar, GrhPath & "BotonRechazarOferta.jpg", GrhPath & "BotonRechazarRolloverOferta.jpg", GrhPath & "BotonRechazarClickOferta.jpg", Me)

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPeaceProp" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub LoadBackGround()
    
    On Error GoTo LoadBackGround_Err
    

    If TipoProp = TIPO_PROPUESTA.ALIANZA Then
        Me.Picture = LoadPicture(DirGraficos & "VentanaOfertaAlianza.jpg")
    Else
        Me.Picture = LoadPicture(DirGraficos & "VentanaOfertaPaz.jpg")

    End If

    
    Exit Sub

LoadBackGround_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPeaceProp" & "->" & "LoadBackGround"
    End If
Resume Next
    
End Sub

Public Property Let ProposalType(ByVal nValue As TIPO_PROPUESTA)
    
    On Error GoTo ProposalType_Err
    
    TipoProp = nValue

    
    Exit Property

ProposalType_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPeaceProp" & "->" & "ProposalType"
    End If
Resume Next
    
End Property

Private Sub imgAceptar_Click()
    
    On Error GoTo imgAceptar_Click_Err
    

    If TipoProp = PAZ Then
        Call WriteGuildAcceptPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildAcceptAlliance(lista.List(lista.ListIndex))

    End If
    
    Me.Hide
    
    Unload Me

    
    Exit Sub

imgAceptar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPeaceProp" & "->" & "imgAceptar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgCerrar_Click()
    
    On Error GoTo imgCerrar_Click_Err
    
    Unload Me

    
    Exit Sub

imgCerrar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPeaceProp" & "->" & "imgCerrar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgDetalle_Click()
    
    On Error GoTo imgDetalle_Click_Err
    

    If TipoProp = PAZ Then
        Call WriteGuildPeaceDetails(lista.List(lista.ListIndex))
    Else
        Call WriteGuildAllianceDetails(lista.List(lista.ListIndex))

    End If

    
    Exit Sub

imgDetalle_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPeaceProp" & "->" & "imgDetalle_Click"
    End If
Resume Next
    
End Sub

Private Sub imgRechazar_Click()
    
    On Error GoTo imgRechazar_Click_Err
    

    If TipoProp = PAZ Then
        Call WriteGuildRejectPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildRejectAlliance(lista.List(lista.ListIndex))

    End If
    
    Me.Hide
    
    Unload Me

    
    Exit Sub

imgRechazar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPeaceProp" & "->" & "imgRechazar_Click"
    End If
Resume Next
    
End Sub

