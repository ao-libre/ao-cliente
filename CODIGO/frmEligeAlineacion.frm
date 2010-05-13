VERSION 5.00
Begin VB.Form frmEligeAlineacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5265
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgReal 
      Height          =   765
      Left            =   795
      Tag             =   "1"
      Top             =   300
      Width           =   5745
   End
   Begin VB.Image imgNeutral 
      Height          =   570
      Left            =   810
      Tag             =   "1"
      Top             =   2220
      Width           =   5730
   End
   Begin VB.Image imgLegal 
      Height          =   705
      Left            =   810
      Tag             =   "1"
      Top             =   1320
      Width           =   5715
   End
   Begin VB.Image imgCaos 
      Height          =   675
      Left            =   822
      Tag             =   "1"
      Top             =   4117
      Width           =   5700
   End
   Begin VB.Image imgCriminal 
      Height          =   705
      Left            =   818
      Tag             =   "1"
      Top             =   3150
      Width           =   5865
   End
   Begin VB.Image imgSalir 
      Height          =   315
      Left            =   5520
      Tag             =   "1"
      Top             =   4800
      Width           =   930
   End
End
Attribute VB_Name = "frmEligeAlineacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmEligeAlineacion.frm
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

Private cBotonCriminal As clsGraphicalButton
Private cBotonCaos As clsGraphicalButton
Private cBotonLegal As clsGraphicalButton
Private cBotonNeutral As clsGraphicalButton
Private cBotonReal As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Enum eAlineacion
    ieREAL = 0
    ieCAOS = 1
    ieNeutral = 2
    ieLegal = 4
    ieCriminal = 5
End Enum

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(App.path & "\graficos\VentanaFundarClan.jpg")
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonCriminal = New clsGraphicalButton
    Set cBotonCaos = New clsGraphicalButton
    Set cBotonLegal = New clsGraphicalButton
    Set cBotonNeutral = New clsGraphicalButton
    Set cBotonReal = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonCriminal.Initialize(imgCriminal, "", _
                                    GrhPath & "BotonCriminal.jpg", _
                                    GrhPath & "BotonCriminal.jpg", Me)

    Call cBotonCaos.Initialize(imgCaos, "", _
                                    GrhPath & "BotonCaos.jpg", _
                                    GrhPath & "BotonCaos.jpg", Me)

    Call cBotonLegal.Initialize(imgLegal, "", _
                                    GrhPath & "BotonLegal.jpg", _
                                    GrhPath & "BotonLegal.jpg", Me)

    Call cBotonNeutral.Initialize(imgNeutral, "", _
                                    GrhPath & "BotonNeutral.jpg", _
                                    GrhPath & "BotonNeutral.jpg", Me)

    Call cBotonReal.Initialize(imgReal, "", _
                                    GrhPath & "BotonReal.jpg", _
                                    GrhPath & "BotonReal.jpg", Me)
                                    
    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonSalirAlineacion.jpg", _
                                    GrhPath & "BotonSalirRolloverAlineacion.jpg", _
                                    GrhPath & "BotonSalirClickAlineacion.jpg", Me)


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgCaos_Click()
    Call WriteGuildFundation(eAlineacion.ieCAOS)
    Unload Me
End Sub

Private Sub imgCriminal_Click()
    Call WriteGuildFundation(eAlineacion.ieCriminal)
    Unload Me
End Sub

Private Sub imgLegal_Click()
    Call WriteGuildFundation(eAlineacion.ieLegal)
    Unload Me
End Sub

Private Sub imgNeutral_Click()
    Call WriteGuildFundation(eAlineacion.ieNeutral)
    Unload Me
End Sub

Private Sub imgReal_Click()
    Call WriteGuildFundation(eAlineacion.ieREAL)
    Unload Me
End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub
