VERSION 5.00
Begin VB.Form frmEligeAlineacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5370
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEligeAlineacion.frx":0000
   ScaleHeight     =   5370
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin AOLibre.uAOButton imgSalir 
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmEligeAlineacion.frx":2BF3D
      PICF            =   "frmEligeAlineacion.frx":2C967
      PICH            =   "frmEligeAlineacion.frx":2D629
      PICV            =   "frmEligeAlineacion.frx":2E5BB
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label imgCaos 
      Caption         =   $"frmEligeAlineacion.frx":2F4BD
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   4200
      Width           =   5655
   End
   Begin VB.Label imgCriminal 
      Caption         =   $"frmEligeAlineacion.frx":2F590
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   3240
      Width           =   5655
   End
   Begin VB.Label imgNeutral 
      Caption         =   $"frmEligeAlineacion.frx":2F66C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   5655
   End
   Begin VB.Label imgLegal 
      Caption         =   $"frmEligeAlineacion.frx":2F718
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label imgReal 
      Caption         =   $"frmEligeAlineacion.frx":2F7E0
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   5655
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

Public LastButtonPressed As clsGraphicalButton

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
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
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
    LastButtonPressed.ToggleToNormal
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
