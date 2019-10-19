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
   Begin AOLibre.uAOButton imgSalir 
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
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
   Begin VB.Label lblTitleCaos 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo Caos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label lblTitleCriminal 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo Criminal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblTitleNeutral 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo Neutral"
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
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblTitleLegal 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo Legal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblTitleReal 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo Real"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label imgCaos 
      BackStyle       =   0  'Transparent
      Caption         =   "Texto del caos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   4200
      Width           =   5655
   End
   Begin VB.Label imgCriminal 
      BackStyle       =   0  'Transparent
      Caption         =   "Texto del CRIMINAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   3240
      Width           =   5655
   End
   Begin VB.Label imgNeutral 
      BackStyle       =   0  'Transparent
      Caption         =   "Texto del Neutral"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   5655
   End
   Begin VB.Label imgLegal 
      BackStyle       =   0  'Transparent
      Caption         =   "Texto del Legal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label imgReal 
      BackStyle       =   0  'Transparent
      Caption         =   "Texto del Real"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
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

    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaFundarClan.jpg")
    
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)
End Sub

Private Sub LoadTextsForm()
    imgNeutral.Caption = JsonLanguage.item("FRM_ELIJEALINEACION_IMGNEUTRAL").item("TEXTO")
    imgLegal.Caption = JsonLanguage.item("FRM_ELIJEALINEACION_IMGLEGAL").item("TEXTO")
    imgCriminal.Caption = JsonLanguage.item("FRM_ELIJEALINEACION_IMGCRIMINAL").item("TEXTO")
    imgCaos.Caption = JsonLanguage.item("FRM_ELIJEALINEACION_IMGCAOS").item("TEXTO")
    imgReal.Caption = JsonLanguage.item("FRM_ELIJEALINEACION_IMGREAL").item("TEXTO")
    lblTitleNeutral.Caption = JsonLanguage.item("FRM_ELIJEALINEACION_IMGNEUTRAL_TITLE").item("TEXTO")
    lblTitleLegal.Caption = JsonLanguage.item("FRM_ELIJEALINEACION_IMGLEGAL_TITLE").item("TEXTO")
    lblTitleCriminal.Caption = JsonLanguage.item("FRM_ELIJEALINEACION_IMGCRIMINAL_TITLE").item("TEXTO")
    lblTitleCaos.Caption = JsonLanguage.item("FRM_ELIJEALINEACION_IMGCAOS_TITLE").item("TEXTO")
    lblTitleReal.Caption = JsonLanguage.item("FRM_ELIJEALINEACION_IMGREAL_TITLE").item("TEXTO")
    imgSalir.Caption = JsonLanguage.item("FRM_ELIJEALINEACION_IMGSALIR").item("TEXTO")
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
