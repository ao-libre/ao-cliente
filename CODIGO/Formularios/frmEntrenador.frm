VERSION 5.00
Begin VB.Form frmEntrenador 
   BorderStyle     =   0  'None
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4215
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
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   870
      TabIndex        =   0
      Top             =   675
      Width           =   2355
   End
   Begin AOLibre.uAOButton imgSalir 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton imgLuchar 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Luchar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Con que criatura deseas combatir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmEntrenador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matias Fernando Pequeno
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
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez



'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^,
'   le puse borde a la ventana y le cambie la letra a
'   una mas linda :)
'
'[END]'

Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
        
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaEntrenador.jpg")

    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)
End Sub

Private Sub LoadTextsForm()
    lblTitle.Caption = JsonLanguage.item("FRM_ENTRENADOR_TITLE").item("TEXTO")
    imgLuchar.Caption = JsonLanguage.item("FRM_ENTRENADOR_LUCHAR").item("TEXTO")
    imgSalir.Caption = JsonLanguage.item("FRM_ENTRENADOR_SALIR").item("TEXTO")
End Sub

Private Sub imgLuchar_Click()
    Call WriteTrain(lstCriaturas.ListIndex + 1)
    Unload Me
End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub
