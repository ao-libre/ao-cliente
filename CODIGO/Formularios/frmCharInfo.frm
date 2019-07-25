VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   0  'None
   Caption         =   "Informacion del personaje"
   ClientHeight    =   6585
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6390
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
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPeticiones 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   1080
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3210
      Width           =   5730
   End
   Begin VB.TextBox txtMiembro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   1080
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4695
      Width           =   5730
   End
   Begin AOLibre.uAOButton imgCerrar 
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCharInfo.frx":0000
      PICF            =   "frmCharInfo.frx":0A2A
      PICH            =   "frmCharInfo.frx":16EC
      PICV            =   "frmCharInfo.frx":267E
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
   Begin AOLibre.uAOButton imgEchar 
      Height          =   495
      Left            =   1320
      TabIndex        =   16
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      TX              =   "Echar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCharInfo.frx":3580
      PICF            =   "frmCharInfo.frx":3FAA
      PICH            =   "frmCharInfo.frx":4C6C
      PICV            =   "frmCharInfo.frx":5BFE
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
   Begin AOLibre.uAOButton imgPeticion 
      Height          =   495
      Left            =   2640
      TabIndex        =   17
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      TX              =   "Peticion"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCharInfo.frx":6B00
      PICF            =   "frmCharInfo.frx":752A
      PICH            =   "frmCharInfo.frx":81EC
      PICV            =   "frmCharInfo.frx":917E
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
      Height          =   495
      Left            =   3960
      TabIndex        =   18
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      TX              =   "Rechazar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCharInfo.frx":A080
      PICF            =   "frmCharInfo.frx":AAAA
      PICH            =   "frmCharInfo.frx":B76C
      PICV            =   "frmCharInfo.frx":C6FE
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
      Height          =   495
      Left            =   5160
      TabIndex        =   19
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      TX              =   "Aceptar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCharInfo.frx":D600
      PICF            =   "frmCharInfo.frx":E02A
      PICH            =   "frmCharInfo.frx":ECEC
      PICV            =   "frmCharInfo.frx":FC7E
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
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label Nombre 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   700
      Width           =   1440
   End
   Begin VB.Label Nivel 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   1750
      Width           =   1185
   End
   Begin VB.Label Clase 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1225
      Width           =   1575
   End
   Begin VB.Label Raza 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   960
      Width           =   1560
   End
   Begin VB.Label Genero 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   2010
      Width           =   1365
   End
   Begin VB.Label Banco 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   2250
      Width           =   1425
   End
   Begin VB.Label guildactual 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   960
      Width           =   2265
   End
   Begin VB.Label ejercito 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   1230
      Width           =   1785
   End
   Begin VB.Label Ciudadanos 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   4905
      TabIndex        =   4
      Top             =   1500
      Width           =   1185
   End
   Begin VB.Label criminales 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   1770
      Width           =   1185
   End
   Begin VB.Label reputacion 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   4905
      TabIndex        =   2
      Top             =   2040
      Width           =   1185
   End
End
Attribute VB_Name = "frmCharInfo"
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

Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton
Private cBotonPeticion As clsGraphicalButton
Private cBotonRechazar As clsGraphicalButton
Private cBotonEchar As clsGraphicalButton
Private cBotonAceptar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Public Enum CharInfoFrmType
    frmMembers
    frmMembershipRequests
End Enum

Public frmType As CharInfoFrmType

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(App.path & "\Graficos\VentanaInfoPj.jpg")
    
    Me.imgCerrar.Caption = JsonLanguage.Item("FRMCHARINFO_CERRAR").Item("TEXTO")
    Me.imgEchar.Caption = JsonLanguage.Item("FRMCHARINFO_ECHAR").Item("TEXTO")
    Me.imgPeticion.Caption = JsonLanguage.Item("FRMCHARINFO_PETICION").Item("TEXTO")
    Me.imgRechazar.Caption = JsonLanguage.Item("FRMCHARINFO_RECHAZAR").Item("TEXTO")
    Me.imgAceptar.Caption = JsonLanguage.Item("FRMCHARINFO_ACEPTAR").Item("TEXTO")
End Sub

Private Sub imgAceptar_Click()
    Call WriteGuildAcceptNewMember(Nombre)
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgEchar_Click()
    Call WriteGuildKickMember(Nombre)
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me
End Sub

Private Sub imgPeticion_Click()
    Call WriteGuildRequestJoinerInfo(Nombre)
End Sub

Private Sub imgRechazar_Click()
    frmCommet.T = RECHAZOPJ
    frmCommet.Nombre = Nombre.Caption
    frmCommet.Show vbModeless, frmCharInfo
End Sub

Private Sub txtMiembro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

