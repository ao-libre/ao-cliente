VERSION 5.00
Begin VB.Form frmSkills3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSkills3.frx":0000
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgNavegacion 
      Height          =   375
      Left            =   4695
      Top             =   4110
      Width           =   1440
   End
   Begin VB.Image imgCombateSinArmas 
      Height          =   345
      Left            =   4695
      Top             =   3735
      Width           =   2100
   End
   Begin VB.Image imgCombateDistancia 
      Height          =   345
      Left            =   4695
      Top             =   3345
      Width           =   2280
   End
   Begin VB.Image imgDomar 
      Height          =   345
      Left            =   4695
      Top             =   2970
      Width           =   1845
   End
   Begin VB.Image imgLiderazgo 
      Height          =   330
      Left            =   4695
      Top             =   2580
      Width           =   1215
   End
   Begin VB.Image imgHerreria 
      Height          =   345
      Left            =   4695
      Top             =   2205
      Width           =   1065
   End
   Begin VB.Image imgCarpinteria 
      Height          =   360
      Left            =   4695
      Top             =   1830
      Width           =   1365
   End
   Begin VB.Image imgMineria 
      Height          =   360
      Left            =   4695
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Image imgPesca 
      Height          =   330
      Left            =   4695
      Top             =   1110
      Width           =   780
   End
   Begin VB.Image imgEscudos 
      Height          =   270
      Left            =   4695
      Top             =   720
      Width           =   2340
   End
   Begin VB.Image imgComercio 
      Height          =   330
      Left            =   495
      Top             =   4125
      Width           =   1170
   End
   Begin VB.Image imgTalar 
      Height          =   360
      Left            =   495
      Top             =   3750
      Width           =   885
   End
   Begin VB.Image imgSupervivencia 
      Height          =   330
      Left            =   495
      Top             =   3375
      Width           =   1620
   End
   Begin VB.Image imgOcultarse 
      Height          =   345
      Left            =   495
      Top             =   3030
      Width           =   1230
   End
   Begin VB.Image imgApunialar 
      Height          =   360
      Left            =   495
      Top             =   2640
      Width           =   1170
   End
   Begin VB.Image imgMeditar 
      Height          =   345
      Left            =   495
      Top             =   2265
      Width           =   1065
   End
   Begin VB.Image imgCombateArmas 
      Height          =   315
      Left            =   495
      Top             =   1890
      Width           =   2280
   End
   Begin VB.Image imgEvasion 
      Height          =   330
      Left            =   495
      Top             =   1515
      Width           =   2295
   End
   Begin VB.Image imgRobar 
      Height          =   360
      Left            =   495
      Top             =   1125
      Width           =   930
   End
   Begin VB.Image imgMagia 
      Height          =   330
      Left            =   495
      Top             =   750
      Width           =   870
   End
   Begin VB.Label lblHelp 
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
      Height          =   1215
      Left            =   600
      TabIndex        =   21
      Top             =   4710
      Width           =   7815
   End
   Begin VB.Image imgCancelar 
      Height          =   360
      Left            =   510
      Picture         =   "frmSkills3.frx":3DDCA
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   1
      Left            =   3495
      TabIndex        =   20
      Top             =   840
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   2
      Left            =   3495
      TabIndex        =   19
      Top             =   1215
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   3
      Left            =   3495
      TabIndex        =   18
      Top             =   1575
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   4
      Left            =   3495
      TabIndex        =   17
      Top             =   1950
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   5
      Left            =   3495
      TabIndex        =   16
      Top             =   2325
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   6
      Left            =   3495
      TabIndex        =   15
      Top             =   2700
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   7
      Left            =   3495
      TabIndex        =   14
      Top             =   3075
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   8
      Left            =   3495
      TabIndex        =   13
      Top             =   3450
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   9
      Left            =   3495
      TabIndex        =   12
      Top             =   3825
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   10
      Left            =   3495
      TabIndex        =   11
      Top             =   4200
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   11
      Left            =   7635
      TabIndex        =   10
      Top             =   840
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   12
      Left            =   7635
      TabIndex        =   9
      Top             =   1215
      Width           =   405
   End
   Begin VB.Image imgMas1 
      Height          =   300
      Left            =   3960
      Picture         =   "frmSkills3.frx":41CBB
      Top             =   780
      Width           =   300
   End
   Begin VB.Image imgMas2 
      Height          =   300
      Left            =   3960
      Top             =   1155
      Width           =   300
   End
   Begin VB.Image imgMenos2 
      Height          =   300
      Left            =   3120
      Top             =   1155
      Width           =   300
   End
   Begin VB.Image imgMas3 
      Height          =   300
      Left            =   3960
      Top             =   1515
      Width           =   300
   End
   Begin VB.Image imgMenos3 
      Height          =   300
      Left            =   3120
      Top             =   1515
      Width           =   300
   End
   Begin VB.Image imgMas4 
      Height          =   300
      Left            =   3960
      Top             =   1890
      Width           =   300
   End
   Begin VB.Image imgMenos4 
      Height          =   300
      Left            =   3120
      Top             =   1890
      Width           =   300
   End
   Begin VB.Image imgMas5 
      Height          =   300
      Left            =   3960
      Top             =   2265
      Width           =   300
   End
   Begin VB.Image imgMenos5 
      Height          =   300
      Left            =   3120
      Top             =   2265
      Width           =   300
   End
   Begin VB.Image imgMas6 
      Height          =   300
      Left            =   3960
      Top             =   2640
      Width           =   300
   End
   Begin VB.Image imgMenos6 
      Height          =   300
      Left            =   3120
      Top             =   2640
      Width           =   300
   End
   Begin VB.Image imgMas7 
      Height          =   300
      Left            =   3960
      Top             =   3015
      Width           =   300
   End
   Begin VB.Image imgMenos7 
      Height          =   300
      Left            =   3120
      Top             =   3015
      Width           =   300
   End
   Begin VB.Image imgMas8 
      Height          =   300
      Left            =   3960
      Top             =   3390
      Width           =   300
   End
   Begin VB.Image imgMenos8 
      Height          =   300
      Left            =   3120
      Top             =   3390
      Width           =   300
   End
   Begin VB.Image imgMas9 
      Height          =   300
      Left            =   3960
      Top             =   3765
      Width           =   300
   End
   Begin VB.Image imgMenos9 
      Height          =   300
      Left            =   3120
      Top             =   3765
      Width           =   300
   End
   Begin VB.Image imgMas10 
      Height          =   300
      Left            =   3960
      Top             =   4140
      Width           =   300
   End
   Begin VB.Image imgMenos10 
      Height          =   300
      Left            =   3120
      Top             =   4140
      Width           =   300
   End
   Begin VB.Image imgMas11 
      Height          =   300
      Left            =   8100
      Picture         =   "frmSkills3.frx":44CE9
      Top             =   780
      Width           =   300
   End
   Begin VB.Image imgMenos11 
      Height          =   300
      Left            =   7260
      Picture         =   "frmSkills3.frx":47D17
      Top             =   780
      Width           =   300
   End
   Begin VB.Image imgMas12 
      Height          =   285
      Left            =   8100
      Top             =   1155
      Width           =   345
   End
   Begin VB.Image imgMenos12 
      Height          =   285
      Left            =   7260
      Top             =   1155
      Width           =   345
   End
   Begin VB.Image imgMas13 
      Height          =   285
      Left            =   8100
      Top             =   1515
      Width           =   345
   End
   Begin VB.Image imgMenos13 
      Height          =   285
      Left            =   7260
      Top             =   1515
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   13
      Left            =   7635
      TabIndex        =   8
      Top             =   1575
      Width           =   405
   End
   Begin VB.Image imgMas14 
      Height          =   285
      Left            =   8100
      Top             =   1890
      Width           =   345
   End
   Begin VB.Image imgMenos14 
      Height          =   285
      Left            =   7260
      Top             =   1890
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   14
      Left            =   7635
      TabIndex        =   7
      Top             =   1950
      Width           =   405
   End
   Begin VB.Image imgMas15 
      Height          =   285
      Left            =   8100
      Top             =   2265
      Width           =   345
   End
   Begin VB.Image imgMenos15 
      Height          =   285
      Left            =   7260
      Top             =   2265
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   15
      Left            =   7635
      TabIndex        =   6
      Top             =   2325
      Width           =   405
   End
   Begin VB.Image imgMas16 
      Height          =   285
      Left            =   8100
      Top             =   2640
      Width           =   345
   End
   Begin VB.Image imgMenos16 
      Height          =   285
      Left            =   7260
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   16
      Left            =   7635
      TabIndex        =   5
      Top             =   2700
      Width           =   405
   End
   Begin VB.Image imgMas17 
      Height          =   285
      Left            =   8100
      Top             =   3015
      Width           =   345
   End
   Begin VB.Image imgMenos17 
      Height          =   285
      Left            =   7260
      Top             =   3015
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   17
      Left            =   7635
      TabIndex        =   4
      Top             =   3075
      Width           =   405
   End
   Begin VB.Image imgMas18 
      Height          =   285
      Left            =   8100
      Top             =   3390
      Width           =   345
   End
   Begin VB.Image imgMenos18 
      Height          =   285
      Left            =   7260
      Top             =   3390
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   18
      Left            =   7635
      TabIndex        =   3
      Top             =   3450
      Width           =   405
   End
   Begin VB.Image imgMenos1 
      Height          =   300
      Left            =   3120
      Picture         =   "frmSkills3.frx":4AD1C
      Top             =   780
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   19
      Left            =   7635
      TabIndex        =   2
      Top             =   3825
      Width           =   405
   End
   Begin VB.Image imgMas19 
      Height          =   285
      Left            =   8100
      Top             =   3765
      Width           =   345
   End
   Begin VB.Image imgMenos19 
      Height          =   285
      Left            =   7260
      Top             =   3765
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   20
      Left            =   7635
      TabIndex        =   1
      Top             =   4200
      Width           =   405
   End
   Begin VB.Image imgMas20 
      Height          =   285
      Left            =   8100
      Top             =   4140
      Width           =   345
   End
   Begin VB.Image imgMenos20 
      Height          =   285
      Left            =   7260
      Top             =   4140
      Width           =   345
   End
   Begin VB.Image imgAceptar 
      Height          =   360
      Left            =   6990
      Picture         =   "frmSkills3.frx":4DD21
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label puntos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4800
      TabIndex        =   0
      Top             =   180
      Width           =   90
   End
End
Attribute VB_Name = "frmSkills3"
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

Private clsFormulario               As clsFormMovementManager

Private cBotonMas(1 To NUMSKILLS)   As clsGraphicalButton
Private cBotonMenos(1 To NUMSKILLS) As clsGraphicalButton
Private cSkillNames(1 To NUMSKILLS) As clsGraphicalButton
Private cBtonAceptar                As clsGraphicalButton
Private cBotonCancelar              As clsGraphicalButton

Public LastButtonPressed            As clsGraphicalButton

Private bPuedeMagia                 As Boolean
Private bPuedeMeditar               As Boolean
Private bPuedeEscudo                As Boolean
Private bPuedeCombateDistancia      As Boolean

Private vsHelp(1 To NUMSKILLS)      As String

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    
    MirandoAsignarSkills = True
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    'Flags para saber que skills se modificaron
    ReDim Flags(1 To NUMSKILLS)
    
    Call ValidarSkills
    
    Me.Picture = LoadPicture(DirGraficos & "VentanaSkills.jpg")
    Call LoadButtons
    
    Call LoadHelp

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
    Dim GrhPath As String
    Dim i       As Long
    
    GrhPath = DirGraficos

    For i = 1 To NUMSKILLS
        Set cBotonMas(i) = New clsGraphicalButton
        Set cBotonMenos(i) = New clsGraphicalButton
        Set cSkillNames(i) = New clsGraphicalButton
    Next i
    
    Set cBtonAceptar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBtonAceptar.Initialize(imgAceptar, GrhPath & "BotonAceptarSkills.jpg", GrhPath & "BotonAceptarRolloverSkills.jpg", GrhPath & "BotonAceptarClickSkills.jpg", Me)

    Call cBotonCancelar.Initialize(imgCancelar, GrhPath & "BotonCacelarSkills.jpg", GrhPath & "BotonCacelarRolloverSkills.jpg", GrhPath & "BotonCacelarClickSkills.jpg", Me)

    Call cBotonMas(1).Initialize(imgMas1, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me, GrhPath & "BotonMasSkills.jpg", Not bPuedeMagia)

    Call cBotonMas(2).Initialize(imgMas2, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)

    Call cBotonMas(3).Initialize(imgMas3, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)

    Call cBotonMas(4).Initialize(imgMas4, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)
    
    Call cBotonMas(5).Initialize(imgMas5, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me, GrhPath & "BotonMasSkills.jpg", Not bPuedeMeditar)

    Call cBotonMas(6).Initialize(imgMas6, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)

    Call cBotonMas(7).Initialize(imgMas7, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)

    Call cBotonMas(8).Initialize(imgMas8, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)
    
    Call cBotonMas(9).Initialize(imgMas9, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)

    Call cBotonMas(10).Initialize(imgMas10, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)

    Call cBotonMas(11).Initialize(imgMas11, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me, GrhPath & "BotonMasSkills.jpg", Not bPuedeEscudo)

    Call cBotonMas(12).Initialize(imgMas12, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)
    
    Call cBotonMas(13).Initialize(imgMas13, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)

    Call cBotonMas(14).Initialize(imgMas14, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)

    Call cBotonMas(15).Initialize(imgMas15, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)

    Call cBotonMas(16).Initialize(imgMas16, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)
    
    Call cBotonMas(17).Initialize(imgMas17, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)

    Call cBotonMas(18).Initialize(imgMas18, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me, GrhPath & "BotonMasSkills.jpg", Not bPuedeCombateDistancia)

    Call cBotonMas(19).Initialize(imgMas19, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)

    Call cBotonMas(20).Initialize(imgMas20, GrhPath & "BotonMasSkills.jpg", GrhPath & "BotonMasRolloverSkills.jpg", GrhPath & "BotonMasClickSkills.jpg", Me)
    
    Call cBotonMenos(1).Initialize(imgMenos1, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me, GrhPath & "BotonMenosSkills.jpg", Not bPuedeMagia)

    Call cBotonMenos(2).Initialize(imgMenos2, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cBotonMenos(3).Initialize(imgMenos3, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cBotonMenos(4).Initialize(imgMenos4, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)
    
    Call cBotonMenos(5).Initialize(imgMenos5, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me, GrhPath & "BotonMenosSkills.jpg", Not bPuedeMeditar)

    Call cBotonMenos(6).Initialize(imgMenos6, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cBotonMenos(7).Initialize(imgMenos7, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cBotonMenos(8).Initialize(imgMenos8, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)
    
    Call cBotonMenos(9).Initialize(imgMenos9, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cBotonMenos(10).Initialize(imgMenos10, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cBotonMenos(11).Initialize(imgMenos11, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me, GrhPath & "BotonMenosSkills.jpg", Not bPuedeEscudo)

    Call cBotonMenos(12).Initialize(imgMenos12, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)
    
    Call cBotonMenos(13).Initialize(imgMenos13, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cBotonMenos(14).Initialize(imgMenos14, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cBotonMenos(15).Initialize(imgMenos15, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cBotonMenos(16).Initialize(imgMenos16, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)
    
    Call cBotonMenos(17).Initialize(imgMenos17, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cBotonMenos(18).Initialize(imgMenos18, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me, GrhPath & "BotonMenosSkills.jpg", Not bPuedeCombateDistancia)

    Call cBotonMenos(19).Initialize(imgMenos19, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cBotonMenos(20).Initialize(imgMenos20, GrhPath & "BotonMenosSkills.jpg", GrhPath & "BotonMenosRolloverSkills.jpg", GrhPath & "BotonMenosClickSkills.jpg", Me)

    Call cSkillNames(1).Initialize(imgMagia, "", GrhPath & "MagiaRollover.jpg", GrhPath & "MagiaRollover.jpg", Me, GrhPath & "MagiaBloqueado.jpg", Not bPuedeMagia, False, False)

    Call cSkillNames(2).Initialize(imgRobar, "", GrhPath & "RobarRollover.jpg", GrhPath & "RobarRollover.jpg", Me, , , False, False)

    Call cSkillNames(3).Initialize(imgEvasion, "", GrhPath & "EvasionRollover.jpg", GrhPath & "EvasionRollover.jpg", Me, , , False, False)
                                    
    Call cSkillNames(4).Initialize(imgCombateArmas, "", GrhPath & "CombateConArmasRollover.jpg", GrhPath & "CombateConArmasRollover.jpg", Me, , , False, False)
    
    Call cSkillNames(5).Initialize(imgMeditar, "", GrhPath & "MeditarRollover.jpg", GrhPath & "MeditarRollover.jpg", Me, GrhPath & "MeditarBloqueado.jpg", Not bPuedeMeditar, False, False)

    Call cSkillNames(6).Initialize(imgApunialar, "", GrhPath & "ApuñalarRollover.jpg", GrhPath & "ApuñalarRollover.jpg", Me, , , False, False)

    Call cSkillNames(7).Initialize(imgOcultarse, "", GrhPath & "OcultarseRollover.jpg", GrhPath & "OcultarseRollover.jpg", Me, , , False, False)

    Call cSkillNames(8).Initialize(imgSupervivencia, "", GrhPath & "SupervivenciaRollover.jpg", GrhPath & "SupervivenciaRollover.jpg", Me, , , False, False)
    
    Call cSkillNames(9).Initialize(imgTalar, "", GrhPath & "TalarRollover.jpg", GrhPath & "TalarRollover.jpg", Me, , , False, False)

    Call cSkillNames(10).Initialize(imgComercio, "", GrhPath & "ComercioRollover.jpg", GrhPath & "ComercioRollover.jpg", Me, , , False, False)

    Call cSkillNames(11).Initialize(imgEscudos, "", GrhPath & "DefensaConEscudosRollover.jpg", GrhPath & "DefensaConEscudosRollover.jpg", Me, GrhPath & "DefensaEscudosBloqueado.jpg", Not bPuedeEscudo, False, False)

    Call cSkillNames(12).Initialize(imgPesca, "", GrhPath & "PescaRollover.jpg", GrhPath & "PescaRollover.jpg", Me, , , False, False)

    Call cSkillNames(13).Initialize(imgMineria, "", GrhPath & "MineriaRollover.jpg", GrhPath & "MineriaRollover.jpg", Me, , , False, False)

    Call cSkillNames(14).Initialize(imgCarpinteria, "", GrhPath & "CarpinteriaRollover.jpg", GrhPath & "CarpinteriaRollover.jpg", Me, , , False, False)

    Call cSkillNames(15).Initialize(imgHerreria, "", GrhPath & "HerreriaRollover.jpg", GrhPath & "HerreriaRollover.jpg", Me, , , False, False)

    Call cSkillNames(16).Initialize(imgLiderazgo, "", GrhPath & "LiderazgoRollover.jpg", GrhPath & "LiderazgoRollover.jpg", Me, , , False, False)

    Call cSkillNames(17).Initialize(imgDomar, "", GrhPath & "DomarAnimalesRollover.jpg", GrhPath & "DomarAnimalesRollover.jpg", Me, , , False, False)

    Call cSkillNames(18).Initialize(imgCombateDistancia, "", GrhPath & "CombateADistanciaRollover.jpg", GrhPath & "CombateADistanciaRollover.jpg", Me, GrhPath & "CombateADistanciaBloqueado.jpg", Not bPuedeCombateDistancia, False, False)

    Call cSkillNames(19).Initialize(imgCombateSinArmas, "", GrhPath & "CombateSinArmasRollover.jpg", GrhPath & "CombateSinArmasRollover.jpg", Me, , , False, False)

    Call cSkillNames(20).Initialize(imgNavegacion, "", GrhPath & "NavegacionRollover.jpg", GrhPath & "NavegacionRollover.jpg", Me, , , False, False)

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub SumarSkillPoint(ByVal SkillIndex As Integer)
    
    On Error GoTo SumarSkillPoint_Err
    

    If Alocados > 0 Then

<<<<<<< HEAD
        If Val(text1(SkillIndex).Caption) < MAXSKILLPOINTS Then
            text1(SkillIndex).Caption = Val(text1(SkillIndex).Caption) + 1
=======
        If Val(Text1(SkillIndex).Caption) < MAXSKILLPOINTS Then
            Text1(SkillIndex).Caption = Val(Text1(SkillIndex).Caption) + 1
>>>>>>> origin/master
            Flags(SkillIndex) = Flags(SkillIndex) + 1
            Alocados = Alocados - 1

        End If
            
    End If
    
    puntos.Caption = Alocados

    
    Exit Sub

SumarSkillPoint_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "SumarSkillPoint"
    End If
Resume Next
    
End Sub

Private Sub RestarSkillPoint(ByVal SkillIndex As Integer)
    
    On Error GoTo RestarSkillPoint_Err
    

    If Alocados < SkillPoints Then
        
<<<<<<< HEAD
        If Val(text1(SkillIndex).Caption) > 0 And Flags(SkillIndex) > 0 Then
            text1(SkillIndex).Caption = Val(text1(SkillIndex).Caption) - 1
=======
        If Val(Text1(SkillIndex).Caption) > 0 And Flags(SkillIndex) > 0 Then
            Text1(SkillIndex).Caption = Val(Text1(SkillIndex).Caption) - 1
>>>>>>> origin/master
            Flags(SkillIndex) = Flags(SkillIndex) - 1
            Alocados = Alocados + 1

        End If

    End If
    
    puntos.Caption = Alocados

    
    Exit Sub

RestarSkillPoint_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "RestarSkillPoint"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal
    lblHelp.Caption = vbNullString

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Form_Unload_Err
    
    MirandoAsignarSkills = False

    
    Exit Sub

Form_Unload_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "Form_Unload"
    End If
Resume Next
    
End Sub

Private Sub imgAceptar_Click()
    
    On Error GoTo imgAceptar_Click_Err
    
    Dim skillChanges(NUMSKILLS) As Byte
    Dim i                       As Long

    For i = 1 To NUMSKILLS
        skillChanges(i) = CByte(Text1(i).Caption) - UserSkills(i)
        'Actualizamos nuestros datos locales
        UserSkills(i) = Val(Text1(i).Caption)
    Next i
    
    Call WriteModifySkills(skillChanges())
    
    If Alocados = 0 Then Call frmMain.LightSkillStar(False)
    
    SkillPoints = Alocados
    
    Unload Me

    
    Exit Sub

imgAceptar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgAceptar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgApunialar_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    
    On Error GoTo imgApunialar_MouseMove_Err
    
    Call ShowHelp(eSkill.Apuñalar)

    
    Exit Sub

imgApunialar_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgApunialar_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgCancelar_Click()
    
    On Error GoTo imgCancelar_Click_Err
    
    Unload Me

    
    Exit Sub

imgCancelar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgCancelar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgCarpinteria_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)
    
    On Error GoTo imgCarpinteria_MouseMove_Err
    
    Call ShowHelp(eSkill.Carpinteria)

    
    Exit Sub

imgCarpinteria_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgCarpinteria_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgCombateArmas_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    
    On Error GoTo imgCombateArmas_MouseMove_Err
    
    Call ShowHelp(eSkill.Armas)

    
    Exit Sub

imgCombateArmas_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgCombateArmas_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgCombateDistancia_MouseMove(Button As Integer, _
                                          Shift As Integer, _
                                          X As Single, _
                                          Y As Single)
    
    On Error GoTo imgCombateDistancia_MouseMove_Err
    
    Call ShowHelp(eSkill.Proyectiles)

    
    Exit Sub

imgCombateDistancia_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgCombateDistancia_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgCombateSinArmas_MouseMove(Button As Integer, _
                                         Shift As Integer, _
                                         X As Single, _
                                         Y As Single)
    
    On Error GoTo imgCombateSinArmas_MouseMove_Err
    
    Call ShowHelp(eSkill.Wrestling)

    
    Exit Sub

imgCombateSinArmas_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgCombateSinArmas_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgComercio_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    
    On Error GoTo imgComercio_MouseMove_Err
    
    Call ShowHelp(eSkill.Comerciar)

    
    Exit Sub

imgComercio_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgComercio_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgDomar_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    On Error GoTo imgDomar_MouseMove_Err
    
    Call ShowHelp(eSkill.Domar)

    
    Exit Sub

imgDomar_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgDomar_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo imgEscudos_MouseMove_Err
    
    Call ShowHelp(eSkill.Defensa)

    
    Exit Sub

imgEscudos_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgEscudos_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo imgEvasion_MouseMove_Err
    
    Call ShowHelp(eSkill.Tacticas)

    
    Exit Sub

imgEvasion_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgEvasion_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgHerreria_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    
    On Error GoTo imgHerreria_MouseMove_Err
    
    Call ShowHelp(eSkill.Herreria)

    
    Exit Sub

imgHerreria_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgHerreria_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgLiderazgo_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    
    On Error GoTo imgLiderazgo_MouseMove_Err
    
    Call ShowHelp(eSkill.Liderazgo)

    
    Exit Sub

imgLiderazgo_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgLiderazgo_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    On Error GoTo imgMagia_MouseMove_Err
    
    Call ShowHelp(eSkill.Magia)

    
    Exit Sub

imgMagia_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMagia_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgMas1_Click()
    
    On Error GoTo imgMas1_Click_Err
    
    Call SumarSkillPoint(1)

    
    Exit Sub

imgMas1_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas1_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas10_Click()
    
    On Error GoTo imgMas10_Click_Err
    
    Call SumarSkillPoint(10)

    
    Exit Sub

imgMas10_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas10_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas11_Click()
    
    On Error GoTo imgMas11_Click_Err
    
    Call SumarSkillPoint(11)

    
    Exit Sub

imgMas11_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas11_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas12_Click()
    
    On Error GoTo imgMas12_Click_Err
    
    Call SumarSkillPoint(12)

    
    Exit Sub

imgMas12_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas12_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas13_Click()
    
    On Error GoTo imgMas13_Click_Err
    
    Call SumarSkillPoint(13)

    
    Exit Sub

imgMas13_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas13_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas14_Click()
    
    On Error GoTo imgMas14_Click_Err
    
    Call SumarSkillPoint(14)

    
    Exit Sub

imgMas14_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas14_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas15_Click()
    
    On Error GoTo imgMas15_Click_Err
    
    Call SumarSkillPoint(15)

    
    Exit Sub

imgMas15_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas15_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas16_Click()
    
    On Error GoTo imgMas16_Click_Err
    
    Call SumarSkillPoint(16)

    
    Exit Sub

imgMas16_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas16_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas17_Click()
    
    On Error GoTo imgMas17_Click_Err
    
    Call SumarSkillPoint(17)

    
    Exit Sub

imgMas17_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas17_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas18_Click()
    
    On Error GoTo imgMas18_Click_Err
    
    Call SumarSkillPoint(18)

    
    Exit Sub

imgMas18_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas18_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas19_Click()
    
    On Error GoTo imgMas19_Click_Err
    
    Call SumarSkillPoint(19)

    
    Exit Sub

imgMas19_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas19_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas2_Click()
    
    On Error GoTo imgMas2_Click_Err
    
    Call SumarSkillPoint(2)

    
    Exit Sub

imgMas2_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas2_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas20_Click()
    
    On Error GoTo imgMas20_Click_Err
    
    Call SumarSkillPoint(20)

    
    Exit Sub

imgMas20_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas20_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas3_Click()
    
    On Error GoTo imgMas3_Click_Err
    
    Call SumarSkillPoint(3)

    
    Exit Sub

imgMas3_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas3_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas4_Click()
    
    On Error GoTo imgMas4_Click_Err
    
    Call SumarSkillPoint(4)

    
    Exit Sub

imgMas4_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas4_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas5_Click()
    
    On Error GoTo imgMas5_Click_Err
    
    Call SumarSkillPoint(5)

    
    Exit Sub

imgMas5_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas5_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas6_Click()
    
    On Error GoTo imgMas6_Click_Err
    
    Call SumarSkillPoint(6)

    
    Exit Sub

imgMas6_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas6_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas7_Click()
    
    On Error GoTo imgMas7_Click_Err
    
    Call SumarSkillPoint(7)

    
    Exit Sub

imgMas7_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas7_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas8_Click()
    
    On Error GoTo imgMas8_Click_Err
    
    Call SumarSkillPoint(8)

    
    Exit Sub

imgMas8_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas8_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMas9_Click()
    
    On Error GoTo imgMas9_Click_Err
    
    Call SumarSkillPoint(9)

    
    Exit Sub

imgMas9_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMas9_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMeditar_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo imgMeditar_MouseMove_Err
    
    Call ShowHelp(eSkill.Meditar)

    
    Exit Sub

imgMeditar_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMeditar_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgMenos1_Click()
    
    On Error GoTo imgMenos1_Click_Err
    
    Call RestarSkillPoint(1)

    
    Exit Sub

imgMenos1_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos1_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos10_Click()
    
    On Error GoTo imgMenos10_Click_Err
    
    Call RestarSkillPoint(10)

    
    Exit Sub

imgMenos10_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos10_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos11_Click()
    
    On Error GoTo imgMenos11_Click_Err
    
    Call RestarSkillPoint(11)

    
    Exit Sub

imgMenos11_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos11_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos12_Click()
    
    On Error GoTo imgMenos12_Click_Err
    
    Call RestarSkillPoint(12)

    
    Exit Sub

imgMenos12_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos12_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos13_Click()
    
    On Error GoTo imgMenos13_Click_Err
    
    Call RestarSkillPoint(13)

    
    Exit Sub

imgMenos13_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos13_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos14_Click()
    
    On Error GoTo imgMenos14_Click_Err
    
    Call RestarSkillPoint(14)

    
    Exit Sub

imgMenos14_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos14_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos15_Click()
    
    On Error GoTo imgMenos15_Click_Err
    
    Call RestarSkillPoint(15)

    
    Exit Sub

imgMenos15_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos15_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos16_Click()
    
    On Error GoTo imgMenos16_Click_Err
    
    Call RestarSkillPoint(16)

    
    Exit Sub

imgMenos16_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos16_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos17_Click()
    
    On Error GoTo imgMenos17_Click_Err
    
    Call RestarSkillPoint(17)

    
    Exit Sub

imgMenos17_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos17_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos18_Click()
    
    On Error GoTo imgMenos18_Click_Err
    
    Call RestarSkillPoint(18)

    
    Exit Sub

imgMenos18_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos18_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos19_Click()
    
    On Error GoTo imgMenos19_Click_Err
    
    Call RestarSkillPoint(19)

    
    Exit Sub

imgMenos19_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos19_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos2_Click()
    
    On Error GoTo imgMenos2_Click_Err
    
    Call RestarSkillPoint(2)

    
    Exit Sub

imgMenos2_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos2_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos20_Click()
    
    On Error GoTo imgMenos20_Click_Err
    
    Call RestarSkillPoint(20)

    
    Exit Sub

imgMenos20_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos20_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos3_Click()
    
    On Error GoTo imgMenos3_Click_Err
    
    Call RestarSkillPoint(3)

    
    Exit Sub

imgMenos3_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos3_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos4_Click()
    
    On Error GoTo imgMenos4_Click_Err
    
    Call RestarSkillPoint(4)

    
    Exit Sub

imgMenos4_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos4_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos5_Click()
    
    On Error GoTo imgMenos5_Click_Err
    
    Call RestarSkillPoint(5)

    
    Exit Sub

imgMenos5_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos5_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos6_Click()
    
    On Error GoTo imgMenos6_Click_Err
    
    Call RestarSkillPoint(6)

    
    Exit Sub

imgMenos6_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos6_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos7_Click()
    
    On Error GoTo imgMenos7_Click_Err
    
    Call RestarSkillPoint(7)

    
    Exit Sub

imgMenos7_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos7_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos8_Click()
    
    On Error GoTo imgMenos8_Click_Err
    
    Call RestarSkillPoint(8)

    
    Exit Sub

imgMenos8_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos8_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMenos9_Click()
    
    On Error GoTo imgMenos9_Click_Err
    
    Call RestarSkillPoint(9)

    
    Exit Sub

imgMenos9_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMenos9_Click"
    End If
Resume Next
    
End Sub

Private Sub LoadHelp()
    
<<<<<<< HEAD
    On Error GoTo LoadHelp_Err
    
    
    vsHelp(eSkill.Magia) = "Magia:" & vbCrLf & "- Representa la habilidad de un personaje de las áreas mágica." & vbCrLf & "- Indica la variedad de hechizos que es capaz de dominar el personaje."

    If Not bPuedeMagia Then
        vsHelp(eSkill.Magia) = vsHelp(eSkill.Magia) & vbCrLf & "* Habilidad inhabilitada para tu clase."

    End If
    
    vsHelp(eSkill.Robar) = "Robar:" & vbCrLf & "- Habilidades de hurto. Nunca por medio de la violencia." & vbCrLf & "- Indica la probabilidad de éxito del personaje al intentar apoderarse de oro de otro, en caso de ser Ladrón, tambien podrá apoderarse de items."
    
    vsHelp(eSkill.Tacticas) = "Evasión en Combate:" & vbCrLf & "- Representa la habilidad general para moverse en combate entre golpes enemigos sin morir o tropezar en el intento." & vbCrLf & "- Indica la posibilidad de evadir un golpe físico del personaje."
    
    vsHelp(eSkill.Armas) = "Combate con Armas:" & vbCrLf & "- Representa la habilidad del personaje para manejar armas de combate cuerpo a cuerpo." & vbCrLf & "- Indica la probabilidad de impactar al oponente con armas cuerpo a cuerpo."
    
    vsHelp(eSkill.Meditar) = "Meditar:" & vbCrLf & "- Representa la capacidad del personaje de concentrarse para abstrarse dentro de su mente, y así revitalizar su fuerza espiritual." & vbCrLf & "- Indica la velocidad a la que el personaje recupera maná (Clases mágicas)."
=======
    vsHelp(eSkill.Magia) = JsonLanguage.Item("HABILIDADES").Item("MAGIA").Item("TEXTO") & ":" & vbCrLf & _
                           JsonLanguage.Item("HABILIDADES").Item("MAGIA").Item("DESCRIPCION")
    If Not bPuedeMagia Then
        vsHelp(eSkill.Magia) = vsHelp(eSkill.Magia) & vbCrLf & _
                               JsonLanguage.Item("AYUDA_NO_USAR_HABILIDAD").Item("TEXTO")
    End If
    
    vsHelp(eSkill.Robar) = JsonLanguage.Item("HABILIDADES").Item("ROBAR").Item("TEXTO") & ":" & vbCrLf & _
                           JsonLanguage.Item("HABILIDADES").Item("ROBAR").Item("DESCRIPCION")
    
    vsHelp(eSkill.Tacticas) = JsonLanguage.Item("HABILIDADES").Item("EVASION_EN_COMBATE").Item("TEXTO") & ":" & vbCrLf & _
                              JsonLanguage.Item("HABILIDADES").Item("EVASION_EN_COMBATE").Item("DESCRIPCION")
    
    vsHelp(eSkill.Armas) = JsonLanguage.Item("HABILIDADES").Item("COMBATE_CON_ARMAS").Item("TEXTO") & ":" & vbCrLf & _
                           JsonLanguage.Item("HABILIDADES").Item("COMBATE_CON_ARMAS").Item("DESCRIPCION")
>>>>>>> origin/master
    
    vsHelp(eSkill.Meditar) = JsonLanguage.Item("HABILIDADES").Item("MEDITAR").Item("TEXTO") & ":" & vbCrLf & _
                             JsonLanguage.Item("HABILIDADES").Item("MEDITAR").Item("DESCRIPCION")
    If Not bPuedeMeditar Then
<<<<<<< HEAD
        vsHelp(eSkill.Meditar) = vsHelp(eSkill.Meditar) & vbCrLf & "* Habilidad inhabilitada para tu clase."

    End If

    vsHelp(eSkill.Apuñalar) = "Apuñalar:" & vbCrLf & "- Representa la destreza para inflingir daño grave con armas cortas." & vbCrLf & "- Indica la posibilidad de apuñalar al enemigo en un ataque. El Asesino es la única clase que no necesitará 10 skills para comenzar a entrenar esta habilidad."

    vsHelp(eSkill.Ocultarse) = "Ocultarse:" & vbCrLf & "- La habilidad propia de un personaje para mimetizarse con el medio y evitar se perciba su presencia." & vbCrLf & "- Indica la facilidad con la que uno puede desaparecer de la vista de los demás y por cuanto tiempo."
    
    vsHelp(eSkill.Supervivencia) = "Superivencia:" & vbCrLf & "- Es el conjunto de habilidades necesarias para sobrevivir fuera de una ciudad en base a lo que la naturaleza ofrece." & vbCrLf & "- Permite conocer la salud de las criaturas guiándose exclusivamente por su aspecto, así como encender fogatas junto a las que descansar."
    
    vsHelp(eSkill.Talar) = "Talar:" & vbCrLf & "- Es la habilidad en el uso del hacha para evitar desperdiciar leña y maximizar la efectividad de cada golpe dado." & vbCrLf & "- Indica la probabilidad de obtener leña por golpe."
    
    vsHelp(eSkill.Comerciar) = "Comercio:" & vbCrLf & "- Es la habilidad para regatear los precios exigidos en la compra y evitar ser regateado al vender." & vbCrLf & "- Indica que tan caro se compra en el comercio con NPCs."
    
    vsHelp(eSkill.Defensa) = "Defensa con Escudos:" & vbCrLf & "- Es la habilidad de interponer correctamente el escudo ante cada embate enemigo para evitar ser impactado sin perder el equilibrio y poder responder rápidamente con la otra mano." & vbCrLf & "- Indica las probabilidades de bloquear un impacto con el escudo."
    
    If Not bPuedeEscudo Then
        vsHelp(eSkill.Defensa) = vsHelp(eSkill.Defensa) & vbCrLf & "* Habilidad inhabilitada para tu clase."
=======
        vsHelp(eSkill.Meditar) = vsHelp(eSkill.Meditar) & vbCrLf & _
                                JsonLanguage.Item("AYUDA_NO_USAR_HABILIDAD").Item("TEXTO")
    End If

    vsHelp(eSkill.Apuñalar) = JsonLanguage.Item("HABILIDADES").Item("APUNALAR").Item("TEXTO") & ":" & vbCrLf & _
                              JsonLanguage.Item("HABILIDADES").Item("APUNALAR").Item("DESCRIPCION")

    vsHelp(eSkill.Ocultarse) = JsonLanguage.Item("HABILIDADES").Item("OCULTARSE").Item("TEXTO") & ":" & vbCrLf & _
                                JsonLanguage.Item("HABILIDADES").Item("OCULTARSE").Item("DESCRIPCION")
    
    vsHelp(eSkill.Supervivencia) = JsonLanguage.Item("HABILIDADES").Item("SUPERVIVENCIA").Item("TEXTO") & ":" & vbCrLf & _
                                    JsonLanguage.Item("HABILIDADES").Item("SUPERVIVENCIA").Item("DESCRIPCION")
    
    vsHelp(eSkill.Talar) = JsonLanguage.Item("HABILIDADES").Item("TALAR").Item("TEXTO") & ":" & vbCrLf & _
                            JsonLanguage.Item("HABILIDADES").Item("TALAR").Item("DESCRIPCION")
    
    vsHelp(eSkill.Comerciar) = JsonLanguage.Item("HABILIDADES").Item("COMERCIO").Item("TEXTO") & ":" & vbCrLf & _
                                JsonLanguage.Item("HABILIDADES").Item("COMERCIO").Item("DESCRIPCION")
    
    vsHelp(eSkill.Defensa) = JsonLanguage.Item("HABILIDADES").Item("DEFENSA_CON_ESCUDOS").Item("TEXTO") & ":" & vbCrLf & _
                              JsonLanguage.Item("HABILIDADES").Item("DEFENSA_CON_ESCUDOS").Item("DESCRIPCION")
    
    If Not bPuedeEscudo Then
        vsHelp(eSkill.Defensa) = vsHelp(eSkill.Defensa) & vbCrLf & _
                                JsonLanguage.Item("AYUDA_NO_USAR_HABILIDAD").Item("TEXTO")
    End If
>>>>>>> origin/master

    End If

<<<<<<< HEAD
    vsHelp(eSkill.Pesca) = "Pesca:" & vbCrLf & "- Es el conjunto de conocimientos básicos para poder armar un señuelo, poner la carnada en el anzuelo y saber dónde buscar peces." & vbCrLf & "- Indica la probabilidad de tener éxito en cada intento de pescar."
    
    vsHelp(eSkill.Mineria) = "Minería:" & vbCrLf & "- Es el conjunto de conocimientos sobre los distintos minerales, el dónde se obtienen, cómo deben ser extraídos y trabajados." & vbCrLf & "- Indica la probabilidad de tener éxito en cada intento de minar y la capacidad, o no de convertir estos minerales en lingotes."
    
    vsHelp(eSkill.Carpinteria) = "Carpintería:" & vbCrLf & "- Es el conjunto de conocimientos para saber serruchar, lijar, encolar y clavar madera con un buen nivel de terminación." & vbCrLf & "- Indica la habilidad en el manejo de estas herramientas, el que tan bueno se es en el oficio de carpintero."
    
    vsHelp(eSkill.Herreria) = "Herrería:" & vbCrLf & "- Es el conjunto de conocimientos para saber procesar cada tipo de mineral para fundirlo, forjarlo y crear aleaciones." & vbCrLf & "- Indica la habilidad en el manejo de estas técnicas, el que tan bueno se es en el oficio de herrero."
    
    vsHelp(eSkill.Liderazgo) = "Liderazgo:" & vbCrLf & "- Es la habilidad propia del personaje para convencer a otros a seguirlo en batalla." & vbCrLf & "- Permite crear clanes y partys"
    
    vsHelp(eSkill.Domar) = "Domar Animales:" & vbCrLf & "- Es la habilidad en el trato con animales para que estos te sigan y ayuden en combate." & vbCrLf & "- Indica la posibilidad de lograr domar a una criatura y qué clases de criaturas se puede domar."
    
    vsHelp(eSkill.Proyectiles) = "Combate a distancia:" & vbCrLf & "- Es el manejo de las armas de largo alcance." & vbCrLf & "- Indica la probabilidad de éxito para impactar a un enemigo con este tipo de armas."
    
    If Not bPuedeCombateDistancia Then
        vsHelp(eSkill.Proyectiles) = vsHelp(eSkill.Proyectiles) & vbCrLf & "* Habilidad inhabilitada para tu clase."

    End If

    vsHelp(eSkill.Wrestling) = "Combate sin armas:" & vbCrLf & "- Es la habilidad del personaje para entrar en combate sin arma alguna salvo sus propios brazos." & vbCrLf & "- Indica la probabilidad de éxito para impactar a un enemigo estando desarmado. El Bandido y Ladrón tienen habilidades extras asociadas a esta habilidad."
    
    vsHelp(eSkill.Navegacion) = "Navegación:" & vbCrLf & "- Es la habilidad para controlar barcos en el mar sin naufragar." & vbCrLf & "- Indica que clase de barcos se pueden utilizar."
    
    
    Exit Sub

LoadHelp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "LoadHelp"
    End If
Resume Next
=======
    vsHelp(eSkill.Pesca) = JsonLanguage.Item("HABILIDADES").Item("PESCA").Item("TEXTO") & ":" & vbCrLf & _
                            JsonLanguage.Item("HABILIDADES").Item("PESCA").Item("DESCRIPCION")
    
    vsHelp(eSkill.Mineria) = JsonLanguage.Item("HABILIDADES").Item("MINERIA").Item("TEXTO") & ":" & vbCrLf & _
                              JsonLanguage.Item("HABILIDADES").Item("MINERIA").Item("DESCRIPCION")
    
    vsHelp(eSkill.Carpinteria) = JsonLanguage.Item("HABILIDADES").Item("CARPINTERIA").Item("TEXTO") & ":" & vbCrLf & _
                                  JsonLanguage.Item("HABILIDADES").Item("CARPINTERIA").Item("DESCRIPCION")
    
    vsHelp(eSkill.Herreria) = JsonLanguage.Item("HABILIDADES").Item("HERRERIA").Item("TEXTO") & ":" & vbCrLf & _
                                JsonLanguage.Item("HABILIDADES").Item("HERRERIA").Item("DESCRIPCION")
    
    vsHelp(eSkill.Liderazgo) = JsonLanguage.Item("HABILIDADES").Item("LIDERAZGO").Item("TEXTO") & ":" & vbCrLf & _
                                JsonLanguage.Item("HABILIDADES").Item("LIDERAZGO").Item("DESCRIPCION")
    
    vsHelp(eSkill.Domar) = JsonLanguage.Item("HABILIDADES").Item("DOMAR_ANIMALES").Item("TEXTO") & ":" & vbCrLf & _
                                JsonLanguage.Item("HABILIDADES").Item("DOMAR_ANIMALES").Item("DESCRIPCION")
    
    vsHelp(eSkill.Proyectiles) = JsonLanguage.Item("HABILIDADES").Item("COMBATE_A_DISTANCIA").Item("TEXTO") & ":" & vbCrLf & _
                                 JsonLanguage.Item("HABILIDADES").Item("COMBATE_A_DISTANCIA").Item("DESCRIPCION")
    
    If Not bPuedeCombateDistancia Then
        vsHelp(eSkill.Proyectiles) = vsHelp(eSkill.Proyectiles) & vbCrLf & _
                                JsonLanguage.Item("AYUDA_NO_USAR_HABILIDAD").Item("TEXTO")
    End If

    vsHelp(eSkill.Wrestling) = JsonLanguage.Item("HABILIDADES").Item("COMBATE_CUERPO_A_CUERPO").Item("TEXTO") & ":" & vbCrLf & _
                                JsonLanguage.Item("HABILIDADES").Item("COMBATE_CUERPO_A_CUERPO").Item("DESCRIPCION")
    
    vsHelp(eSkill.Navegacion) = JsonLanguage.Item("HABILIDADES").Item("NAVEGACION").Item("TEXTO") & ":" & vbCrLf & _
                                JsonLanguage.Item("HABILIDADES").Item("NAVEGACION").Item("DESCRIPCION")
>>>>>>> origin/master
    
End Sub

Private Sub imgMineria_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo imgMineria_MouseMove_Err
    
    Call ShowHelp(eSkill.Mineria)

    
    Exit Sub

imgMineria_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgMineria_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgNavegacion_MouseMove(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)
    
    On Error GoTo imgNavegacion_MouseMove_Err
    
    Call ShowHelp(eSkill.Navegacion)

    
    Exit Sub

imgNavegacion_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgNavegacion_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgOcultarse_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    
    On Error GoTo imgOcultarse_MouseMove_Err
    
    Call ShowHelp(eSkill.Ocultarse)

    
    Exit Sub

imgOcultarse_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgOcultarse_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgPesca_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    On Error GoTo imgPesca_MouseMove_Err
    
    Call ShowHelp(eSkill.Pesca)

    
    Exit Sub

imgPesca_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgPesca_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgRobar_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    On Error GoTo imgRobar_MouseMove_Err
    
    Call ShowHelp(eSkill.Robar)

    
    Exit Sub

imgRobar_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgRobar_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgSupervivencia_MouseMove(Button As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    
    On Error GoTo imgSupervivencia_MouseMove_Err
    
    Call ShowHelp(eSkill.Supervivencia)

    
    Exit Sub

imgSupervivencia_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgSupervivencia_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgTalar_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    On Error GoTo imgTalar_MouseMove_Err
    
    Call ShowHelp(eSkill.Talar)

    
    Exit Sub

imgTalar_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "imgTalar_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    
    On Error GoTo lblHelp_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

lblHelp_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "lblHelp_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub ShowHelp(ByVal eeSkill As eSkill)
    
    On Error GoTo ShowHelp_Err
    
    lblHelp.Caption = vsHelp(eeSkill)

    
    Exit Sub

ShowHelp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "ShowHelp"
    End If
Resume Next
    
End Sub

Private Sub ValidarSkills()
    
    On Error GoTo ValidarSkills_Err
    

    bPuedeMagia = True
    bPuedeMeditar = True
    bPuedeEscudo = True
    bPuedeCombateDistancia = True

    Select Case UserClase

        Case eClass.Warrior, eClass.Hunter, eClass.Worker, eClass.Thief
            bPuedeMagia = False
            bPuedeMeditar = False
        
        Case eClass.Pirat
            bPuedeMagia = False
            bPuedeMeditar = False
            bPuedeEscudo = False
        
        Case eClass.Mage, eClass.Druid
            bPuedeEscudo = False
            bPuedeCombateDistancia = False
            
    End Select
    
    ' Magia
    imgMas1.Enabled = bPuedeMagia
    imgMenos1.Enabled = bPuedeMagia

    ' Meditar
    imgMas5.Enabled = bPuedeMeditar
    imgMenos5.Enabled = bPuedeMeditar

    ' Escudos
    imgMas11.Enabled = bPuedeEscudo
    imgMenos11.Enabled = bPuedeEscudo

    ' Proyectiles
    imgMas18.Enabled = bPuedeCombateDistancia
    imgMenos18.Enabled = bPuedeCombateDistancia

    
    Exit Sub

ValidarSkills_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "ValidarSkills"
    End If
Resume Next
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyUp_Err
    
    If KeyCode = vbKeyEscape Then
        Set clsFormulario = Nothing
        Unload Me
    End If

    Exit Sub

Form_KeyUp_Err:

    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmSkills3" & "->" & "Form_KeyUp"

    End If

    Resume Next
    
End Sub

