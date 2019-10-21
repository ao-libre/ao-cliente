VERSION 5.00
Begin VB.Form frmCommet 
   BorderStyle     =   0  'None
   Caption         =   "Oferta de paz o alianza"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5055
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
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
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
      Height          =   1935
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin AOLibre.uAOButton imgEnviar 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      TX              =   "Enviar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
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
   Begin AOLibre.uAOButton imgCerrar 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
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
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo del form ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmCommet"
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
Private Const MAX_PROPOSAL_LENGTH As Integer = 520

Public Nombre As String

Public T As TIPO

Public Enum TIPO
    ALIANZA = 1
    PAZ = 2
    RECHAZOPJ = 3
End Enum

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaCommet.jpg")
    
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)
End Sub

Private Sub LoadTextsForm()
    imgEnviar.Caption = JsonLanguage.item("FRM_COMMET_ENVIAR").item("TEXTO")
    imgCerrar.Caption = JsonLanguage.item("FRM_COMMET_CERRAR").item("TEXTO")
    
    ' Depende del tipo de form que abramos cargamos un diferente titulo
    Select Case T
        Case TIPO.ALIANZA
            lblTitle.Caption = JsonLanguage.item("FRM_COMMET_ALIANZA").item("TEXTO")
            
        Case TIPO.PAZ
            lblTitle.Caption = JsonLanguage.item("FRM_COMMET_PAZ").item("TEXTO")
            
        Case TIPO.RECHAZOPJ
            lblTitle.Caption = JsonLanguage.item("FRM_COMMET_RECHAZOPJ").item("TEXTO")
            
    End Select
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgEnviar_Click()

    If LenB(Text1) = 0 Then
        If T = PAZ Or T = ALIANZA Then
            MsgBox "Debes redactar un mensaje solicitando la paz o alianza al lider de " & Nombre
        Else
            MsgBox "Debes indicar el motivo por el cual rechazas la membresia de " & Nombre
        End If
        
        Exit Sub
    End If
    
    If T = PAZ Then
        Call WriteGuildOfferPeace(Nombre, Replace(Text1, vbNewLine, "º"))
        
    ElseIf T = ALIANZA Then
        Call WriteGuildOfferAlliance(Nombre, Replace(Text1, vbNewLine, "º"))
        
    ElseIf T = RECHAZOPJ Then
        Call WriteGuildRejectNewMember(Nombre, Replace(Replace(Text1.Text, ",", " "), vbNewLine, " "))
        'Sacamos el char de la lista de aspirantes
        Dim i As Long
        
        'Calculamos la cantidad de aspirantes de antemano para mejorar el rendimiento
        Dim Count_listCount As Long
        Count_listCount = frmGuildLeader.solicitudes.ListCount - 1
            
        For i = 0 To Count_listCount
            If frmGuildLeader.solicitudes.List(i) = Nombre Then
                frmGuildLeader.solicitudes.RemoveItem i
                Exit For
            End If
        Next i
        
        Me.Hide
        Unload frmCharInfo
    End If
    
    Unload Me

End Sub

Private Sub Text1_Change()
    If Len(Text1.Text) > MAX_PROPOSAL_LENGTH Then
        Text1.Text = Left$(Text1.Text, MAX_PROPOSAL_LENGTH)
    End If
End Sub
