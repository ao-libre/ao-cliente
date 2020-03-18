VERSION 5.00
Begin VB.Form frmQuests 
   BorderStyle     =   0  'None
   Caption         =   "Misiones"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   8355
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQuests.frx":0000
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   557
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   4035
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   900
      Width           =   5415
   End
   Begin VB.ListBox lstQuests 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   180
      TabIndex        =   0
      Top             =   915
      Width           =   2355
   End
   Begin AOLibre.uAOButton Salir 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmQuests.frx":22306
      PICF            =   "frmQuests.frx":22D30
      PICH            =   "frmQuests.frx":239F2
      PICV            =   "frmQuests.frx":24984
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
   Begin AOLibre.uAOButton Abandonar 
      Height          =   615
      Left            =   225
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      TX              =   "Abandonar Mision"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmQuests.frx":25886
      PICF            =   "frmQuests.frx":262B0
      PICH            =   "frmQuests.frx":26F72
      PICV            =   "frmQuests.frx":27F04
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Informacion de misiones:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Option Explicit

Private Sub Abandonar_Click()

            If lstQuests.ListCount = 0 Then
                MsgBox "No tienes ninguna mision!", vbOKOnly + vbExclamation
                Exit Sub
            End If
            
            'Chequeamos si tiene algun item seleccionado.
            If lstQuests.ListIndex < 0 Then
                MsgBox "Primero debes seleccionar una mision!", vbOKOnly + vbExclamation
                Exit Sub
            End If
            
            Select Case MsgBox("Estas seguro que deseas abandonar la mision?", vbYesNo + vbExclamation)
                Case vbYes  'Boton Si.
                    'Enviamos el paquete para abandonar la quest
                    Call WriteQuestAbandon(lstQuests.ListIndex + 1)
                    
                Case vbNo   'Boton NO.
                    'Como selecciono que no, no hace nada.
                    Exit Sub
            End Select
            
End Sub

Private Sub Form_Load()
    
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaInfoQuest.jpg")
    
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)

End Sub

Private Sub LoadTextsForm()
    Me.lblDesc.Caption = JsonLanguage.item("FRM_QUESTINFO_DESC").item("TEXTO")
    Me.Abandonar.Caption = JsonLanguage.item("FRM_QUESTINFO_ABAND").item("TEXTO")
    Me.Salir.Caption = JsonLanguage.item("FRM_QUESTINFO_EXIT").item("TEXTO")
End Sub

Private Sub lstQuests_Click()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el click del ListBox lstQuests.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If lstQuests.ListIndex < 0 Then Exit Sub
    
    Call WriteQuestDetailsRequest(lstQuests.ListIndex + 1)
End Sub

Private Sub Salir_Click()
    Unload Me
End Sub
