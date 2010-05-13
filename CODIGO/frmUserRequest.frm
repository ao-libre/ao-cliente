VERSION 5.00
Begin VB.Form frmUserRequest 
   BorderStyle     =   0  'None
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4650
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
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
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
      Height          =   1395
      Left            =   225
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   405
      Width           =   4185
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   120
      Tag             =   "1"
      Top             =   1920
      Width           =   4350
   End
End
Attribute VB_Name = "frmUserRequest"
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

Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private clsFormulario As New clsFormMovementManager

Public Sub recievePeticion(ByVal p As String)

    Text1 = Replace$(p, "º", vbCrLf)
    Me.Show vbModeless, frmMain

End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(App.path & "\graficos\VentanaPeticion.jpg")
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarPeticion.jpg", _
                                    GrhPath & "BotonCerrarRolloverPeticion.jpg", _
                                    GrhPath & "BotonCerrarClickPeticion.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub
