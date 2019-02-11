VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   Caption         =   "Creaci�n de un Clan"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtClanName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1815
      Width           =   3345
   End
   Begin VB.TextBox txtWeb 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   3345
   End
   Begin VB.Image imgSiguiente 
      Height          =   375
      Left            =   2400
      Tag             =   "1"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Image imgCancelar 
      Height          =   375
      Left            =   240
      Tag             =   "1"
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit
Private clsFormulario    As clsFormMovementManager

Private cBotonSiguiente  As clsGraphicalButton
Private cBotonCancelar   As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Deactivate()
    
    On Error GoTo Form_Deactivate_Err
    
    Me.SetFocus

    
    Exit Sub

Form_Deactivate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildFoundation" & "->" & "Form_Deactivate"
    End If
Resume Next
    
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    
    On Error GoTo Form_Load_Err
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Me.Picture = LoadPicture(App.path & "\graficos\VentanaNombreClan.jpg")
        
    Call LoadButtons
    
    If Len(txtClanName.Text) <= 30 Then
        If Not AsciiValidos(txtClanName) Then
            MsgBox "Nombre invalido."
            Exit Sub

        End If

    Else
        MsgBox "Nombre demasiado extenso."
        Exit Sub

    End If

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildFoundation" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonSiguiente = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonSiguiente.Initialize(imgSiguiente, GrhPath & "BotonSiguienteNombreClan.jpg", GrhPath & "BotonSiguienteRolloverNombreClan.jpg", GrhPath & "BotonSiguienteClickNombreClan.jpg", Me)

    Call cBotonCancelar.Initialize(imgCancelar, GrhPath & "BotonCancelarNombreClan.jpg", GrhPath & "BotonCancelarRolloverNombreClan.jpg", GrhPath & "BotonCancelarClickNombreClan.jpg", Me)

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildFoundation" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildFoundation" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgCancelar_Click()
    
    On Error GoTo imgCancelar_Click_Err
    
    Unload Me

    
    Exit Sub

imgCancelar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildFoundation" & "->" & "imgCancelar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgSiguiente_Click()
    
    On Error GoTo imgSiguiente_Click_Err
    
    ClanName = txtClanName.Text
    Site = txtWeb.Text
    Unload Me
    frmGuildDetails.Show , frmMain

    
    Exit Sub

imgSiguiente_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildFoundation" & "->" & "imgSiguiente_Click"
    End If
Resume Next
    
End Sub

Private Sub txtWeb_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    
    On Error GoTo txtWeb_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

txtWeb_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildFoundation" & "->" & "txtWeb_MouseMove"
    End If
Resume Next
    
End Sub
