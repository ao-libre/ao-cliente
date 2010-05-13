VERSION 5.00
Begin VB.Form frmGuildAdm 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4065
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
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBuscar 
      Appearance      =   0  'Flat
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
      Height          =   240
      Left            =   495
      TabIndex        =   1
      Top             =   4650
      Width           =   3105
   End
   Begin VB.ListBox GuildsList 
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
      Height          =   3540
      ItemData        =   "frmGuildAdm.frx":0000
      Left            =   495
      List            =   "frmGuildAdm.frx":0002
      TabIndex        =   0
      Top             =   570
      Width           =   3075
   End
   Begin VB.Image imgDetalles 
      Height          =   375
      Left            =   2280
      Tag             =   "1"
      Top             =   5025
      Width           =   1335
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   480
      Tag             =   "1"
      Top             =   5025
      Width           =   855
   End
End
Attribute VB_Name = "frmGuildAdm"
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

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton
Private cBotonDetalles As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
        
    Me.Picture = LoadPicture(App.path & "\graficos\VentanaListaClanes.jpg")
    
    Call LoadButtons
    
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonDetalles = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarListaClanes.jpg", _
                                    GrhPath & "BotonCerrarRolloverListaClanes.jpg", _
                                    GrhPath & "BotonCerrarClickListaClanes.jpg", Me)

    Call cBotonDetalles.Initialize(imgDetalles, GrhPath & "BotonDetallesListaClanes.jpg", _
                                    GrhPath & "BotonDetallesRolloverListaClanes.jpg", _
                                    GrhPath & "BotonDetallesClickListaClanes.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub guildslist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub imgDetalles_Click()
    frmGuildBrief.EsLeader = False

    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))
End Sub

Private Sub txtBuscar_Change()
    Call FiltrarListaClanes(txtBuscar.Text)
End Sub

Private Sub txtBuscar_GotFocus()
    With txtBuscar
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Public Sub FiltrarListaClanes(ByRef sCompare As String)

    Dim lIndex As Long
    
    If UBound(GuildNames) <> 0 Then
        With guildslist
            'Limpio la lista
            .Clear
            
            .Visible = False
            
            ' Recorro los arrays
            For lIndex = 0 To UBound(GuildNames)
                ' Si coincide con los patrones
                If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                    ' Lo agrego a la lista
                    .AddItem GuildNames(lIndex)
                End If
            Next lIndex
            
            .Visible = True
        End With
    End If

End Sub
