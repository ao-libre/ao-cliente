VERSION 5.00
Begin VB.Form frmGuildMember 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMiembros 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2565
      Left            =   3075
      TabIndex        =   3
      Top             =   675
      Width           =   2610
   End
   Begin VB.ListBox lstClanes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2565
      Left            =   195
      TabIndex        =   2
      Top             =   690
      Width           =   2610
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   225
      Left            =   225
      TabIndex        =   1
      Top             =   3630
      Width           =   2550
   End
   Begin VB.Label lblCantMiembros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Height          =   195
      Left            =   4635
      TabIndex        =   0
      Top             =   3510
      Width           =   360
   End
   Begin VB.Image imgCerrar 
      Height          =   495
      Left            =   3000
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Image imgNoticias 
      Height          =   495
      Left            =   150
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Image imgDetalles 
      Height          =   375
      Left            =   150
      Top             =   4200
      Width           =   2655
   End
End
Attribute VB_Name = "frmGuildMember"
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

Private clsFormulario    As clsFormMovementManager

Private cBotonNoticias   As clsGraphicalButton
Private cBotonDetalles   As clsGraphicalButton
Private cBotonCerrar     As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Me.Picture = LoadPicture(DirGraficos & "VentanaMiembroClan.jpg")
    
    Call LoadButtons
    
    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildMember" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonNoticias = New clsGraphicalButton
    Set cBotonDetalles = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonDetalles.Initialize(imgDetalles, GrhPath & "BotonDetallesMiembroClan.jpg", GrhPath & "BotonDetallesRolloverMiembroClan.jpg", GrhPath & "BotonDetallesClickMiembroClan.jpg", Me)

    Call cBotonNoticias.Initialize(imgNoticias, GrhPath & "BotonNoticiasMiembroClan.jpg", GrhPath & "BotonNoticiasRolloverMiembroClan.jpg", GrhPath & "BotonNoticiasClickMiembroClan.jpg", Me)

    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarMimebroClan.jpg", GrhPath & "BotonCerrarRolloverMimebroClan.jpg", GrhPath & "BotonCerrarClickMimebroClan.jpg", Me)

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildMember" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildMember" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgCerrar_Click()
    
    On Error GoTo imgCerrar_Click_Err
    
    Unload Me

    
    Exit Sub

imgCerrar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildMember" & "->" & "imgCerrar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgDetalles_Click()
    
    On Error GoTo imgDetalles_Click_Err
    

    If lstClanes.ListIndex = -1 Then Exit Sub
    
    frmGuildBrief.EsLeader = False

    Call WriteGuildRequestDetails(lstClanes.List(lstClanes.ListIndex))

    
    Exit Sub

imgDetalles_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildMember" & "->" & "imgDetalles_Click"
    End If
Resume Next
    
End Sub

Private Sub imgNoticias_Click()
    
    On Error GoTo imgNoticias_Click_Err
    
    bShowGuildNews = True
    Call WriteShowGuildNews

    
    Exit Sub

imgNoticias_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildMember" & "->" & "imgNoticias_Click"
    End If
Resume Next
    
End Sub

Private Sub txtSearch_Change()
    
    On Error GoTo txtSearch_Change_Err
    
    Call FiltrarListaClanes(txtSearch.Text)

    
    Exit Sub

txtSearch_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildMember" & "->" & "txtSearch_Change"
    End If
Resume Next
    
End Sub

Private Sub txtSearch_GotFocus()
    
    On Error GoTo txtSearch_GotFocus_Err
    

    With txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)

    End With

    
    Exit Sub

txtSearch_GotFocus_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildMember" & "->" & "txtSearch_GotFocus"
    End If
Resume Next
    
End Sub

Public Sub FiltrarListaClanes(ByRef sCompare As String)
    
    On Error GoTo FiltrarListaClanes_Err
    

    Dim lIndex As Long
    
    If UBound(GuildNames) <> 0 Then

        With lstClanes
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

    
    Exit Sub

FiltrarListaClanes_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildMember" & "->" & "FiltrarListaClanes"
    End If
Resume Next
    
End Sub

