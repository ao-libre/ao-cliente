VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   0  'None
   Caption         =   "Administración del Clan"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5985
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
   ScaleHeight     =   494
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFiltrarMiembros 
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
      Height          =   225
      Left            =   3075
      TabIndex        =   6
      Top             =   2340
      Width           =   2580
   End
   Begin VB.TextBox txtFiltrarClanes 
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
      Left            =   195
      TabIndex        =   5
      Top             =   2340
      Width           =   2580
   End
   Begin VB.TextBox txtguildnews 
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
      Height          =   690
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3435
      Width           =   5475
   End
   Begin VB.ListBox solicitudes 
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
      Height          =   810
      ItemData        =   "frmGuildLeader.frx":0000
      Left            =   195
      List            =   "frmGuildLeader.frx":0002
      TabIndex        =   2
      Top             =   5100
      Width           =   2595
   End
   Begin VB.ListBox members 
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
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":0004
      Left            =   3060
      List            =   "frmGuildLeader.frx":0006
      TabIndex        =   1
      Top             =   540
      Width           =   2595
   End
   Begin VB.ListBox guildslist 
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
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":0008
      Left            =   180
      List            =   "frmGuildLeader.frx":000A
      TabIndex        =   0
      Top             =   540
      Width           =   2595
   End
   Begin VB.Image imgCerrar 
      Height          =   495
      Left            =   3000
      Tag             =   "1"
      Top             =   6705
      Width           =   2775
   End
   Begin VB.Image imgPropuestasAlianzas 
      Height          =   495
      Left            =   3000
      Tag             =   "1"
      Top             =   6195
      Width           =   2775
   End
   Begin VB.Image imgPropuestasPaz 
      Height          =   495
      Left            =   3000
      Tag             =   "1"
      Top             =   5685
      Width           =   2775
   End
   Begin VB.Image imgEditarURL 
      Height          =   495
      Left            =   3000
      Tag             =   "1"
      Top             =   5175
      Width           =   2775
   End
   Begin VB.Image imgEditarCodex 
      Height          =   495
      Left            =   3000
      Tag             =   "1"
      Top             =   4665
      Width           =   2775
   End
   Begin VB.Image imgActualizar 
      Height          =   390
      Left            =   150
      Tag             =   "1"
      Top             =   4230
      Width           =   5550
   End
   Begin VB.Image imgDetallesSolicitudes 
      Height          =   375
      Left            =   120
      Tag             =   "1"
      Top             =   6045
      Width           =   2655
   End
   Begin VB.Image imgDetallesMiembros 
      Height          =   375
      Left            =   3060
      Tag             =   "1"
      Top             =   2700
      Width           =   2655
   End
   Begin VB.Image imgDetallesClan 
      Height          =   375
      Left            =   165
      Tag             =   "1"
      Top             =   2700
      Width           =   2655
   End
   Begin VB.Image imgElecciones 
      Height          =   375
      Left            =   120
      Tag             =   "1"
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Label Miembros 
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
      Left            =   1815
      TabIndex        =   3
      Top             =   6510
      Width           =   255
   End
End
Attribute VB_Name = "frmGuildLeader"
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

Private Const MAX_NEWS_LENGTH     As Integer = 512
Private clsFormulario             As clsFormMovementManager

Private cBotonElecciones          As clsGraphicalButton
Private cBotonActualizar          As clsGraphicalButton
Private cBotonDetallesClan        As clsGraphicalButton
Private cBotonDetallesMiembros    As clsGraphicalButton
Private cBotonDetallesSolicitudes As clsGraphicalButton
Private cBotonEditarCodex         As clsGraphicalButton
Private cBotonEditarURL           As clsGraphicalButton
Private cBotonPropuestasPaz       As clsGraphicalButton
Private cBotonPropuestasAlianzas  As clsGraphicalButton
Private cBotonCerrar              As clsGraphicalButton

Public LastButtonPressed          As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    
    On Error GoTo Form_Load_Err
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(App.path & "\graficos\VentanaAdministrarClan.jpg")
    
    Call LoadButtons

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonElecciones = New clsGraphicalButton
    Set cBotonActualizar = New clsGraphicalButton
    Set cBotonDetallesClan = New clsGraphicalButton
    Set cBotonDetallesMiembros = New clsGraphicalButton
    Set cBotonDetallesSolicitudes = New clsGraphicalButton
    Set cBotonEditarCodex = New clsGraphicalButton
    Set cBotonEditarURL = New clsGraphicalButton
    Set cBotonPropuestasPaz = New clsGraphicalButton
    Set cBotonPropuestasAlianzas = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonElecciones.Initialize(imgElecciones, GrhPath & "BotonElecciones.jpg", GrhPath & "BotonEleccionesRollover.jpg", GrhPath & "BotonEleccionesClick.jpg", Me)

    Call cBotonActualizar.Initialize(imgActualizar, GrhPath & "BotonActualizar.jpg", GrhPath & "BotonActualizarRollover.jpg", GrhPath & "BotonActualizarClick.jpg", Me)

    Call cBotonDetallesClan.Initialize(imgDetallesClan, GrhPath & "BotonDetallesAdministrarClan.jpg", GrhPath & "BotonDetallesRolloverAdministrarClan.jpg", GrhPath & "BotonDetallesClickAdministrarClan.jpg", Me)

    Call cBotonDetallesMiembros.Initialize(imgDetallesMiembros, GrhPath & "BotonDetallesAdministrarClan.jpg", GrhPath & "BotonDetallesRolloverAdministrarClan.jpg", GrhPath & "BotonDetallesClickAdministrarClan.jpg", Me)
                                    
    Call cBotonDetallesSolicitudes.Initialize(imgDetallesSolicitudes, GrhPath & "BotonDetallesAdministrarClan.jpg", GrhPath & "BotonDetallesRolloverAdministrarClan.jpg", GrhPath & "BotonDetallesClickAdministrarClan.jpg", Me)

    Call cBotonEditarCodex.Initialize(imgEditarCodex, GrhPath & "BotonEditarCodex.jpg", GrhPath & "BotonEditarCodexRollover.jpg", GrhPath & "BotonEditarCodexClick.jpg", Me)

    Call cBotonEditarURL.Initialize(imgEditarURL, GrhPath & "BotonEditarURL.jpg", GrhPath & "BotonEditarURLRollover.jpg", GrhPath & "BotonEditarURLClick.jpg", Me)

    Call cBotonPropuestasPaz.Initialize(imgPropuestasPaz, GrhPath & "BotonPropuestaPaz.jpg", GrhPath & "BotonPropuestaPazRollover.jpg", GrhPath & "BotonPropuestaPazClick.jpg", Me)

    Call cBotonPropuestasAlianzas.Initialize(imgPropuestasAlianzas, GrhPath & "BotonPropuestasAlianzas.jpg", GrhPath & "BotonPropuestasAlianzasRollover.jpg", GrhPath & "BotonPropuestasAlianzasClick.jpg", Me)

    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarAdministrarClan.jpg", GrhPath & "BotonCerrarRolloverAdministrarClan.jpg", GrhPath & "BotonCerrarClickAdministrarClan.jpg", Me)

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub guildslist_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo guildslist_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

guildslist_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "guildslist_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgActualizar_Click()
    
    On Error GoTo imgActualizar_Click_Err
    
    Dim k As String

    k = Replace(txtguildnews, vbNewLine, "º")
    
    Call WriteGuildUpdateNews(k)

    
    Exit Sub

imgActualizar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "imgActualizar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgCerrar_Click()
    
    On Error GoTo imgCerrar_Click_Err
    
    Unload Me
    frmMain.SetFocus

    
    Exit Sub

imgCerrar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "imgCerrar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgDetallesClan_Click()
    
    On Error GoTo imgDetallesClan_Click_Err
    
    frmGuildBrief.EsLeader = True
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))

    
    Exit Sub

imgDetallesClan_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "imgDetallesClan_Click"
    End If
Resume Next
    
End Sub

Private Sub imgDetallesMiembros_Click()
    
    On Error GoTo imgDetallesMiembros_Click_Err
    

    If members.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.List(members.ListIndex))

    
    Exit Sub

imgDetallesMiembros_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "imgDetallesMiembros_Click"
    End If
Resume Next
    
End Sub

Private Sub imgDetallesSolicitudes_Click()
    
    On Error GoTo imgDetallesSolicitudes_Click_Err
    

    If solicitudes.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(solicitudes.List(solicitudes.ListIndex))

    
    Exit Sub

imgDetallesSolicitudes_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "imgDetallesSolicitudes_Click"
    End If
Resume Next
    
End Sub

Private Sub imgEditarCodex_Click()
    
    On Error GoTo imgEditarCodex_Click_Err
    
    Call frmGuildDetails.Show(vbModal, frmGuildLeader)

    
    Exit Sub

imgEditarCodex_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "imgEditarCodex_Click"
    End If
Resume Next
    
End Sub

Private Sub imgEditarURL_Click()
    
    On Error GoTo imgEditarURL_Click_Err
    
    Call frmGuildURL.Show(vbModeless, frmGuildLeader)

    
    Exit Sub

imgEditarURL_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "imgEditarURL_Click"
    End If
Resume Next
    
End Sub

Private Sub imgElecciones_Click()
    
    On Error GoTo imgElecciones_Click_Err
    
    Call WriteGuildOpenElections
    Unload Me

    
    Exit Sub

imgElecciones_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "imgElecciones_Click"
    End If
Resume Next
    
End Sub

Private Sub imgPropuestasAlianzas_Click()
    
    On Error GoTo imgPropuestasAlianzas_Click_Err
    
    Call WriteGuildAlliancePropList

    
    Exit Sub

imgPropuestasAlianzas_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "imgPropuestasAlianzas_Click"
    End If
Resume Next
    
End Sub

Private Sub imgPropuestasPaz_Click()
    
    On Error GoTo imgPropuestasPaz_Click_Err
    
    Call WriteGuildPeacePropList

    
    Exit Sub

imgPropuestasPaz_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "imgPropuestasPaz_Click"
    End If
Resume Next
    
End Sub

Private Sub members_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    
    On Error GoTo members_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

members_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "members_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub solicitudes_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    
    On Error GoTo solicitudes_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

solicitudes_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "solicitudes_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub txtguildnews_Change()
    
    On Error GoTo txtguildnews_Change_Err
    

    If Len(txtguildnews.Text) > MAX_NEWS_LENGTH Then txtguildnews.Text = Left$(txtguildnews.Text, MAX_NEWS_LENGTH)

    
    Exit Sub

txtguildnews_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "txtguildnews_Change"
    End If
Resume Next
    
End Sub

Private Sub txtguildnews_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    
    On Error GoTo txtguildnews_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

txtguildnews_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "txtguildnews_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub txtFiltrarClanes_Change()
    
    On Error GoTo txtFiltrarClanes_Change_Err
    
    Call FiltrarListaClanes(txtFiltrarClanes.Text)

    
    Exit Sub

txtFiltrarClanes_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "txtFiltrarClanes_Change"
    End If
Resume Next
    
End Sub

Private Sub txtFiltrarClanes_GotFocus()
    
    On Error GoTo txtFiltrarClanes_GotFocus_Err
    

    With txtFiltrarClanes
        .SelStart = 0
        .SelLength = Len(.Text)

    End With

    
    Exit Sub

txtFiltrarClanes_GotFocus_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "txtFiltrarClanes_GotFocus"
    End If
Resume Next
    
End Sub

Private Sub FiltrarListaClanes(ByRef sCompare As String)
    
    On Error GoTo FiltrarListaClanes_Err
    

    Dim lIndex As Long
    
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

    
    Exit Sub

FiltrarListaClanes_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "FiltrarListaClanes"
    End If
Resume Next
    
End Sub

Private Sub txtFiltrarMiembros_Change()
    
    On Error GoTo txtFiltrarMiembros_Change_Err
    
    Call FiltrarListaMiembros(txtFiltrarMiembros.Text)

    
    Exit Sub

txtFiltrarMiembros_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "txtFiltrarMiembros_Change"
    End If
Resume Next
    
End Sub

Private Sub txtFiltrarMiembros_GotFocus()
    
    On Error GoTo txtFiltrarMiembros_GotFocus_Err
    

    With txtFiltrarMiembros
        .SelStart = 0
        .SelLength = Len(.Text)

    End With

    
    Exit Sub

txtFiltrarMiembros_GotFocus_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "txtFiltrarMiembros_GotFocus"
    End If
Resume Next
    
End Sub

Private Sub FiltrarListaMiembros(ByRef sCompare As String)
    
    On Error GoTo FiltrarListaMiembros_Err
    

    Dim lIndex As Long
    
    With members
        'Limpio la lista
        .Clear
        
        .Visible = False
        
        ' Recorro los arrays
        For lIndex = 0 To UBound(GuildMembers)

            ' Si coincide con los patrones
            If InStr(1, UCase$(GuildMembers(lIndex)), UCase$(sCompare)) Then
                ' Lo agrego a la lista
                .AddItem GuildMembers(lIndex)

            End If

        Next lIndex
        
        .Visible = True

    End With

    
    Exit Sub

FiltrarListaMiembros_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmGuildLeader" & "->" & "FiltrarListaMiembros"
    End If
Resume Next
    
End Sub

