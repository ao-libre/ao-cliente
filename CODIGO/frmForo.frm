VERSION 5.00
Begin VB.Form frmForo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmForo.frx":0000
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   414
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTitulo 
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
      Height          =   315
      Left            =   1140
      MaxLength       =   35
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   4620
   End
   Begin VB.TextBox txtPost 
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
      Height          =   3960
      Left            =   780
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmForo.frx":2BBCD
      Top             =   1935
      Visible         =   0   'False
      Width           =   4770
   End
   Begin VB.ListBox lstTitulos 
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
      Height          =   5100
      Left            =   765
      TabIndex        =   0
      Top             =   825
      Width           =   4785
   End
   Begin VB.Label lblTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   1125
      TabIndex        =   4
      Top             =   960
      Width           =   4695
   End
   Begin VB.Image imgMarcoTexto 
      Height          =   465
      Left            =   1095
      Top             =   840
      Width           =   4725
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   4080
      Picture         =   "frmForo.frx":2BBD3
      Top             =   6060
      Width           =   1455
   End
   Begin VB.Image imgListaMsg 
      Height          =   360
      Left            =   2400
      Picture         =   "frmForo.frx":31CBF
      Top             =   6060
      Width           =   1455
   End
   Begin VB.Image imgDejarMsg 
      Height          =   360
      Left            =   720
      Picture         =   "frmForo.frx":35ED9
      Top             =   6060
      Width           =   1455
   End
   Begin VB.Label lblAutor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   1125
      TabIndex        =   3
      Top             =   1455
      Width           =   4650
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   2
      Left            =   4320
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   1
      Left            =   2520
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   0
      Left            =   960
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgDejarAnuncio 
      Height          =   255
      Left            =   2400
      Top             =   6060
      Width           =   1455
   End
End
Attribute VB_Name = "frmForo"
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

Private clsFormulario          As clsFormMovementManager

Private cBotonDejarAnuncio     As clsGraphicalButton
Private cBotonDejarMsg         As clsGraphicalButton
Private cBotonCerrar           As clsGraphicalButton
Private cBotonListaMsg         As clsGraphicalButton
Public LastButtonPressed       As clsGraphicalButton

' Para controlar las imagenes de fondo y el envio de posteos
Private ForoActual             As eForumType
Private VerListaMsg            As Boolean
Private Lectura                As Boolean

Public ForoLimpio              As Boolean
Private Sticky                 As Boolean

' Para restringir la visibilidad de los foros
Public Privilegios             As Byte
Public ForosVisibles           As eForumType
Public CanPostSticky           As Byte

' Imagenes de fondo
Private FondosDejarMsg(0 To 2) As Picture
Private FondosListaMsg(0 To 2) As Picture

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Form_Unload_Err
    
    MirandoForo = False
    Privilegios = 0

    
    Exit Sub

Form_Unload_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "Form_Unload"
    End If
Resume Next
    
End Sub

Private Sub imgDejarAnuncio_Click()
    
    On Error GoTo imgDejarAnuncio_Click_Err
    
    Lectura = False
    VerListaMsg = False
    Sticky = True
    
    'Switch to proper background
    ToogleScreen

    
    Exit Sub

imgDejarAnuncio_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "imgDejarAnuncio_Click"
    End If
Resume Next
    
End Sub

Private Sub imgDejarMsg_Click()
    
    On Error GoTo imgDejarMsg_Click_Err
    

    If Not cBotonDejarMsg.IsEnabled Then Exit Sub
    
    Dim PostStyle As Byte
    
    If Not VerListaMsg Then
        If Not Lectura Then
        
            If Sticky Then
                PostStyle = GetStickyPost
            Else
                PostStyle = GetNormalPost

            End If

            Call WriteForumPost(txtTitulo.Text, txtPost.Text, PostStyle)
            
            ' Actualizo localmente
            Call clsForos.AddPost(ForoActual, txtTitulo.Text, UserName, txtPost.Text, Sticky)
            Call UpdateList
            
            VerListaMsg = True

        End If

    Else
        VerListaMsg = False
        Sticky = False

    End If
    
    Lectura = False
    
    'Switch to proper background
    ToogleScreen

    
    Exit Sub

imgDejarMsg_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "imgDejarMsg_Click"
    End If
Resume Next
    
End Sub

Private Sub imgCerrar_Click()
    
    On Error GoTo imgCerrar_Click_Err
    
    Unload Me

    
    Exit Sub

imgCerrar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "imgCerrar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgListaMsg_Click()
    
    On Error GoTo imgListaMsg_Click_Err
    
    VerListaMsg = True
    ToogleScreen

    
    Exit Sub

imgListaMsg_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "imgListaMsg_Click"
    End If
Resume Next
    
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    
    On Error GoTo Form_Load_Err
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Call LoadButtons
    
    ' Initial config
    ForoActual = eForumType.ieGeneral
    VerListaMsg = True
    UpdateList
    
    ' Default background
    ToogleScreen
    
    ForoLimpio = False
    MirandoForo = True
    
    ' Si no es caos o gms, no puede ver el tab de caos.
    If (Privilegios And eForumVisibility.ieCAOS_MEMBER) = 0 Then imgTab(2).Visible = False
    
    ' Si no es armada o gm, no puede ver el tab de armadas.
    If (Privilegios And eForumVisibility.ieREAL_MEMBER) = 0 Then imgTab(1).Visible = False
    
    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    

    Dim GrhPath As String
    GrhPath = DirGraficos
    
    ' Load pictures
    Set FondosListaMsg(eForumType.ieGeneral) = LoadPicture(GrhPath & "ForoGeneral.jpg")
    Set FondosListaMsg(eForumType.ieREAL) = LoadPicture(GrhPath & "ForoReal.jpg")
    Set FondosListaMsg(eForumType.ieCAOS) = LoadPicture(GrhPath & "ForoCaos.jpg")
    
    Set FondosDejarMsg(eForumType.ieGeneral) = LoadPicture(GrhPath & "ForoMsgGeneral.jpg")
    Set FondosDejarMsg(eForumType.ieREAL) = LoadPicture(GrhPath & "ForoMsgReal.jpg")
    Set FondosDejarMsg(eForumType.ieCAOS) = LoadPicture(GrhPath & "ForoMsgCaos.jpg")
    
    imgMarcoTexto.Picture = LoadPicture(GrhPath & "MarcoTextBox.jpg")

    Set cBotonDejarAnuncio = New clsGraphicalButton
    Set cBotonDejarMsg = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonListaMsg = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton

    ' Initialize buttons
    Call cBotonDejarAnuncio.Initialize(imgDejarAnuncio, GrhPath & "BotonDejarAnuncioForo.jpg", GrhPath & "BotonDejarAnuncioRolloverForo.jpg", GrhPath & "BotonDejarAnuncioClickForo.jpg", Me)
                                            
    Call cBotonDejarMsg.Initialize(imgDejarMsg, GrhPath & "BotonDejarMsgForo.jpg", GrhPath & "BotonDejarMsgRolloverForo.jpg", GrhPath & "BotonDejarMsgClickForo.jpg", Me, GrhPath & "BotonDejarMsgDisabledForo.jpg")
                                            
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarForo.jpg", GrhPath & "BotonCerrarRolloverForo.jpg", GrhPath & "BotonCerrarClickForo.jpg", Me)
                                            
    Call cBotonListaMsg.Initialize(imgListaMsg, GrhPath & "BotonListaMsgForo.jpg", GrhPath & "BotonListaMsgRolloverForo.jpg", GrhPath & "BotonListaMsgClickForo.jpg", Me)

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgTab_Click(Index As Integer)
    
    On Error GoTo imgTab_Click_Err
    

    Call Audio.PlayWave(SND_CLICK)
    
    If Index <> ForoActual Then
        ForoActual = Index
        VerListaMsg = True
        Lectura = False
        UpdateList
        ToogleScreen
    Else

        If Not VerListaMsg Then
            VerListaMsg = True
            Lectura = False
            ToogleScreen

        End If

    End If

    
    Exit Sub

imgTab_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "imgTab_Click"
    End If
Resume Next
    
End Sub

Private Sub ToogleScreen()
    
    On Error GoTo ToogleScreen_Err
    
    
    Dim PostOffset As Integer
    
    imgMarcoTexto.Visible = Not VerListaMsg And Not Lectura
    txtTitulo.Visible = Not VerListaMsg And Not Lectura
    lblTitulo.Visible = Not VerListaMsg And Lectura
    
    Call cBotonDejarMsg.EnableButton(VerListaMsg Or Lectura)
    
    txtPost.Visible = Not VerListaMsg
    
    imgDejarAnuncio.Visible = VerListaMsg And PuedeDejarAnuncios
    imgListaMsg.Visible = Not VerListaMsg
    lstTitulos.Visible = VerListaMsg
    
    If VerListaMsg Then
        Me.Picture = FondosListaMsg(ForoActual)
    Else

        If Lectura Then

            With lstTitulos
                PostOffset = .ItemData(.ListIndex)
                
                ' Normal post?
                If PostOffset < STICKY_FORUM_OFFSET Then
                    lblTitulo.Caption = Foros(ForoActual).GeneralTitle(PostOffset)
                    txtPost.Text = Foros(ForoActual).GeneralPost(PostOffset)
                    lblAutor.Caption = Foros(ForoActual).GeneralAuthor(PostOffset)
                
                    ' Sticky post
                Else
                    PostOffset = PostOffset - STICKY_FORUM_OFFSET
                    
                    lblTitulo.Caption = Foros(ForoActual).StickyTitle(PostOffset)
                    txtPost.Text = Foros(ForoActual).StickyPost(PostOffset)
                    lblAutor.Caption = Foros(ForoActual).StickyAuthor(PostOffset)

                End If

            End With

        Else
            lblAutor.Caption = UserName
            txtTitulo.Text = vbNullString
            txtPost.Text = vbNullString
            
            txtTitulo.SetFocus

        End If
        
        txtPost.Locked = Lectura
        Me.Picture = FondosDejarMsg(ForoActual)

    End If
    
    
    Exit Sub

ToogleScreen_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "ToogleScreen"
    End If
Resume Next
    
End Sub

Private Function PuedeDejarAnuncios() As Boolean
    
    On Error GoTo PuedeDejarAnuncios_Err
    
    
    ' No puede
    If CanPostSticky = 0 Then Exit Function

    If ForoActual = eForumType.ieGeneral Then

        ' Solo puede dejar en el general si es gm
        If CanPostSticky <> 2 Then Exit Function

    End If
    
    PuedeDejarAnuncios = True
    
    
    Exit Function

PuedeDejarAnuncios_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "PuedeDejarAnuncios"
    End If
Resume Next
    
End Function

Private Sub lstTitulos_Click()
    
    On Error GoTo lstTitulos_Click_Err
    
    VerListaMsg = False
    Lectura = True
    ToogleScreen

    
    Exit Sub

lstTitulos_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "lstTitulos_Click"
    End If
Resume Next
    
End Sub

Private Sub lstTitulos_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo lstTitulos_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

lstTitulos_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "lstTitulos_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub txtPost_Change()
    
    On Error GoTo txtPost_Change_Err
    

    If Lectura Then Exit Sub
    
    Call cBotonDejarMsg.EnableButton(Len(txtTitulo.Text) <> 0 And Len(txtPost.Text) <> 0)

    
    Exit Sub

txtPost_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "txtPost_Change"
    End If
Resume Next
    
End Sub

Private Sub txtPost_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    
    On Error GoTo txtPost_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

txtPost_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "txtPost_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub txtTitulo_Change()
    
    On Error GoTo txtTitulo_Change_Err
    

    If Lectura Then Exit Sub
    
    Call cBotonDejarMsg.EnableButton(Len(txtTitulo.Text) <> 0 And Len(txtPost.Text) <> 0)

    
    Exit Sub

txtTitulo_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "txtTitulo_Change"
    End If
Resume Next
    
End Sub

Private Sub UpdateList()
    
    On Error GoTo UpdateList_Err
    
    Dim PostIndex As Long
    
    lstTitulos.Clear
    
    With lstTitulos

        ' Sticky first
        For PostIndex = 1 To clsForos.GetNroSticky(ForoActual)
            .AddItem "[ANUNCIO] " & Foros(ForoActual).StickyTitle(PostIndex) & " (" & Foros(ForoActual).StickyAuthor(PostIndex) & ")"
            .ItemData(.NewIndex) = STICKY_FORUM_OFFSET + PostIndex
        Next PostIndex
    
        ' Then normal posts
        For PostIndex = 1 To clsForos.GetNroPost(ForoActual)
            .AddItem Foros(ForoActual).GeneralTitle(PostIndex) & " (" & Foros(ForoActual).GeneralAuthor(PostIndex) & ")"
            .ItemData(.NewIndex) = PostIndex
        Next PostIndex

    End With
    
    
    Exit Sub

UpdateList_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "UpdateList"
    End If
Resume Next
    
End Sub

Private Function GetStickyPost() As Byte
    
    On Error GoTo GetStickyPost_Err
    

    Select Case ForoActual

        Case 0
            GetStickyPost = eForumMsgType.ieGENERAL_STICKY
            
        Case 1
            GetStickyPost = eForumMsgType.ieREAL_STICKY
            
        Case 2
            GetStickyPost = eForumMsgType.ieCAOS_STICKY
            
    End Select
    
    
    Exit Function

GetStickyPost_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "GetStickyPost"
    End If
Resume Next
    
End Function

Private Function GetNormalPost() As Byte
    
    On Error GoTo GetNormalPost_Err
    

    Select Case ForoActual

        Case 0
            GetNormalPost = eForumMsgType.ieGeneral
            
        Case 1
            GetNormalPost = eForumMsgType.ieREAL
            
        Case 2
            GetNormalPost = eForumMsgType.ieCAOS
            
    End Select
    
    
    Exit Function

GetNormalPost_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "GetNormalPost"
    End If
Resume Next
    
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyUp_Err
    
    If KeyCode = vbKeyEscape Then
        Set clsFormulario = Nothing
        Unload Me
    End If

    Exit Sub

Form_KeyUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmForo" & "->" & "Form_KeyUp"
    End If
Resume Next
    
End Sub
