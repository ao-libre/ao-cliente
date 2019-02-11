VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
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
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantMensajes 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "5"
      Top             =   2415
      Width           =   255
   End
   Begin VB.TextBox txtLevel 
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
      Height          =   255
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "40"
      Top             =   4110
      Width           =   255
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Index           =   0
      Left            =   1380
      TabIndex        =   0
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Max             =   100
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Top             =   1260
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   10
      Max             =   100
      TickStyle       =   3
   End
   Begin VB.Image imgChkDesactivarFragShooter 
      Height          =   225
      Left            =   435
      Top             =   4740
      Width           =   210
   End
   Begin VB.Image imgChkAlMorir 
      Height          =   225
      Left            =   435
      Top             =   4425
      Width           =   210
   End
   Begin VB.Image imgChkRequiredLvl 
      Height          =   225
      Left            =   435
      Top             =   4110
      Width           =   210
   End
   Begin VB.Image imgChkNoMostrarNews 
      Height          =   225
      Left            =   2475
      Top             =   3315
      Width           =   210
   End
   Begin VB.Image imgChkMostrarNews 
      Height          =   225
      Left            =   435
      Top             =   3315
      Width           =   210
   End
   Begin VB.Image imgChkPantalla 
      Height          =   225
      Left            =   1950
      Top             =   2430
      Width           =   210
   End
   Begin VB.Image imgChkConsola 
      Height          =   225
      Left            =   435
      Top             =   2430
      Width           =   210
   End
   Begin VB.Image imgChkEfectosSonido 
      Height          =   225
      Left            =   435
      Top             =   1545
      Width           =   210
   End
   Begin VB.Image imgChkSonidos 
      Height          =   225
      Left            =   435
      Top             =   1260
      Width           =   210
   End
   Begin VB.Image imgChkMusica 
      Height          =   225
      Left            =   435
      Top             =   990
      Width           =   210
   End
   Begin VB.Image imgTutorial 
      Height          =   285
      Left            =   2520
      Top             =   6240
      Width           =   2010
   End
   Begin VB.Image imgSoporte 
      Height          =   285
      Left            =   360
      Top             =   6240
      Width           =   2010
   End
   Begin VB.Image imgRadio 
      Height          =   285
      Left            =   2520
      Top             =   5880
      Width           =   2010
   End
   Begin VB.Image imgManual 
      Height          =   285
      Left            =   360
      Top             =   5880
      Width           =   2010
   End
   Begin VB.Image imgCambiarPasswd 
      Height          =   285
      Left            =   2520
      Top             =   5520
      Width           =   2010
   End
   Begin VB.Image imgMapa 
      Height          =   285
      Left            =   360
      Top             =   5520
      Width           =   2010
   End
   Begin VB.Image imgMsgPersonalizado 
      Height          =   285
      Left            =   2520
      Top             =   5160
      Width           =   2010
   End
   Begin VB.Image imgConfigTeclas 
      Height          =   285
      Left            =   360
      Top             =   5160
      Width           =   2010
   End
   Begin VB.Image imgSalir 
      Height          =   285
      Left            =   1440
      Top             =   6600
      Width           =   2010
   End
End
Attribute VB_Name = "frmOpciones"
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

Private clsFormulario          As clsFormMovementManager

Private cBotonConfigTeclas     As clsGraphicalButton
Private cBotonMsgPersonalizado As clsGraphicalButton
Private cBotonMapa             As clsGraphicalButton
Private cBotonCambiarPasswd    As clsGraphicalButton
Private cBotonManual           As clsGraphicalButton
Private cBotonRadio            As clsGraphicalButton
Private cBotonSoporte          As clsGraphicalButton
Private cBotonTutorial         As clsGraphicalButton
Private cBotonSalir            As clsGraphicalButton

Public LastButtonPressed       As clsGraphicalButton

Private picCheckBox            As Picture

Private bMusicActivated        As Boolean
Private bSoundActivated        As Boolean
Private bSoundEffectsActivated As Boolean

Private loading                As Boolean

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgCambiarPasswd_Click()
    
    On Error GoTo imgCambiarPasswd_Click_Err
    
    Call frmNewPassword.Show(vbModal, Me)

    
    Exit Sub

imgCambiarPasswd_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgCambiarPasswd_Click"
    End If
Resume Next
    
End Sub

Private Sub imgChkAlMorir_Click()
    
    On Error GoTo imgChkAlMorir_Click_Err
    
    ClientSetup.bDie = Not ClientSetup.bDie
    
    If ClientSetup.bDie Then
        imgChkAlMorir.Picture = picCheckBox
    Else
        Set imgChkAlMorir.Picture = Nothing

    End If

    
    Exit Sub

imgChkAlMorir_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgChkAlMorir_Click"
    End If
Resume Next
    
End Sub

Private Sub imgChkDesactivarFragShooter_Click()
    
    On Error GoTo imgChkDesactivarFragShooter_Click_Err
    
    ClientSetup.bActive = Not ClientSetup.bActive
    
    If ClientSetup.bActive Then
        Set imgChkDesactivarFragShooter.Picture = Nothing
    Else
        imgChkDesactivarFragShooter.Picture = picCheckBox

    End If

    
    Exit Sub

imgChkDesactivarFragShooter_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgChkDesactivarFragShooter_Click"
    End If
Resume Next
    
End Sub

Private Sub imgChkRequiredLvl_Click()
    
    On Error GoTo imgChkRequiredLvl_Click_Err
    
    ClientSetup.bKill = Not ClientSetup.bKill
    
    If ClientSetup.bKill Then
        imgChkRequiredLvl.Picture = picCheckBox
    Else
        Set imgChkRequiredLvl.Picture = Nothing

    End If

    
    Exit Sub

imgChkRequiredLvl_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgChkRequiredLvl_Click"
    End If
Resume Next
    
End Sub

Private Sub txtCantMensajes_Change()
    
    On Error GoTo txtCantMensajes_Change_Err
    
    txtCantMensajes.Text = Val(txtCantMensajes.Text)
    
    If txtCantMensajes.Text > 0 Then
        DialogosClanes.CantidadDialogos = txtCantMensajes.Text
    Else
        txtCantMensajes.Text = 5

    End If

    
    Exit Sub

txtCantMensajes_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "txtCantMensajes_Change"
    End If
Resume Next
    
End Sub

Private Sub txtLevel_Change()
    
    On Error GoTo txtLevel_Change_Err
    

    If Not IsNumeric(txtLevel) Then txtLevel = 0
    txtLevel = Trim$(txtLevel)
    ClientSetup.byMurderedLevel = CByte(txtLevel)

    
    Exit Sub

txtLevel_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "txtLevel_Change"
    End If
Resume Next
    
End Sub

Private Sub imgChkConsola_Click()
    
    On Error GoTo imgChkConsola_Click_Err
    
    DialogosClanes.Activo = False
    
    imgChkConsola.Picture = picCheckBox
    Set imgChkPantalla.Picture = Nothing

    
    Exit Sub

imgChkConsola_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgChkConsola_Click"
    End If
Resume Next
    
End Sub

Private Sub imgChkEfectosSonido_Click()
    
    On Error GoTo imgChkEfectosSonido_Click_Err
    

    If loading Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)
        
    bSoundEffectsActivated = Not bSoundEffectsActivated
    
    Audio.SoundEffectsActivated = bSoundEffectsActivated
    
    If bSoundEffectsActivated Then
        imgChkEfectosSonido.Picture = picCheckBox
    Else
        Set imgChkEfectosSonido.Picture = Nothing

    End If
            
    
    Exit Sub

imgChkEfectosSonido_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgChkEfectosSonido_Click"
    End If
Resume Next
    
End Sub

Private Sub imgChkMostrarNews_Click()
    
    On Error GoTo imgChkMostrarNews_Click_Err
    
    ClientSetup.bGuildNews = True
    
    imgChkMostrarNews.Picture = picCheckBox
    Set imgChkNoMostrarNews.Picture = Nothing

    
    Exit Sub

imgChkMostrarNews_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgChkMostrarNews_Click"
    End If
Resume Next
    
End Sub

Private Sub imgChkMusica_Click()
    
    On Error GoTo imgChkMusica_Click_Err
    

    If loading Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)
    
    bMusicActivated = Not bMusicActivated
            
    If Not bMusicActivated Then
        Audio.MusicActivated = False
        Slider1(0).Enabled = False
        Set imgChkMusica.Picture = Nothing
    Else

        If Not Audio.MusicActivated Then  'Prevent the music from reloading
            Audio.MusicActivated = True
            Slider1(0).Enabled = True
            Slider1(0).value = Audio.MusicVolume

        End If
        
        imgChkMusica.Picture = picCheckBox

    End If

    
    Exit Sub

imgChkMusica_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgChkMusica_Click"
    End If
Resume Next
    
End Sub

Private Sub imgChkNoMostrarNews_Click()
    
    On Error GoTo imgChkNoMostrarNews_Click_Err
    
    ClientSetup.bGuildNews = False
    
    imgChkNoMostrarNews.Picture = picCheckBox
    Set imgChkMostrarNews.Picture = Nothing

    
    Exit Sub

imgChkNoMostrarNews_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgChkNoMostrarNews_Click"
    End If
Resume Next
    
End Sub

Private Sub imgChkPantalla_Click()
    
    On Error GoTo imgChkPantalla_Click_Err
    
    DialogosClanes.Activo = True
    
    imgChkPantalla.Picture = picCheckBox
    Set imgChkConsola.Picture = Nothing

    
    Exit Sub

imgChkPantalla_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgChkPantalla_Click"
    End If
Resume Next
    
End Sub

Private Sub imgChkSonidos_Click()
    
    On Error GoTo imgChkSonidos_Click_Err
    

    If loading Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)
    
    bSoundActivated = Not bSoundActivated
    
    If Not bSoundActivated Then
        Audio.SoundActivated = False
        RainBufferIndex = 0
        frmMain.IsPlaying = PlayLoop.plNone
        Slider1(1).Enabled = False
        
        Set imgChkSonidos.Picture = Nothing
    Else
        Audio.SoundActivated = True
        Slider1(1).Enabled = True
        Slider1(1).value = Audio.SoundVolume
        
        imgChkSonidos.Picture = picCheckBox

    End If

    
    Exit Sub

imgChkSonidos_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgChkSonidos_Click"
    End If
Resume Next
    
End Sub

Private Sub imgConfigTeclas_Click()
    
    On Error GoTo imgConfigTeclas_Click_Err
    

    If Not loading Then Call Audio.PlayWave(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)

    
    Exit Sub

imgConfigTeclas_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgConfigTeclas_Click"
    End If
Resume Next
    
End Sub

Private Sub imgManual_Click()
    
    On Error GoTo imgManual_Click_Err
    

    If Not loading Then Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://ao.alkon.com.ar/manual/", "", App.path, SW_SHOWNORMAL)

    
    Exit Sub

imgManual_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgManual_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMapa_Click()
    
    On Error GoTo imgMapa_Click_Err
    
    Call frmMapa.Show(vbModal, Me)

    
    Exit Sub

imgMapa_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgMapa_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMsgPersonalizado_Click()
    
    On Error GoTo imgMsgPersonalizado_Click_Err
    
    Call frmMessageTxt.Show(vbModeless, Me)

    
    Exit Sub

imgMsgPersonalizado_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgMsgPersonalizado_Click"
    End If
Resume Next
    
End Sub

Private Sub imgRadio_Click()
    
    On Error GoTo imgRadio_Click_Err
    

    If Not loading Then Call Audio.PlayWave(SND_CLICK)

    Call ShellExecute(0, "Open", "http://www.radioargentum.com/repro.html", "", App.path, SW_SHOWNORMAL)

    
    Exit Sub

imgRadio_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgRadio_Click"
    End If
Resume Next
    
End Sub

Private Sub imgSalir_Click()
    
    On Error GoTo imgSalir_Click_Err
    
    Unload Me
    frmMain.SetFocus

    
    Exit Sub

imgSalir_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgSalir_Click"
    End If
Resume Next
    
End Sub

Private Sub imgSoporte_Click()
    
    On Error GoTo imgSoporte_Click_Err
    
    
    If Not loading Then Call Audio.PlayWave(SND_CLICK)
    
    Call ShellExecute(0, "Open", "http://www.aostaff.com.ar/soporte/", "", App.path, SW_SHOWNORMAL)

    
    Exit Sub

imgSoporte_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgSoporte_Click"
    End If
Resume Next
    
End Sub

Private Sub imgTutorial_Click()
    
    On Error GoTo imgTutorial_Click_Err
    
    frmTutorial.Show vbModeless

    
    Exit Sub

imgTutorial_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "imgTutorial_Click"
    End If
Resume Next
    
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    
    On Error GoTo Form_Load_Err
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(App.path & "\graficos\VentanaOpciones.jpg")
    LoadButtons
    
    loading = True      'Prevent sounds when setting check's values
    LoadUserConfig
    loading = False     'Enable sounds when setting check's values

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonConfigTeclas = New clsGraphicalButton
    Set cBotonMsgPersonalizado = New clsGraphicalButton
    Set cBotonMapa = New clsGraphicalButton
    Set cBotonCambiarPasswd = New clsGraphicalButton
    Set cBotonManual = New clsGraphicalButton
    Set cBotonRadio = New clsGraphicalButton
    Set cBotonSoporte = New clsGraphicalButton
    Set cBotonTutorial = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonConfigTeclas.Initialize(imgConfigTeclas, GrhPath & "BotonConfigurarTeclas.jpg", GrhPath & "BotonConfigurarTeclasRollover.jpg", GrhPath & "BotonConfigurarTeclasClick.jpg", Me)
                                    
    Call cBotonMsgPersonalizado.Initialize(imgMsgPersonalizado, GrhPath & "BotonMsgPersonalizadoTeclas.jpg", GrhPath & "BotonMsgPersonalizadoRollover.jpg", GrhPath & "BotonMsgPersonalizadoClick.jpg", Me)
                                    
    Call cBotonMapa.Initialize(imgMapa, GrhPath & "BotonMapaAo.jpg", GrhPath & "BotonMapaAoRollover.jpg", GrhPath & "BotonMapaAoClick.jpg", Me)
                                    
    Call cBotonCambiarPasswd.Initialize(imgCambiarPasswd, GrhPath & "BotonCambiarContrasenia.jpg", GrhPath & "BotonCambiarContraseniaRollover.jpg", GrhPath & "BotonCambiarContraseniaClick.jpg", Me)
                                    
    Call cBotonManual.Initialize(imgManual, GrhPath & "BotonManualAo.jpg", GrhPath & "BotonManualAoRollover.jpg", GrhPath & "BotonManualAoClick.jpg", Me)
                                    
    Call cBotonRadio.Initialize(imgRadio, GrhPath & "BotonEscucharRadioAo.jpg", GrhPath & "BotonEscucharRadioAoRollover.jpg", GrhPath & "BotonEscucharRadioAoClick.jpg", Me)
                                    
    Call cBotonSoporte.Initialize(imgSoporte, GrhPath & "BotonSoporte.jpg", GrhPath & "BotonSoporteRollover.jpg", GrhPath & "BotonSoporteClick.jpg", Me)
                                    
    Call cBotonTutorial.Initialize(imgTutorial, GrhPath & "BotonTutorial.jpg", GrhPath & "BotonTutorialRollover.jpg", GrhPath & "BotonTutorialClick.jpg", Me)
                                    
    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonSalirOpciones.jpg", GrhPath & "BotonSalirRolloverOpciones.jpg", GrhPath & "BotonSalirClickOpciones.jpg", Me)
                                    
    Set picCheckBox = LoadPicture(GrhPath & "CheckBoxOpciones.jpg")

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub LoadUserConfig()
    
    On Error GoTo LoadUserConfig_Err
    

    ' Load music config
    bMusicActivated = Audio.MusicActivated
    Slider1(0).Enabled = bMusicActivated
    
    If bMusicActivated Then
        imgChkMusica.Picture = picCheckBox
        
        Slider1(0).value = Audio.MusicVolume

    End If
    
    ' Load Sound config
    bSoundActivated = Audio.SoundActivated
    Slider1(1).Enabled = bSoundActivated
    
    If bSoundActivated Then
        imgChkSonidos.Picture = picCheckBox
        
        Slider1(1).value = Audio.SoundVolume

    End If
    
    ' Load Sound Effects config
    bSoundEffectsActivated = Audio.SoundEffectsActivated

    If bSoundEffectsActivated Then imgChkEfectosSonido.Picture = picCheckBox
    
    txtCantMensajes.Text = CStr(DialogosClanes.CantidadDialogos)
    
    If DialogosClanes.Activo Then
        imgChkPantalla.Picture = picCheckBox
    Else
        imgChkConsola.Picture = picCheckBox

    End If
    
    If ClientSetup.bGuildNews Then
        imgChkMostrarNews.Picture = picCheckBox
    Else
        imgChkNoMostrarNews.Picture = picCheckBox

    End If
        
    If ClientSetup.bKill Then imgChkRequiredLvl.Picture = picCheckBox
    If ClientSetup.bDie Then imgChkAlMorir.Picture = picCheckBox
    If Not ClientSetup.bActive Then imgChkDesactivarFragShooter.Picture = picCheckBox
    
    txtLevel = ClientSetup.byMurderedLevel

    
    Exit Sub

LoadUserConfig_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "LoadUserConfig"
    End If
Resume Next
    
End Sub

Private Sub Slider1_Change(Index As Integer)
    
    On Error GoTo Slider1_Change_Err
    

    Select Case Index

        Case 0
            Audio.MusicVolume = Slider1(0).value

        Case 1
            Audio.SoundVolume = Slider1(1).value

    End Select

    
    Exit Sub

Slider1_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "Slider1_Change"
    End If
Resume Next
    
End Sub

Private Sub Slider1_Scroll(Index As Integer)
    
    On Error GoTo Slider1_Scroll_Err
    

    Select Case Index

        Case 0
            Audio.MusicVolume = Slider1(0).value

        Case 1
            Audio.SoundVolume = Slider1(1).value

    End Select

    
    Exit Sub

Slider1_Scroll_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmOpciones" & "->" & "Slider1_Scroll"
    End If
Resume Next
    
End Sub
