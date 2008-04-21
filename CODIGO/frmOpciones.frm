VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOpciones 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4740
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
   ScaleHeight     =   5325
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Cambiar Contraseña"
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton cmdCustomKeys 
      Caption         =   "Configurar Teclas"
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton customMsgCmd 
      Caption         =   "Mensajes Personalizados"
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton cmdManual 
      Caption         =   "Manual de Argentum Online"
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Audio"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   4215
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Sonidos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Musica"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   10
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Diálogos de clan"
      ForeColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   255
      TabIndex        =   2
      Top             =   1665
      Width           =   4230
      Begin VB.TextBox txtCantMensajes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2925
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "5"
         Top             =   315
         Width           =   450
      End
      Begin VB.OptionButton optPantalla 
         BackColor       =   &H00000000&
         Caption         =   "En pantalla,"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1800
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optConsola 
         BackColor       =   &H00000000&
         Caption         =   "En consola"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   1560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mensajes"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3480
         TabIndex        =   6
         Top             =   345
         Width           =   750
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   960
      MouseIcon       =   "frmOpciones.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4800
      Width           =   2790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   180
      Width           =   2775
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

Private loading As Boolean

Private Sub Check1_Click(index As Integer)
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    
    Select Case index
        Case 0
            If Check1(0).value = vbUnchecked Then
                Audio.MusicActivated = False
                Slider1(0).Enabled = False
            ElseIf Not Audio.MusicActivated Then  'Prevent the music from reloading
                Audio.MusicActivated = True
                Slider1(0).Enabled = True
                Slider1(0).value = Audio.MusicVolume
            End If
        
        Case 1
            If Check1(1).value = vbUnchecked Then
                Audio.SoundActivated = False
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone
                Slider1(1).Enabled = False
            Else
                Audio.SoundActivated = True
                Slider1(1).Enabled = True
                Slider1(1).value = Audio.SoundVolume
            End If
    End Select
End Sub

Private Sub cmdCustomKeys_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub cmdManual_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://ao.alkon.com.ar/aomanual/", "", App.path, 0)
End Sub

Private Sub cmdChangePassword_Click()
    Call frmNewPassword.Show(vbModal, Me)
End Sub

Private Sub Command2_Click()
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub customMsgCmd_Click()
    Call frmMessageTxt.Show(vbModeless, Me)
End Sub

Private Sub Form_Load()
    loading = True      'Prevent sounds when setting check's values
    
    If Audio.MusicActivated Then
        Check1(0).value = vbChecked
        Slider1(0).Enabled = True
        Slider1(0).value = Audio.MusicVolume
    Else
        Check1(0).value = vbUnchecked
        Slider1(0).Enabled = False
    End If
    
    If Audio.SoundActivated Then
        Check1(1).value = vbChecked
        Slider1(1).Enabled = True
        Slider1(1).value = Audio.SoundVolume
    Else
        Check1(1).value = vbUnchecked
        Slider1(1).Enabled = False
    End If
    
    txtCantMensajes.Text = CStr(DialogosClanes.CantidadDialogos)
    
    If DialogosClanes.Activo Then
        optConsola.value = False
        optPantalla.value = True
    Else
        optConsola.value = True
        optPantalla.value = False
    End If
    
    loading = False     'Enable sounds when setting check's values
End Sub

Private Sub optConsola_Click()
    DialogosClanes.Activo = False
End Sub

Private Sub optPantalla_Click()
    DialogosClanes.Activo = True
End Sub

Private Sub Slider1_Change(index As Integer)
    Select Case index
        Case 0
            Audio.MusicVolume = Slider1(0).value
        Case 1
            Audio.SoundVolume = Slider1(1).value
    End Select
End Sub

Private Sub Slider1_Scroll(index As Integer)
    Select Case index
        Case 0
            Audio.MusicVolume = Slider1(0).value
        Case 1
            Audio.SoundVolume = Slider1(1).value
    End Select
End Sub

Private Sub txtCantMensajes_LostFocus()
    txtCantMensajes.Text = Trim$(txtCantMensajes.Text)
    If IsNumeric(txtCantMensajes.Text) Then
        DialogosClanes.CantidadDialogos = Trim$(txtCantMensajes.Text)
    Else
        txtCantMensajes.Text = 5
    End If
End Sub
