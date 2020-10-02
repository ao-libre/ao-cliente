VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online Libre"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   5.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin AOLibre.uAOButton btnSalir 
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":1DC21
      PICF            =   "frmConnect.frx":1E64B
      PICH            =   "frmConnect.frx":1F30D
      PICV            =   "frmConnect.frx":2029F
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
   Begin AOLibre.uAOButton btnCrearCuenta 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Crear Cuenta"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":211A1
      PICF            =   "frmConnect.frx":21BCB
      PICH            =   "frmConnect.frx":2288D
      PICV            =   "frmConnect.frx":2381F
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton btnTeclas 
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Teclas"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":24721
      PICF            =   "frmConnect.frx":2514B
      PICH            =   "frmConnect.frx":25E0D
      PICV            =   "frmConnect.frx":26D9F
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
   Begin AOLibre.uAOButton btnConectarse 
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Conectarse"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":27CA1
      PICF            =   "frmConnect.frx":286CB
      PICH            =   "frmConnect.frx":2938D
      PICV            =   "frmConnect.frx":2A31F
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
   Begin AOLibre.uAOCheckbox chkRecordar 
      Height          =   345
      Left            =   5280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4680
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmConnect.frx":2B221
   End
   Begin VB.TextBox txtPasswd 
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
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3720
      Width           =   2460
   End
   Begin VB.TextBox txtNombre 
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
      Left            =   4905
      TabIndex        =   0
      Top             =   3210
      Width           =   2460
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
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
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5760
      TabIndex        =   4
      Text            =   "190.245.145.3"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
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
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4890
      TabIndex        =   3
      Text            =   "7667"
      Top             =   2760
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v0.0.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   510
   End
   Begin VB.Label lblRecordarme 
      BackStyle       =   0  'Transparent
      Caption         =   "Recordarme"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   4800
      Width           =   2055
   End
End
Attribute VB_Name = "frmConnect"
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
'
'Matias Fernando Pequeno
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Codigo Postal 1405

Option Explicit

Private Lector As clsIniManager

Private Const AES_PASSWD As String = "tumamaentanga"

Private Sub btnConectarse_Click()
    'update user info
    AccountName = txtNombre.Text
    AccountPassword = txtPasswd.Text

    'Clear spell list
    frmMain.hlst.Clear

    If Me.chkRecordar.Checked = False Then
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Remember", "False")
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "UserName", vbNullString)
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Password", vbNullString)
    Else
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Remember", "True")
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "UserName", AccountName)
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Password", Cripto.AesEncryptString(AccountPassword, AES_PASSWD))
    End If

    If CheckUserData() = True Then
        Call Protocol.Connect(E_MODO.Normal)
    End If
End Sub

Private Sub btnSalir_Click()
    Call CloseClient
End Sub

Private Sub btnTeclas_Click()
    Load frmKeypad
    frmKeypad.Show vbModal
    Unload frmKeypad
    txtPasswd.SetFocus
End Sub

Private Sub Form_Activate()
    If Len(CurServerIp) Then IPTxt.Text = CurServerIp
    If CurServerPort Then PortTxt.Text = CurServerPort

    If CBool(GetVar(Game.path(INIT) & "Config.ini", "LOGIN", "Remember")) = True Then
        Me.txtNombre = GetVar(Game.path(INIT) & "Config.ini", "LOGIN", "UserName")
        Me.txtPasswd = Cripto.AesDecryptString(GetVar(Game.path(INIT) & "Config.ini", "LOGIN", "Password"), AES_PASSWD)
        Me.chkRecordar.Checked = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Call CloseClient
    End If
End Sub


Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]

    version.Caption = GetVersionOfTheGame() & " (clic aca para cambiar ip y puerto)"

    'Solo hay 2 imagenes de cargando, cambiar 2 por el numero maximo si se quiere cambiar
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaConectar" & RandomNumber(1, 2) & ".jpg")
    
    Call LoadTextsForm

End Sub

Private Sub LoadTextsForm()
    btnConectarse.Caption = JsonLanguage.item("BTN_CONECTARSE").item("TEXTO")
    btnCrearCuenta.Caption = JsonLanguage.item("BTN_CREAR_CUENTA").item("TEXTO")
    lblRecordarme.Caption = JsonLanguage.item("LBL_RECORDARME").item("TEXTO")
    btnSalir.Caption = JsonLanguage.item("BTN_SALIR").item("TEXTO")
    btnTeclas.Caption = JsonLanguage.item("LBL_TECLAS").item("TEXTO")
End Sub

Private Sub IPTxt_Change()
    CurServerIp = IPTxt.Text
End Sub

Private Sub PortTxt_Change()
    CurServerPort = Val(PortTxt.Text)
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then btnConectarse_Click
End Sub

Private Sub btnCrearCuenta_Click()
    Call ShellExecute(0, "Open", "https://www.argentum20.com", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub version_Click()
    IPTxt.Visible = True
    PortTxt.Visible = True
End Sub
