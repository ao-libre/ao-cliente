VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
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
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkRecordar 
      BackColor       =   &H80000000&
      Caption         =   "Recordar Cuenta"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Timer tEfectos 
      Left            =   1680
      Top             =   1080
   End
   Begin AOLibre.uAOButton btnActualizarLista 
      Height          =   495
      Left            =   9120
      TabIndex        =   7
      Top             =   6840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      TX              =   "Actualizar Lista"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":000C
      PICF            =   "frmConnect.frx":0A36
      PICH            =   "frmConnect.frx":16F8
      PICV            =   "frmConnect.frx":268A
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InetCtlsObjects.Inet InetIpApi 
      Left            =   840
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetReddit 
      Left            =   120
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox lstRedditPosts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   4320
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.ListBox lstServers 
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
      ForeColor       =   &H0000FF00&
      Height          =   4905
      ItemData        =   "frmConnect.frx":358C
      Left            =   8685
      List            =   "frmConnect.frx":358E
      TabIndex        =   6
      Top             =   1680
      Width           =   2775
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   3210
      Width           =   2460
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
      TabIndex        =   2
      Text            =   "7666"
      Top             =   2760
      Width           =   825
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
      TabIndex        =   3
      Text            =   "localhost"
      Top             =   2760
      Width           =   1575
   End
   Begin AOLibre.uAOButton btnCrearCuenta 
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Crear Cuenta"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":3590
      PICF            =   "frmConnect.frx":3FBA
      PICH            =   "frmConnect.frx":4C7C
      PICV            =   "frmConnect.frx":5C0E
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
   Begin AOLibre.uAOButton btnRecuperar 
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Recuperar Pass"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":6B10
      PICF            =   "frmConnect.frx":753A
      PICH            =   "frmConnect.frx":81FC
      PICV            =   "frmConnect.frx":918E
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
   Begin AOLibre.uAOButton btnManual 
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Manual"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":A090
      PICF            =   "frmConnect.frx":AABA
      PICH            =   "frmConnect.frx":B77C
      PICV            =   "frmConnect.frx":C70E
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
   Begin AOLibre.uAOButton btnReglamento 
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Reglamento"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":D610
      PICF            =   "frmConnect.frx":E03A
      PICH            =   "frmConnect.frx":ECFC
      PICV            =   "frmConnect.frx":FC8E
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
   Begin AOLibre.uAOButton btnCodigoFuente 
      Height          =   375
      Left            =   6840
      TabIndex        =   12
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Codigo Fuente"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":10B90
      PICF            =   "frmConnect.frx":115BA
      PICH            =   "frmConnect.frx":1227C
      PICV            =   "frmConnect.frx":1320E
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
   Begin AOLibre.uAOButton btnCrearServer 
      Height          =   375
      Left            =   8400
      TabIndex        =   13
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Crear Server"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":14110
      PICF            =   "frmConnect.frx":14B3A
      PICH            =   "frmConnect.frx":157FC
      PICV            =   "frmConnect.frx":1678E
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
   Begin AOLibre.uAOButton btnSalir 
      Height          =   375
      Left            =   9960
      TabIndex        =   14
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":17690
      PICF            =   "frmConnect.frx":180BA
      PICH            =   "frmConnect.frx":18D7C
      PICV            =   "frmConnect.frx":19D0E
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
      TabIndex        =   15
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Conectarse"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":1AC10
      PICF            =   "frmConnect.frx":1B63A
      PICH            =   "frmConnect.frx":1C2FC
      PICV            =   "frmConnect.frx":1D28E
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
   Begin AOLibre.uAOButton btnTeclas 
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Teclas"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":1E190
      PICF            =   "frmConnect.frx":1EBBA
      PICH            =   "frmConnect.frx":1F87C
      PICV            =   "frmConnect.frx":2080E
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
   Begin AOLibre.uAOButton btnVerForo 
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   6075
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      TX              =   "Ver Foro"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":21710
      PICF            =   "frmConnect.frx":2213A
      PICH            =   "frmConnect.frx":22DFC
      PICV            =   "frmConnect.frx":23D8E
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
      Caption         =   "Label1"
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
      TabIndex        =   0
      Top             =   240
      Width           =   555
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

Private Type tRedditPost
    Title As String
    URL As String
End Type

Dim Posts() As tRedditPost

' Animación de los Controles...
Private Type tAnimControl
    Activo As Boolean
    Velocidad As Double
    Top As Integer
End Type
Private AnimControl(1 To 11) As tAnimControl
Private Fuerza As Double

Public LastButtonPressed As clsGraphicalButton

Private Lector As ClsIniReader

Private Const AES_PASSWD As String = "tumamaentanga"

Private Function RefreshServerList() As String
'***************************************************
'Author: Recox
'Last Modification: 01/04/2019
'01/04/2019: Recox - Descarga y llena el listado de servers
'***************************************************
        Call DownloadServersFile("https://raw.githubusercontent.com/ao-libre/ao-cliente/master/INIT/sinfo.dat")
        Call CargarServidores
End Function

Private Sub btnActualizarLista_Click()
'***************************************************
'Author: Recox
'Last Modification: 01/04/2019
'01/04/2019: Recox - Boton para actualizar la lista de servers
'***************************************************
    frmConnect.lstServers.Clear
    frmConnect.lstServers.AddItem ("Actualizando Servers...")
    frmConnect.lstServers.AddItem ("Por Favor Espere")
    Call RefreshServerList
    MsgBox "Se actualizo con exito la lista de servers"
End Sub

Private Sub btnCodigoFuente_Click()
'***********************************
'IMPORTANTE!
'
'No debe eliminarse la posibilidad de bajar el codigo de sus servidor de esta forma.
'Caso contrario estarian violando la licencia Affero GPL y con ella derechos de autor,
'incurriendo de esta forma en un delito punible por ley.
'
'Argentum Online es libre, es de todos. Mantengamoslo asi. Si tanto te gusta el juego y queres los
'cambios que hacemos nosotros, comparti los tuyos. Es un cambio justo. Si no estas de acuerdo,
'no uses nuestro codigo, pues nadie te obliga o bien utiliza una version anterior a la 0.12.0.
'***********************************
    Call ShellExecute(0, "Open", "https://github.com/ao-libre", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub btnConectarse_Click()
    'update user info
    AccountName = txtNombre.Text
    AccountPassword = txtPasswd.Text

    'Clear spell list
    frmMain.hlst.Clear

    If CheckUserData() = True Then
        Call Protocol.Connect(E_MODO.Normal)
    End If
End Sub

Private Sub btnCrearServer_Click()
    Call ShellExecute(0, "Open", "https://www.reddit.com/r/argentumonlineoficial/comments/9dow3q/como_montar_mi_propio_servidor/", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub btnManual_Click()
    Call ShellExecute(0, "Open", "http://wiki.argentumonline.org", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub btnRecuperar_Click()
    Call Protocol.Connect(E_MODO.CambiarContrasena)
End Sub

Private Sub btnReglamento_Click()
    Call ShellExecute(0, "Open", "http://wiki.argentumonline.org/reglamento", "", App.Path, SW_SHOWNORMAL)
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

Private Sub btnVerForo_Click()
    Call ShellExecute(0, "Open", "https://www.reddit.com/r/argentumonlineoficial/", vbNullString, App.Path, SW_SHOWNORMAL)
End Sub

Private Sub chkRecordar_Click()
    If Me.chkRecordar.Value = vbUnchecked Then
        Call WriteVar(Path(INIT) & "Config.ini", "Login", "Remember", 0)
        Call WriteVar(Path(INIT) & "Config.ini", "Login", "UserName", vbNullString)
        Call WriteVar(Path(INIT) & "Config.ini", "Login", "Password", vbNullString)
    Else
        Call WriteVar(Path(INIT) & "Config.ini", "Login", "UserName", Me.txtNombre)
        Call WriteVar(Path(INIT) & "Config.ini", "Login", "Password", Cripto.AesEncryptString(Me.txtPasswd, AES_PASSWD))
        Call WriteVar(Path(INIT) & "Config.ini", "Login", "Remember", 1)
    End If
End Sub

Private Sub Form_Activate()
    
    If CurServer <> 0 Then
        IPTxt = ServersLst(1).Ip
        PortTxt = ServersLst(1).Puerto
    Else
        IPTxt = IPdelServidor
        PortTxt = PuertoDelServidor
    End If
    
    Set Lector = New ClsIniReader
    Lector.Initialize (Path(INIT) & "Config.ini")
    
    If Lector.GetValue("LOGIN", "Remember") = 1 Then
        Me.txtNombre = Lector.GetValue("LOGIN", "UserName")
        Me.txtPasswd = Cripto.AesDecryptString(Lector.GetValue("LOGIN", "Password"), AES_PASSWD)
        Me.chkRecordar.Value = 1
    End If
    
    Call GetPostsFromReddit
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

    Call RefreshServerList
    
    If CurServer <> 0 Then
        IPTxt = ServersLst(CurServer).Ip
        PortTxt = ServersLst(CurServer).Puerto
    Else
        IPTxt = ServersLst(1).Ip
        PortTxt = ServersLst(1).Puerto
    End If

    version.Caption = GetVersionOfTheGame()

    Me.Picture = LoadPicture(App.Path & "\graficos\VentanaConectar.jpg")
    
    Call LoadButtons

    ' GSZAO - Animación...
    
    'TODO: Agregar los movimientos faltantes, me aburri (Recox)
    'btnConectarse.Top = 10
    'AnimControl(1).Activo = True
    'AnimControl(1).Velocidad = 0
   ' AnimControl(1).Top = 200
    
    'btnActualizarLista.Top = 10
    'AnimControl(2).Activo = True
    'AnimControl(2).Velocidad = 0
    'AnimControl(2).Top = 350
    
    btnCodigoFuente.Top = 10
    AnimControl(3).Activo = True
    AnimControl(3).Velocidad = 0
    AnimControl(3).Top = 560
    
    btnCrearCuenta.Top = 10
    AnimControl(4).Activo = True
    AnimControl(4).Velocidad = 0
    AnimControl(4).Top = 560
    
    btnCrearServer.Top = 10
    AnimControl(5).Activo = True
    AnimControl(5).Velocidad = 0
    AnimControl(5).Top = 560
    
    btnManual.Top = 10
    AnimControl(6).Activo = True
    AnimControl(6).Velocidad = 0
    AnimControl(6).Top = 560
    
    btnRecuperar.Top = 10
    AnimControl(7).Activo = True
    AnimControl(7).Velocidad = 0
    AnimControl(7).Top = 560
    
    'btnReglamento.Top = 10
    AnimControl(8).Activo = True
    'AnimControl(8).Velocidad = 0
    'AnimControl(8).Top = 560

    'btnTeclas.Top = 10
    AnimControl(10).Activo = True
    'AnimControl(10).Velocidad = 0
    'AnimControl(10).Top = 560
    
    Fuerza = 1.7 ' Gravedad... 1.7
    tEfectos.Interval = 10
    tEfectos.Enabled = True

End Sub

Private Sub tEfectos_Timer()
    Dim oTop As Integer
    Dim i    As Integer

    For i = 1 To 7

        If AnimControl(i).Activo = True Then

            Select Case i

                Case 1: oTop = btnConectarse.Top

                Case 2: oTop = btnActualizarLista.Top

                Case 3: oTop = btnCodigoFuente.Top

                Case 4: oTop = btnCrearCuenta.Top

                Case 5: oTop = btnCrearServer.Top

                Case 6: oTop = btnManual.Top

                Case 7: oTop = btnRecuperar.Top

                Case 8: oTop = btnReglamento.Top

                Case 9: oTop = btnSalir.Top

                Case 10: oTop = btnTeclas.Top

                Case 11: oTop = btnVerForo.Top
            End Select

            If oTop > AnimControl(i).Top Then
                oTop = AnimControl(i).Top
                AnimControl(i).Velocidad = AnimControl(i).Velocidad * -0.6
            End If

            If AnimControl(i).Velocidad >= -0.6 And AnimControl(i).Velocidad <= -0.5 Then
                AnimControl(i).Activo = False
            Else
                AnimControl(i).Velocidad = AnimControl(i).Velocidad + Fuerza
                oTop = oTop + AnimControl(i).Velocidad
            End If

            Select Case i

                Case 1: btnActualizarLista.Top = oTop

                Case 2: btnConectarse.Top = oTop

                Case 3: btnCodigoFuente.Top = oTop

                Case 4: btnCrearCuenta.Top = oTop

                Case 5: btnCrearServer.Top = oTop

                Case 6: btnManual.Top = oTop

                Case 7: btnRecuperar.Top = oTop

                Case 8: btnReglamento.Top = oTop

                Case 9: btnSalir.Top = oTop

                Case 10: btnTeclas.Top = oTop

                Case 11: btnVerForo.Top = oTop
            End Select
        End If
    Next

    If AnimControl(1).Activo = False And AnimControl(2).Activo = False And AnimControl(3).Activo = False And AnimControl(4).Activo = False And AnimControl(5).Activo = False And AnimControl(6).Activo = False And AnimControl(7).Activo = False And AnimControl(8).Activo = False And AnimControl(9).Activo = False And AnimControl(10).Activo = False And AnimControl(11).Activo = False Then
        tEfectos.Enabled = False
        
        ' GSZAO - Animaciï¿½n...
        btnConectarse.Top = AnimControl(1).Top
        btnActualizarLista.Top = AnimControl(2).Top
        btnCodigoFuente.Top = AnimControl(3).Top
        btnCrearCuenta.Top = AnimControl(4).Top
        btnCrearServer.Top = AnimControl(5).Top
        btnManual.Top = AnimControl(6).Top
        btnRecuperar.Top = AnimControl(7).Top
        btnReglamento.Top = AnimControl(8).Top
        btnSalir.Top = AnimControl(9).Top
        btnTeclas.Top = AnimControl(10).Top
        btnVerForo.Top = AnimControl(11).Top
    End If
    
End Sub

Private Sub LoadButtons()
    btnActualizarLista.Caption = JsonLanguage.Item("BTN_ACTUALIZAR_LISTA").Item("TEXTO")
    btnCodigoFuente.Caption = JsonLanguage.Item("BTN_CODIGO_FUENTE").Item("TEXTO")
    btnConectarse.Caption = JsonLanguage.Item("BTN_CONECTARSE").Item("TEXTO")
    btnCrearCuenta.Caption = JsonLanguage.Item("BTN_CREAR_CUENTA").Item("TEXTO")
    btnCrearServer.Caption = JsonLanguage.Item("BTN_CREAR_SERVER").Item("TEXTO")
    btnManual.Caption = JsonLanguage.Item("BTN_MANUAL").Item("TEXTO")
    btnRecuperar.Caption = JsonLanguage.Item("BTN_RECUPERAR").Item("TEXTO")
    btnReglamento.Caption = JsonLanguage.Item("BTN_REGLAMENTO").Item("TEXTO")
End Sub

Private Sub lstRedditPosts_Click()
    Call ShellExecute(0, "Open", Posts(lstRedditPosts.ListIndex + 1).URL, "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub lstServers_Click()
    IPTxt.Text = ServersLst(lstServers.ListIndex + 1).Ip
    PortTxt.Text = ServersLst(lstServers.ListIndex + 1).Puerto
    
    'Variable Global declarada en Declares.bas
    MundoSeleccionado = ServersLst(lstServers.ListIndex + 1).Mundo
    
    'En caso que no haya un mundo seleccionado en la propiedad Mundo
    'Seleccionamos Alkon como mundo default
    If LenB(MundoSeleccionado) = 0 Then
        MundoSeleccionado = "Alkon"
    End If

    CurServer = lstServers.ListIndex + 1
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then btnConectarse_Click
End Sub

Private Sub GetPostsFromReddit()
On Error Resume Next

    Dim ResponseReddit As String
    Dim JsonObject As Object
    Dim Endpoint As String
    
    Endpoint = GetVar(Path(INIT) & "Config.ini", "Parameters", "SubRedditEndpoint")
    ResponseReddit = InetReddit.OpenURL(Endpoint)
    Set JsonObject = JSON.parse(ResponseReddit)
    
    Dim qtyPostsOnReddit As Integer
    qtyPostsOnReddit = JsonObject.Item("data").Item("children").Count
    ReDim Posts(qtyPostsOnReddit)
    
    'Clear lstRedditPosts before populate it again to prevent repeated values.
    lstRedditPosts.Clear
    
    'Long funciona mas rapido en los loops que Integer
    Dim i As Long
    i = 1
    Do While i <= qtyPostsOnReddit
        Posts(i).Title = JsonObject.Item("data").Item("children").Item(i).Item("data").Item("title")
        Posts(i).URL = JsonObject.Item("data").Item("children").Item(i).Item("data").Item("url")
        
        lstRedditPosts.AddItem JsonObject.Item("data").Item("children").Item(i).Item("data").Item("title")
        
        i = i + 1
    Loop
End Sub

Private Sub btnCrearCuenta_Click()
    Call Protocol.Connect(E_MODO.CrearCuenta)
End Sub
