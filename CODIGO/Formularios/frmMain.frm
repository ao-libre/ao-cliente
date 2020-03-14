VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   360
   ClientTop       =   -3300
   ClientWidth     =   15345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   DrawStyle       =   6  'Inside Solid
   FillColor       =   &H00008080&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00004080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   768
   ScaleMode       =   0  'User
   ScaleWidth      =   1023
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer timerTiempoRestanteInvisibleMensaje 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   2880
   End
   Begin VB.Timer timerTiempoRestanteParalisisMensaje 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   2880
   End
   Begin AOLibre.uAOProgress uAOProgressExperienceLevel 
      Height          =   255
      Left            =   12480
      TabIndex        =   40
      ToolTipText     =   "Experiencia necesaria para pasar de nivel"
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      Min             =   1
      Value           =   55
      Animate         =   0   'False
      UseBackground   =   0   'False
      ForeColor       =   -2147483624
      BackColor       =   4210816
      BorderColor     =   0
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   14520
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   21
      Top             =   10965
      Width           =   420
   End
   Begin VB.PictureBox MiniMapa 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   10080
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   300
      Width           =   1500
      Begin VB.Shape UserAreaMinimap 
         BackColor       =   &H80000004&
         BorderColor     =   &H80000002&
         FillColor       =   &H000080FF&
         Height          =   225
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Shape UserM 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         Height          =   45
         Left            =   705
         Top             =   705
         Width           =   45
      End
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   14115
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   29
      Top             =   10965
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   13710
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   24
      Top             =   10965
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   13290
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   23
      Top             =   10965
      Width           =   420
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   12360
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   16
      Top             =   2400
      Width           =   2400
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   240
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "frmMain.frx":418A7
      ToolTipText     =   "Chat"
      Top             =   2250
      Visible         =   0   'False
      Width           =   11415
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   9720
      Top             =   2880
   End
   Begin VB.TextBox SendCMSTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "frmMain.frx":418D7
      ToolTipText     =   "Chat"
      Top             =   10680
      Visible         =   0   'False
      Width           =   11490
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   9120
      Top             =   2880
   End
   Begin VB.Timer SonidosMapas 
      Interval        =   20000
      Left            =   8280
      Top             =   2880
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1485
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   300
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   2619
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":4190D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox hlst 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   12240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4800
      Width           =   2565
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   8880
      Left            =   120
      MousePointer    =   99  'Custom
      ScaleHeight     =   592
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   768
      TabIndex        =   30
      Top             =   2160
      Width           =   11520
      Begin VB.Timer trainingMacro 
         Enabled         =   0   'False
         Interval        =   3200
         Left            =   10680
         Top             =   600
      End
   End
   Begin AOLibre.uAOButton imgMapa 
      Height          =   255
      Left            =   13680
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   9270
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      TX              =   "Mapa"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":4198A
      PICF            =   "frmMain.frx":423B4
      PICH            =   "frmMain.frx":43076
      PICV            =   "frmMain.frx":44008
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
   Begin AOLibre.uAOButton imgGrupo 
      Height          =   255
      Left            =   13680
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   9570
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      TX              =   "Grupo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":44F0A
      PICF            =   "frmMain.frx":45934
      PICH            =   "frmMain.frx":465F6
      PICV            =   "frmMain.frx":47588
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
   Begin AOLibre.uAOButton imgOpciones 
      Height          =   255
      Left            =   13680
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9870
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      TX              =   "Opciones"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":4848A
      PICF            =   "frmMain.frx":48EB4
      PICH            =   "frmMain.frx":49B76
      PICV            =   "frmMain.frx":4AB08
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
   Begin AOLibre.uAOButton imgEstadisticas 
      Height          =   255
      Left            =   13680
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   10200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      TX              =   "Estadisticas"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":4BA0A
      PICF            =   "frmMain.frx":4C434
      PICH            =   "frmMain.frx":4D0F6
      PICV            =   "frmMain.frx":4E088
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
   Begin AOLibre.uAOButton imgClanes 
      Height          =   255
      Left            =   13680
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   10560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      TX              =   "Clanes"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":4EF8A
      PICF            =   "frmMain.frx":4F9B4
      PICH            =   "frmMain.frx":50676
      PICV            =   "frmMain.frx":51608
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
   Begin AOLibre.uAOButton CmdInventario 
      Height          =   375
      Left            =   12240
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      TX              =   "Inventario"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":5250A
      PICF            =   "frmMain.frx":52F34
      PICH            =   "frmMain.frx":53BF6
      PICV            =   "frmMain.frx":54B88
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
   Begin AOLibre.uAOButton CmdHechizos 
      Height          =   375
      Left            =   13560
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      TX              =   "Hechizos"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":55A8A
      PICF            =   "frmMain.frx":564B4
      PICH            =   "frmMain.frx":57176
      PICV            =   "frmMain.frx":58108
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
   Begin AOLibre.uAOButton CmdLanzar 
      Height          =   495
      Left            =   12000
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      TX              =   "Lanzar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":5900A
      PICF            =   "frmMain.frx":59A34
      PICH            =   "frmMain.frx":5A6F6
      PICV            =   "frmMain.frx":5B688
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
   Begin AOLibre.uAOButton cmdInfo 
      Height          =   495
      Left            =   13800
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      TX              =   "Info"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":5C58A
      PICF            =   "frmMain.frx":5CFB4
      PICH            =   "frmMain.frx":5DC76
      PICV            =   "frmMain.frx":5EC08
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
   Begin VB.Label lblPorcLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "33.33%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   14040
      TabIndex        =   42
      Top             =   1260
      Width           =   555
   End
   Begin VB.Label lblMapName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del mapa"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   9000
      TabIndex        =   39
      Top             =   1875
      Width           =   2535
   End
   Begin VB.Image imgAsignarSkill 
      Height          =   450
      Left            =   13320
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   450
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13920
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   14580
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   14880
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   45
      Width           =   495
   End
   Begin VB.Label lblFPS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "101"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   6480
      TabIndex        =   20
      Top             =   60
      Width           =   795
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   0
      Left            =   14790
      MousePointer    =   99  'Custom
      Top             =   3360
      MouseIcon       =   "frmMain.frx":5FB0A
      Picture         =   "frmMain.frx":5FC5C
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   1
      Left            =   14790
      MousePointer    =   99  'Custom
      Top             =   3105
      MouseIcon       =   "frmMain.frx":5FFA0
      Picture         =   "frmMain.frx":600F2
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image xz 
      Height          =   255
      Index           =   0
      Left            =   14760
      Top             =   -120
      Width           =   255
   End
   Begin VB.Image xzz 
      Height          =   195
      Index           =   1
      Left            =   14805
      Top             =   -120
      Width           =   225
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del pj"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   12120
      TabIndex        =   41
      Top             =   360
      Width           =   2985
   End
   Begin VB.Label lblLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   13305
      TabIndex        =   19
      Top             =   1260
      Width           =   270
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      ForeColor       =   &H80000018&
      Height          =   225
      Left            =   12705
      TabIndex        =   18
      Top             =   1260
      Width           =   465
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0080FFFF&
      Height          =   210
      Left            =   14400
      TabIndex        =   15
      Top             =   8760
      Width           =   90
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   13080
      TabIndex        =   9
      Top             =   8760
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   12480
      TabIndex        =   8
      Top             =   8745
      Width           =   210
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000 X:00 Y: 00"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   12240
      TabIndex        =   7
      Top             =   11085
      Width           =   975
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000/000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   11160
      Width           =   855
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   9270
      TabIndex        =   5
      Top             =   11160
      Width           =   855
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2940
      TabIndex        =   4
      Top             =   11160
      Width           =   855
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   1170
      TabIndex        =   3
      Top             =   11160
      Width           =   855
   End
   Begin VB.Image imgScroll 
      Height          =   240
      Index           =   1000
      Left            =   14760
      MousePointer    =   99  'Custom
      Top             =   3105
      Width           =   225
   End
   Begin VB.Image InvEqu 
      Height          =   6255
      Left            =   12075
      Top             =   1800
      Width           =   2970
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   12120
      TabIndex        =   11
      Top             =   9660
      Width           =   1095
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   12120
      TabIndex        =   10
      Top             =   9300
      Width           =   1095
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   12120
      TabIndex        =   12
      Top             =   9990
      Width           =   1095
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   12120
      TabIndex        =   13
      Top             =   10320
      Width           =   1095
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   12120
      TabIndex        =   14
      Top             =   10650
      Width           =   1095
   End
   Begin VB.Shape shpEnergia 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   12120
      Top             =   9315
      Width           =   1125
   End
   Begin VB.Shape shpMana 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   12120
      Top             =   9660
      Width           =   1125
   End
   Begin VB.Shape shpVida 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   12120
      Top             =   9990
      Width           =   1125
   End
   Begin VB.Shape shpHambre 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   12120
      Top             =   10335
      Width           =   1125
   End
   Begin VB.Shape shpSed 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   12120
      Top             =   10650
      Width           =   1125
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmMain
'    Project    : ARGENTUM
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
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

Option Explicit

Public TX                  As Byte
Public TY                  As Byte
Public MouseX              As Long
Public MouseY              As Long
Public MouseBoton          As Long
Public MouseShift          As Long
Private clicX              As Long
Private clicY              As Long

Public IsPlaying           As Byte

Private clsFormulario      As clsFormMovementManager

Private cBotonDiamArriba   As clsGraphicalButton
Private cBotonDiamAbajo    As clsGraphicalButton
Private cBotonAsignarSkill As clsGraphicalButton

Public LastButtonPressed   As clsGraphicalButton

Public picSkillStar        As Picture

Public WithEvents Client   As clsSocket
Attribute Client.VB_VarHelpID = -1

Private ChangeHechi        As Boolean, ChangeHechiNum As Integer

Private FirstTimeChat      As Boolean
Private FirstTimeClanChat  As Boolean

'Usado para controlar que no se dispare el binding de la tecla CTRL cuando se usa CTRL+Tecla.
Dim CtrlMaskOn             As Boolean
Dim SkinSeleccionado       As String

Private Const NEWBIE_USER_GOLD_COLOR As Long = vbCyan
Private Const USER_GOLD_COLOR As Long = vbYellow

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call Protocol.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
End Sub

Private Sub Form_Activate()

    Call Inventario.DrawInventory

End Sub

Private Sub Form_Load()
    SkinSeleccionado = GetVar(Game.path(INIT) & "Config.ini", "Parameters", "SkinSelected")
    
    If Not ResolucionCambiada Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        Call clsFormulario.Initialize(Me, 120)
    End If
        
    Call LoadButtons
    
    With Me
        'Lo hardcodeo porque de lo contrario se ve un borde blanco.
        .Height = 11550
        .Label6 = JsonLanguage.item("NIVEL").item("TEXTO") & ": "
    End With

    Call LoadTextsForm
    'Call LoadAOCustomControlsPictures(Me)
    'Todo: Poner la carga de botones como en el frmCambiaMotd.frm para mantener coherencia con el resto de la aplicacion
    'y poder borrar los frx de este archivo
        
    ' Detect links in console
    Call EnableURLDetect(RecTxt.hWnd, Me.hWnd)
    
    ' Make the console transparent
    Call SetWindowLong(RecTxt.hWnd, -20, &H20&)
    
    CtrlMaskOn = False
    
    FirstTimeChat = True
    FirstTimeClanChat = True
    
End Sub

Private Sub LoadTextsForm()
    CmdLanzar.Caption = JsonLanguage.item("LBL_LANZAR").item("TEXTO")
    CmdInventario.Caption = JsonLanguage.item("LBL_INVENTARIO").item("TEXTO")
    CmdHechizos.Caption = JsonLanguage.item("LBL_HECHIZOS").item("TEXTO")
    cmdInfo.Caption = JsonLanguage.item("LBL_INFO").item("TEXTO")
    imgMapa.Caption = JsonLanguage.item("LBL_MAPA").item("TEXTO")
    imgGrupo.Caption = JsonLanguage.item("LBL_GRUPO").item("TEXTO")
    imgOpciones.Caption = JsonLanguage.item("LBL_OPCIONES").item("TEXTO")
    imgEstadisticas.Caption = JsonLanguage.item("LBL_ESTADISTICAS").item("TEXTO")
    imgClanes.Caption = JsonLanguage.item("LBL_CLANES").item("TEXTO")
End Sub

Private Sub LoadButtons()
    Dim i As Integer

    Set cBotonDiamArriba = New clsGraphicalButton
    Set cBotonDiamAbajo = New clsGraphicalButton
    Set cBotonAsignarSkill = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton

    Set picSkillStar = LoadPicture(Game.path(Skins) & SkinSeleccionado & "\BotonAsignarSkills.bmp")

    If SkillPoints > 0 Then imgAsignarSkill.Picture = picSkillStar
    
    imgAsignarSkill.MouseIcon = picMouseIcon
    lblDropGold.MouseIcon = picMouseIcon
    lblCerrar.MouseIcon = picMouseIcon
    lblMinimizar.MouseIcon = picMouseIcon
    
    For i = 0 To 2
        picSM(i).MouseIcon = picMouseIcon
    Next i

End Sub

Public Sub LightSkillStar(ByVal bTurnOn As Boolean)

    If bTurnOn Then
        imgAsignarSkill.Picture = picSkillStar
    Else
        Set imgAsignarSkill.Picture = Nothing
    End If
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)

    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub
        Dim sTemp As String
    
        Select Case Index

            Case 1 'subir

                If hlst.ListIndex = 0 Then Exit Sub

            Case 0 'bajar

                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        End Select
    
        Call WriteMoveSpell(Index = 1, hlst.ListIndex + 1)
        
        Select Case Index

            Case 1 'subir
                sTemp = hlst.List(hlst.ListIndex - 1)
                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex - 1

            Case 0 'bajar
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1
        End Select
    End If
End Sub

Public Sub ControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)
    
    Dim GrhIndex As Long
    Dim DR       As RECT

    GrhIndex = GRH_INI_SM + Index + SM_CANT * (CInt(Mostrar) + 1)

    With GrhData(GrhIndex)
    
        DR.Left = 0
        DR.Top = 0
        DR.Right = .pixelWidth
        DR.Bottom = .pixelHeight
        
    End With

    Call DrawGrhtoHdc(picSM(Index), GrhIndex, DR)
    
    Select Case Index
        
        Case eSMType.sResucitation
            
            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("COLOR").item(3), _
                                     True, False, True)
                                        
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("TEXTO")
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_RESU_OFF").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_OFF").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_OFF").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_OFF").item("COLOR").item(3), _
                                     True, False, True)
                                        
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("TEXTO")
            End If
            
        Case eSMType.sSafemode
            
            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, UCase$(JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("TEXTO").item(1)), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(3), _
                                      True, False, True)
                                        
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("TEXTO").item(2)
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, UCase$(JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("TEXTO").item(1)), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(3), _
                                      True, False, True)
                                        
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("TEXTO").item(2)
            End If
            
        Case eSMType.mWork
            
            If Mostrar Then
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_MACRO_ACTIVADO").item("TEXTO")
            Else
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_MACRO_DESACTIVADO").item("TEXTO")
            End If
            
    End Select

    SMStatus(Index) = Mostrar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 18/11/2010
    '18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
    '18/11/2010: Amraphen - Agregue el handle correspondiente para las nuevas configuraciones de teclas (CTRL+0..9).
    '***************************************************
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
    
        'Verificamos si se esta presionando la tecla CTRL.
        If Shift = 2 Then
            
            If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
                If KeyCode = vbKey0 Then
                    'Si es CTRL+0 muestro la ventana de configuracion de teclas.
                    Call frmCustomKeys.Show(vbModal, Me)
                    
                ElseIf KeyCode >= vbKey1 And KeyCode <= vbKey9 Then

                    'Si es CTRL+1..9 cambio la configuracion.
                    If KeyCode - vbKey0 = CustomKeys.CurrentConfig Then Exit Sub
                    
                    CustomKeys.CurrentConfig = KeyCode - vbKey0
                    
                    Dim sMsg As String
                    sMsg = JsonLanguage.item("CUSTOMKEYS_CONFIG_CARGADA").item("TEXTO")
                        
                    If CustomKeys.CurrentConfig = 0 Then
                        sMsg = Replace$(sMsg, "VAR_CONFIG_ELEGIDA", JsonLanguage.item("PREDETERMINADA").item("TEXTO"))
                    Else
                        sMsg = Replace$(sMsg, "VAR_CONFIG_ELEGIDA", JsonLanguage.item("PERSONALIZADA").item("TEXTO"))
                        sMsg = Replace$(sMsg, "VAR_CONFIG_CUSTOM_NUMERO", CStr(CustomKeys.CurrentConfig))
                    End If

                    Call ShowConsoleMsg(sMsg, JsonLanguage.item("CUSTOMKEYS_CONFIG_CARGADA").item("COLOR").item(1), _
                                              JsonLanguage.item("CUSTOMKEYS_CONFIG_CARGADA").item("COLOR").item(2), _
                                              JsonLanguage.item("CUSTOMKEYS_CONFIG_CARGADA").item("COLOR").item(3), _
                                        True)
                                        
                End If
                
                CtrlMaskOn = True
                
                Exit Sub
                
            End If
            
        End If
        
        If KeyCode = vbKeyControl Then

            'Chequeo que no se haya usado un CTRL + tecla antes de disparar las bindings.
            If CtrlMaskOn Then
                CtrlMaskOn = False
                Exit Sub
            End If
        End If
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then

            Select Case KeyCode

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    Audio.MusicActivated = Not Audio.MusicActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    Audio.SoundActivated = Not Audio.SoundActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFxs)
                    Audio.SoundEffectsActivated = Not Audio.SoundEffectsActivated
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)

                    If trainingMacro.Enabled Then Call DesactivarMacroHechizos
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)

                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call WriteSafeToggle

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle
                    
            End Select
        Else
            
            'Evito que se muestren los mensajes personalizados cuando se cambie una configuracion de teclas.
            If Shift = 2 Then Exit Sub
            
            Select Case KeyCode
            
                    'Custom messages!
                Case vbKey0 To vbKey9
                    Dim CustomMessage As String
                    
                    CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)

                    If LenB(CustomMessage) <> 0 Then

                        ' No se pueden mandar mensajes personalizados de clan o privado!
                        If UCase$(Left$(CustomMessage, 5)) <> "/CMSG" And Left$(CustomMessage, 1) <> "\" Then
                            
                            Call ParseUserCommand(CustomMessage)
                        End If
                    End If
                    
            End Select
            
        End If
        
    End If
    
    Select Case KeyCode

        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)

            If SendTxt.Visible Then Exit Sub
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not MirandoForo) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendCMSTXT.Visible = True
                SendCMSTXT.SetFocus
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call ScreenCapture
                
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)

            If UserMinMAN = UserMaxMAN Then Exit Sub
            
            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            Call WriteMeditate
        
        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)

            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If trainingMacro.Enabled Then
                DesactivarMacroHechizos
            Else
                ActivarMacroHechizos
            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)

            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If macrotrabajo.Enabled Then
                Call DesactivarMacroTrabajo
            Else
                Call ActivarMacroTrabajo
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)

            If frmMain.macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteQuit
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)

            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
            Else

                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
            End If
            
            If trainingMacro.Enabled Then Call DesactivarMacroHechizos
            If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
            
            If frmCustomKeys.Visible Then Exit Sub 'Chequeo si esta visible la ventana de configuracion de teclas.
            
            Call WriteAttack
            
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)

            If SendCMSTXT.Visible Then Exit Sub
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not MirandoForo) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
            
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call DisableURLDetect
    
End Sub

Private Sub GldLbl_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_GOLD_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_GOLD_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_GOLD_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_GOLD_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub imgAsignarSkill_Click()
    Dim i As Integer
    
    LlegaronSkills = False
    Call WriteRequestSkills
    Call FlushBuffer
    
    Do While Not LlegaronSkills
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    LlegaronSkills = False
    
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    
    Alocados = SkillPoints
    frmSkills3.puntos.Caption = SkillPoints
    frmSkills3.Show , frmMain

End Sub

Private Sub imgClanes_Click()
    
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub imgEstadisticas_Click()

    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
    Call WriteRequestFame
    Call FlushBuffer

    Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    frmEstadisticas.Iniciar_Labels
    frmEstadisticas.Show , frmMain
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
End Sub

Private Sub imgGrupo_Click()
    
    Call WriteRequestPartyForm
End Sub

Private Sub imgMapa_Click()
    
    Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub imgOpciones_Click()
    
    Call frmOpciones.Show(vbModeless, frmMain)
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub lblScroll_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub lblArmor_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_ARMOR_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_ARMOR_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_ARMOR_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_ARMOR_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmCerrar.Show vbModal, Me
End Sub


Private Sub lblDext_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_DEXT_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_DEXT_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_DEXT_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_DEXT_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblEnergia_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_ENERGIA_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_ENERGIA_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_ENERGIA_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_ENERGIA_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub



Private Sub lblHambre_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_HAMBRE_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_HAMBRE_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_HAMBRE_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_HAMBRE_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblHelm_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_HELM_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_HELM_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_HELM_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_HELM_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblMana_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_MANA_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_MANA_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_MANA_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_MANA_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblSed_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SED_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_SED_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_SED_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_SED_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblShielder_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SHIELDER_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_SHIELDER_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_SHIELDER_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_SHIELDER_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblStrg_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_STRG_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_STRG_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_STRG_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_STRG_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblVida_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_VIDA_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_VIDA_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_VIDA_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_VIDA_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblMinimizar_Click()
    Me.WindowState = 1
End Sub


Private Sub lblWeapon_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_WEAPON_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_WEAPON_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_WEAPON_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_WEAPON_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub macrotrabajo_Timer()

    If Inventario.SelectedItem = 0 Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    'If Not Application.IsAppActive() Then  'Implemento lo propuesto por GD, se puede usar macro aun que se este en otra ventana
    '    Call DesactivarMacroTrabajo
    '    Exit Sub
    'End If
    
    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not MirandoHerreria) Then
        Call WriteWorkLeftClick(TX, TY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
    If Not MirandoCarpinteria Then Call UsarItem
End Sub

Public Sub ActivarMacroTrabajo()
    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_MACRO_ACTIVADO").item("TEXTO"), 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mWork, True)
End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_MACRO_DESACTIVADO").item("TEXTO"), 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mWork, False)
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(TX, TY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(TX, TY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub PicMH_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(3), _
                        False, False, True)
End Sub

Private Sub Coord_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub picSM_DblClick(Index As Integer)

    Select Case Index

        Case eSMType.sResucitation
            Call WriteResuscitationToggle
        
        Case eSMType.sSafemode
            Call WriteSafeToggle
        
        Case eSMType.mWork

            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If

            If macrotrabajo.Enabled Then
                Call DesactivarMacroTrabajo
            Else
                Call ActivarMacroTrabajo
            End If
        
        Case eSMType.mSpells

            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If
        
            If trainingMacro.Enabled Then
                Call DesactivarMacroHechizos
            Else
                Call ActivarMacroHechizos
            End If
    End Select
End Sub

Private Sub RecTxt_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    StartCheckingLinks
End Sub

Private Sub SendCMSTXT_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Para borrar el mensaje del chat de clanes
    If FirstTimeClanChat Then
        SendCMSTXT.Text = vbNullString
        FirstTimeClanChat = False
        ' Color original
        SendCMSTXT.ForeColor = &H80000018
    End If
End Sub

Private Sub SendTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Para borrar el mensaje de fondo
    If FirstTimeChat Then
        SendTxt.Text = vbNullString
        FirstTimeChat = False
        ' Cambiamos el color de texto al original
        SendTxt.ForeColor = &HE0E0E0
    End If
    
    ' Control + Shift
    If Shift = 3 Then
        On Error GoTo errhandler
        
        ' Only allow numeric keys
        If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
            
            ' Get Msg Number
            Dim NroMsg As Integer
            NroMsg = KeyCode - vbKey0 - 1
            
            ' Pressed "0", so Msg Number is 9
            If NroMsg = -1 Then NroMsg = 9
            
            'Como es KeyDown, si mantenes _
             apretado el mensaje llena la consola

            If CustomMessages.Message(NroMsg) = SendTxt.Text Then
                Exit Sub
            End If
            
            CustomMessages.Message(NroMsg) = SendTxt.Text
            
            Dim MENSAJE_PERSONALIZADO As String
                MENSAJE_PERSONALIZADO = JsonLanguage.item("MENSAJE_PERSONALIZADO").item("TEXTO")
                MENSAJE_PERSONALIZADO = Replace$(MENSAJE_PERSONALIZADO, "VAR_MENSAJE", SendTxt.Text)
                MENSAJE_PERSONALIZADO = Replace$(MENSAJE_PERSONALIZADO, "VAR_MENSAJE_NUMERO", NroMsg + 1)
            
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg(MENSAJE_PERSONALIZADO, .Red, .Green, .Blue, .bold, .italic)
            End With
            
        End If
        
    End If
    
    Exit Sub
    
errhandler:

    'Did detected an invalid message??
    If Err.number = CustomMessages.InvalidMessageErrCode Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_CUSTOM_INVALIDO").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
        End With
    End If
    
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = vbNullString
        SendTxt.Text = vbNullString
        KeyCode = 0
        SendTxt.Visible = False
        
        If picInv.Visible Then
            picInv.SetFocus
        Else
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub Second_Timer()

    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()

    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
    Else

        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
            Else

                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    If Not Comerciando Then frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()

    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem()

    If pausa Then Exit Sub
    
    If Comerciando Then Exit Sub
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        Call WriteUseItem(Inventario.SelectedItem)
    End If
    
End Sub

Private Sub EquiparItem()

    If UserEstado = 1 Then
    
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
        
    Else
    
        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
            Call WriteEquipItem(Inventario.SelectedItem)
        End If
        
    End If
End Sub

Private Sub cmdLanzar_Click()
    
    If hlst.List(hlst.ListIndex) <> JsonLanguage.item("NADA").item("TEXTO") And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub cmdINFO_Click()
    
    If hlst.ListIndex <> -1 Then
        Dim Index As Integer
        Index = DevolverIndexHechizo(hlst.List(hlst.ListIndex))
        Dim Msj As String
     
        If Index <> 0 Then Msj = "%%%%%%%%%%%% " & JsonLanguage.item("MENSAJE_INFO_HECHIZOS").item("TEXTO").item(1) & " %%%%%%%%%%%%" & vbCrLf & JsonLanguage.item("MENSAJE_INFO_HECHIZOS").item("TEXTO").item(2) & ": " & Hechizos(Index).Nombre & vbCrLf & JsonLanguage.item("MENSAJE_INFO_HECHIZOS").item("TEXTO").item(3) & ": " & Hechizos(Index).Desc & vbCrLf & JsonLanguage.item("MENSAJE_INFO_HECHIZOS").item("TEXTO").item(4) & ": " & Hechizos(Index).SkillRequerido & vbCrLf & JsonLanguage.item("MENSAJE_INFO_HECHIZOS").item("TEXTO").item(5) & ": " & Hechizos(Index).ManaRequerida & vbCrLf & JsonLanguage.item("MENSAJE_INFO_HECHIZOS").item("TEXTO").item(6) & ": " & Hechizos(Index).EnergiaRequerida & vbCrLf & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
                                             
        Call ShowConsoleMsg(Msj, JsonLanguage.item("MENSAJE_INFO_HECHIZOS").item("COLOR").item(1), JsonLanguage.item("MENSAJE_INFO_HECHIZOS").item("COLOR").item(2), JsonLanguage.item("MENSAJE_INFO_HECHIZOS").item("COLOR").item(3))
        
    End If
End Sub

Private Sub DespInv_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub MainViewPic_DblClick()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteDoubleClick(TX, TY)
    End If
End Sub

Private Sub SendTxt_Click()
    SendTxt.Tag = 0 ' GSZAO
End Sub


Private Sub MainViewPic_Click()

    If Cartel Then Cartel = False
    
    Dim MENSAJE_ADVERTENCIA As String
    Dim VAR_LANZANDO        As String
    
    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, TX, TY)
        
        If Not InGameArea() Then Exit Sub
        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then

                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1

                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If

                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(TX, TY)
                Else

                    If trainingMacro.Enabled Then Call DesactivarMacroHechizos
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0

                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            VAR_LANZANDO = JsonLanguage.item("PROYECTILES").item("TEXTO")
                            MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                            MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                            
                            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0

                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                VAR_LANZANDO = JsonLanguage.item("PROYECTILES").item("TEXTO")
                                MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                                MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                                
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0

                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    VAR_LANZANDO = JsonLanguage.item("HECHIZOS").item("TEXTO")
                                    MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                                    MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                                    
                                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        Else

                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0

                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    VAR_LANZANDO = JsonLanguage.item("HECHIZOS").item("TEXTO")
                                    MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                                    MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                                    
                                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(TX, TY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                'Call WriteRightClick(tx, tY) 'Proximamnete lo implementaremos..
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then

            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, TX, TY)
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_DblClick()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteDoubleClick(TX, TY)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewPic.Left
    MouseY = Y - MainViewPic.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewPic.Width Then
        MouseX = MainViewPic.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewPic.Height Then
        MouseY = MainViewPic.Height
    End If
    
    LastButtonPressed.ToggleToNormal
    
    ' Disable links checking (not over consola)
    StopCheckingLinks
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub lblDropGold_Click()

    Inventario.SelectGold

    If UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If
    
End Sub

Private Sub cmdInventario_Click()
    Call Audio.PlayWave(SND_CLICK)

    'InvEqu.Picture = LoadPicture(Game.path(Skins) & SkinSeleccionado & "\Centroinventario.jpg")

    ' Activo controles de inventario
    picInv.Visible = True

    ' Desactivo controles de hechizo
    'hlst.Visible = False
    'cmdINFO.Visible = False
    'CmdLanzar.Visible = False
    
    'cmdMoverHechi(0).Visible = False
    'cmdMoverHechi(1).Visible = False
    
    DoEvents
    Call Inventario.DrawInventory
    
End Sub

Private Sub CmdHechizos_Click()
    
    Call Audio.PlayWave(SND_CLICK)

    'InvEqu.Picture = LoadPicture(Game.path(Skins) & SkinSeleccionado & "\Centrohechizos.jpg")
    
    ' Activo controles de hechizos
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    ' Desactivo controles de inventario
    'PicInv.Visible = False

End Sub

Private Sub picInv_DblClick()

    'Esta validacion es para que el juego no rompa si hacemos doble click
    'En un slot vacio (Recox)
    If Inventario.SelectedItem = 0 Then Exit Sub
    If MirandoCarpinteria Or MirandoHerreria Then Exit Sub
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If trainingMacro.Enabled Then Call DesactivarMacroHechizos
    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
    
    Select Case Inventario.OBJType(Inventario.SelectedItem)
        
        Case eObjType.otcasco
            Call EquiparItem
    
        Case eObjType.otArmadura
            Call EquiparItem

        Case eObjType.otescudo
            Call EquiparItem
        
        Case eObjType.otWeapon
            Call EquiparItem
        
        Case eObjType.otAnillo
            Call EquiparItem
        
        Case Else
            Call UsarItem
            
    End Select
    
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()

    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar

    If Not Application.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    
    ElseIf (Not Comerciando) And _
           (Not MirandoAsignarSkills) And _
           (Not frmMSG.Visible) And _
           (Not MirandoForo) And _
           (Not frmEstadisticas.Visible) And _
           (Not frmCantidad.Visible) And _
           (Not MirandoParty) Then

        If PicInv.Visible Then
            PicInv.SetFocus
			
        ElseIf hlst.Visible Then
            hlst.SetFocus

        End If

    End If

End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If

End Sub

Private Sub SendTxt_Change()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 3/06/2006
    '3/06/2006: Maraxus - impedi se inserten caracteres no imprimibles
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = JsonLanguage.item("MENSAJE_SOY_CHEATER").item("TEXTO")
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i         As Long
        Dim tempstr   As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 04/01/2020
'12/28/2007: Recox - Arregle el chat de clanes, ahora funciona correctamente y se puede mandar el mensaje con la misma tecla que se abre la consola.
'**************************************************************
 
    'Send text
    If KeyCode = vbKeyReturn Or KeyCode = CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild) Then

        'Say
        If LenB(stxtbuffercmsg) <> 0 Then
            Call WriteGuildMessage(stxtbuffercmsg)
        End If

        stxtbuffercmsg = vbNullString
        SendCMSTXT.Text = vbNullString
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
        
        If picInv.Visible Then
            picInv.SetFocus
        Else
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
End Sub

Private Sub SendCMSTXT_Change()

    If Len(SendCMSTXT.Text) > 160 Then
        'stxtbuffercmsg = JsonLanguage.item("MENSAJE_SOY_CHEATER").item("TEXTO")
        stxtbuffercmsg = vbNullString ' GSZAO
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i         As Long
        Dim tempstr   As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendCMSTXT.Text)
            CharAscii = Asc(mid$(SendCMSTXT.Text, i, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendCMSTXT.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendCMSTXT.Text = tempstr
        End If
        
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub

Private Sub AbrirMenuViewPort()
    #If (ConMenuseConextuales = 1) Then

        If TX >= MinXBorder And TY >= MinYBorder And TY <= MaxYBorder And TX <= MaxXBorder Then

            If MapData(TX, TY).CharIndex > 0 Then
                If charlist(MapData(TX, TY).CharIndex).invisible = False Then
        
                    Dim m As frmMenuseFashion
                    Set m = New frmMenuseFashion
            
                    Load m
                    m.SetCallback Me
                    m.SetMenuId 1
                    m.ListaInit 2, False
            
                    If LenB(charlist(MapData(TX, TY).CharIndex).Nombre) <> 0 Then
                        m.ListaSetItem 0, charlist(MapData(TX, TY).CharIndex).Nombre, True
                    Else
                        m.ListaSetItem 0, "<NPC>", True
                    End If
                    m.ListaSetItem 1, JsonLanguage.item("COMERCIAR").item("TEXTO")
            
                    m.ListaFin
                    m.Show , Me

                End If
            End If
        End If

    #End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)

    Select Case MenuId

        Case 0 'Inventario

            Select Case Sel

                Case 0

                Case 1

                Case 2 'Tirar
                    Call TirarItem

                Case 3 'Usar

                    If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
                        Call UsarItem
                    End If

                Case 3 'equipar
                    Call EquiparItem
            End Select
    
        Case 1 'Menu del ViewPort del engine

            Select Case Sel

                Case 0 'Nombre
                    Call WriteLeftClick(TX, TY)
        
                Case 1 'Comerciar
                    Call WriteLeftClick(TX, TY)
                    Call WriteCommerceStart
            End Select
    End Select
End Sub

Private Sub SonidosMapas_Timer()
    Sonidos.ReproducirSonidosDeMapas
End Sub
 
''''''''''''''''''''''''''''''''''''''
'     WINDOWS API                            '
''''''''''''''''''''''''''''''''''''''
Private Sub Client_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
    Second.Enabled = True
    
    Select Case EstadoLogin

        Case E_MODO.CrearNuevoPj, E_MODO.Normal
            Call Login
        
        Case E_MODO.CrearCuenta
            Call Audio.PlayBackgroundMusic("7", MusicTypes.Mp3)
            frmCrearCuenta.Show

        Case E_MODO.Dados
            Call Audio.PlayBackgroundMusic("7", MusicTypes.Mp3)
            frmCrearPersonaje.Show
            
        Case E_MODO.CambiarContrasena
            Call Audio.PlayBackgroundMusic("7", MusicTypes.Mp3)
            frmRecuperarCuenta.Show

        Case E_MODO.ObtenerDatosServer
            Call WriteObtenerDatosServer
        
    End Select
 
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
    Dim RD     As String
    Dim data() As Byte
    
    Client.GetData RD, vbByte, bytesTotal
    data = StrConv(RD, vbFromUnicode)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
    
End Sub

Private Sub Client_CloseSck()
    
    Debug.Print "Cerrando la conexion via API de Windows..."

    Call ResetAllInfo
End Sub

Private Sub Client_Error(ByVal number As Integer, _
                         Description As String, _
                         ByVal sCode As Long, _
                         ByVal source As String, _
                         ByVal HelpFile As String, _
                         ByVal HelpContext As Long, _
                         CancelDisplay As Boolean)
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    
    frmConnect.MousePointer = 1
    
    Second.Enabled = False
 
    If Client.State <> sckClosed Then Client.CloseSck

    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
 
End Sub

Private Function InGameArea() As Boolean
'********************************************************************
'Author: NicoNZ
'Last Modification: 29/09/2019
'Checks if last click was performed within or outside the game area.
'********************************************************************
    If clicX < 0 Or clicX > frmMain.MainViewPic.Width Then Exit Function
    If clicY < 0 Or clicY > frmMain.MainViewPic.Height Then Exit Function
    
    InGameArea = True
End Function

Private Sub hlst_Click()
    
    With hlst
    
        If ChangeHechi Then
    
            Dim NewLugar As Integer: NewLugar = .ListIndex
            Dim AntLugar As String: AntLugar = .List(NewLugar)
            
            Call WriteDragAndDropHechizos(ChangeHechiNum + 1, NewLugar + 1)
        
            .BackColor = vbBlack
            .List(NewLugar) = .List(ChangeHechiNum)
            .List(ChangeHechiNum) = AntLugar
        
            ChangeHechi = False
            ChangeHechiNum = 0

        End If

        .BackColor = vbBlack

    End With

End Sub

Private Sub hlst_DblClick()
    ChangeHechi = True
    ChangeHechiNum = hlst.ListIndex
    hlst.BackColor = vbRed

End Sub

'Incorporado por ReyarB
Private Sub Minimapa_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    If Button = vbRightButton Then
        Call WriteWarpChar("YO", UserMap, CByte(X), CByte(Y))
        Call ActualizarMiniMapa
    End If
End Sub
'fin Incorporado ReyarB

Public Sub ActualizarMiniMapa()
    '***************************************************
    'Author: Martin Gomez (Samke)
    'Last Modify Date: 05/01/2020
    'Integrado por Reyarb
    'Se agrego campo de vision del render (Recox)
    'Ajustadas las coordenadas para centrarlo (WyroX)
    '***************************************************
    Me.UserM.Left = UserPos.X - 2
    Me.UserM.Top = UserPos.Y - 2
    Me.UserAreaMinimap.Left = UserPos.X - 10
    Me.UserAreaMinimap.Top = UserPos.Y - 8
    Me.MiniMapa.Refresh
End Sub

Public Sub ActivarMacroHechizos()

    If Not hlst.Visible Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, True)
        Exit Sub
    End If
    
    trainingMacro.Interval = INT_MACRO_HECHIS
    trainingMacro.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mSpells, True)
End Sub

Public Sub DesactivarMacroHechizos()
    trainingMacro.Enabled = False
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True)
    Call ControlSM(eSMType.mSpells, False)
End Sub

Private Sub timerTiempoRestanteInvisibleMensaje_Timer()
    If UserInvisible Then
        UserInvisibleSegundosRestantes = UserInvisibleSegundosRestantes - 1
    End If
End Sub

Private Sub timerTiempoRestanteParalisisMensaje_Timer()
    If UserParalizado Then
        UserParalizadoSegundosRestantes = UserParalizadoSegundosRestantes - 1
    End If
End Sub

Private Sub trainingMacro_Timer()

    If Not hlst.Visible Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    'Macros are disabled if focus is not on Argentum!
    If Not Application.IsAppActive() Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    If Comerciando Then Exit Sub
    
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.CastSpell, False) Then
        Call WriteCastSpell(hlst.ListIndex + 1)
        Call WriteWork(eSkill.Magia)
    End If
    
    Call ConvertCPtoTP(MouseX, MouseY, TX, TY)
    
    If UsingSkill = Magia And Not MainTimer.Check(TimersIndex.CastSpell) Then Exit Sub
    
    If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub
    
    Call WriteWorkLeftClick(TX, TY, UsingSkill)
    UsingSkill = 0
End Sub

Public Sub UpdateProgressExperienceLevelBar(ByVal UserExp As Long)
    If UserLvl = STAT_MAXELV Then
        frmMain.lblPorcLvl.Caption = "[N/A]"

        'Si no tiene mas niveles que subir ponemos la barra al maximo.
        frmMain.uAOProgressExperienceLevel.Max = 100
        frmMain.uAOProgressExperienceLevel.Value = 100
    Else
        frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
        frmMain.uAOProgressExperienceLevel.Max = UserPasarNivel
        frmMain.uAOProgressExperienceLevel.Value = UserExp
    End If
End Sub

Public Sub SetGoldColor()

    If UserGLD >= CLng(UserLvl) * 10000 And UserLvl > 12 Then 'Si el nivel es mayor de 12, es decir, no es newbie.
        'Changes color
        frmMain.GldLbl.ForeColor = USER_GOLD_COLOR
    Else
        'Changes color
        frmMain.GldLbl.ForeColor = NEWBIE_USER_GOLD_COLOR
    End If

End Sub
