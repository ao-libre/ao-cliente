VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   11505
   ClientLeft      =   360
   ClientTop       =   -3300
   ClientWidth     =   15330
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
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":7F6A
   ScaleHeight     =   767
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1022
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox pHechizos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3180
      Left            =   12000
      Picture         =   "frmMain.frx":3F80C
      ScaleHeight     =   212
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   49
      Top             =   3000
      Visible         =   0   'False
      Width           =   2700
      Begin VB.Image BarraHechizosCentro 
         Height          =   2640
         Left            =   2430
         Top             =   270
         Width           =   270
      End
      Begin VB.Image BarritaHechizos 
         Height          =   120
         Left            =   2430
         Picture         =   "frmMain.frx":41D1C
         Top             =   285
         Width           =   270
      End
      Begin VB.Image BarraHechizosDown 
         Height          =   270
         Left            =   2430
         Top             =   2910
         Width           =   270
      End
      Begin VB.Image BarraHechizosUp 
         Height          =   270
         Left            =   2430
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.Timer timerPasarSegundo 
      Interval        =   1000
      Left            =   4560
      Top             =   2640
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   13920
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   18
      Top             =   10920
      Width           =   420
   End
   Begin VB.PictureBox MiniMapa 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   9540
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   600
      Width           =   1500
      Begin VB.Shape UserAreaMinimap 
         BackColor       =   &H80000004&
         BorderColor     =   &H80000002&
         FillColor       =   &H000080FF&
         Height          =   315
         Left            =   555
         Top             =   585
         Width           =   375
      End
      Begin VB.Shape UserM 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         Height          =   45
         Left            =   720
         Top             =   720
         Width           =   45
      End
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   13440
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   26
      Top             =   10920
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   12840
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   21
      Top             =   10920
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   12360
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   20
      Top             =   10920
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3810
      Left            =   11775
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   15
      Top             =   3120
      Width           =   3150
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
      Height          =   360
      Left            =   600
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "frmMain.frx":42083
      ToolTipText     =   "Chat"
      Top             =   1830
      Visible         =   0   'False
      Width           =   8535
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
      Height          =   360
      Left            =   360
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "frmMain.frx":420B3
      ToolTipText     =   "Chat"
      Top             =   10800
      Visible         =   0   'False
      Width           =   10695
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
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9120
      Left            =   120
      MousePointer    =   99  'Custom
      ScaleHeight     =   608
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   736
      TabIndex        =   27
      Top             =   2280
      Width           =   11040
      Begin InetCtlsObjects.Inet InetDownloadFfmpeg 
         Left            =   120
         Top             =   1560
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Timer tmrCounters 
         Left            =   5040
         Top             =   360
      End
      Begin VB.Timer trainingMacro 
         Enabled         =   0   'False
         Interval        =   3200
         Left            =   10200
         Top             =   600
      End
   End
   Begin AOLibre.uAOButton btnMapa 
      Height          =   255
      Left            =   11160
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Mapa"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":420E9
      PICF            =   "frmMain.frx":42B13
      PICH            =   "frmMain.frx":437D5
      PICV            =   "frmMain.frx":44767
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
   Begin AOLibre.uAOButton btnGrupo 
      Height          =   255
      Left            =   12600
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Grupo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":45669
      PICF            =   "frmMain.frx":46093
      PICH            =   "frmMain.frx":46D55
      PICV            =   "frmMain.frx":47CE7
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
   Begin AOLibre.uAOButton btnOpciones 
      Height          =   255
      Left            =   12360
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Opciones"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":48BE9
      PICF            =   "frmMain.frx":49613
      PICH            =   "frmMain.frx":4A2D5
      PICV            =   "frmMain.frx":4B267
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
   Begin AOLibre.uAOButton btnEstadisticas 
      Height          =   255
      Left            =   12600
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Estadisticas"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":4C169
      PICF            =   "frmMain.frx":4CB93
      PICH            =   "frmMain.frx":4D855
      PICV            =   "frmMain.frx":4E7E7
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
   Begin AOLibre.uAOButton btnClanes 
      Height          =   255
      Left            =   14040
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Clanes"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":4F6E9
      PICF            =   "frmMain.frx":50113
      PICH            =   "frmMain.frx":50DD5
      PICV            =   "frmMain.frx":51D67
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
   Begin AOLibre.uAOButton btnInventario 
      Height          =   375
      Left            =   11520
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      TX              =   "Inventario"
      ENAB            =   -1  'True
      FCOL            =   8421504
      OCOL            =   16777215
      PICE            =   "frmMain.frx":52C69
      PICF            =   "frmMain.frx":53DED
      PICH            =   "frmMain.frx":55791
      PICV            =   "frmMain.frx":57825
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
   Begin AOLibre.uAOButton btnHechizos 
      Height          =   375
      Left            =   13440
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      TX              =   "Hechizos"
      ENAB            =   -1  'True
      FCOL            =   8421504
      OCOL            =   16777215
      PICE            =   "frmMain.frx":59889
      PICF            =   "frmMain.frx":5AA0D
      PICH            =   "frmMain.frx":5C3B1
      PICV            =   "frmMain.frx":5E445
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
   Begin AOLibre.uAOButton btnLanzar 
      Height          =   615
      Left            =   11760
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      TX              =   "Lanzar"
      ENAB            =   -1  'True
      FCOL            =   8421504
      OCOL            =   16777215
      PICE            =   "frmMain.frx":604A9
      PICF            =   "frmMain.frx":6162D
      PICH            =   "frmMain.frx":62FD1
      PICV            =   "frmMain.frx":65065
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
   Begin AOLibre.uAOButton btnInfo 
      Height          =   615
      Left            =   13800
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      TX              =   "Info"
      ENAB            =   -1  'True
      FCOL            =   8421504
      OCOL            =   16777215
      PICE            =   "frmMain.frx":670C9
      PICF            =   "frmMain.frx":6824D
      PICH            =   "frmMain.frx":69BF1
      PICV            =   "frmMain.frx":6BC85
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
   Begin AOLibre.uAOButton btnRetos 
      Height          =   255
      Left            =   12720
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   10560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Retos"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":6DCE9
      PICF            =   "frmMain.frx":6E713
      PICH            =   "frmMain.frx":6F3D5
      PICV            =   "frmMain.frx":70367
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
   Begin AOLibre.uAOButton btnAmigos 
      Height          =   255
      Left            =   11280
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   10440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Amigos"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":71269
      PICF            =   "frmMain.frx":71C93
      PICH            =   "frmMain.frx":72955
      PICV            =   "frmMain.frx":738E7
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
   Begin AOLibre.uAOButton btnQuests 
      Height          =   255
      Left            =   14040
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   10560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Quests"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":747E9
      PICF            =   "frmMain.frx":75213
      PICH            =   "frmMain.frx":75ED5
      PICV            =   "frmMain.frx":76E67
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
   Begin AOLibre.uAOButton btnReportarBug 
      Height          =   495
      Left            =   11400
      TabIndex        =   43
      Top             =   10920
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   873
      TX              =   "Reportar Bug"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":77D69
      PICF            =   "frmMain.frx":77D85
      PICH            =   "frmMain.frx":77DA1
      PICV            =   "frmMain.frx":77DBD
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
   Begin AOLibre.uAOButton btnGrabarVideo 
      Height          =   495
      Left            =   14400
      TabIndex        =   44
      Top             =   10920
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   873
      TX              =   "Grabar Video"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":77DD9
      PICF            =   "frmMain.frx":77DF5
      PICH            =   "frmMain.frx":77E11
      PICV            =   "frmMain.frx":77E2D
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
   Begin VB.PictureBox pConsola 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   135
      Picture         =   "frmMain.frx":77E49
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   620
      TabIndex        =   48
      Top             =   480
      Width           =   9300
      Begin VB.Image BarraConsolaCentro 
         Height          =   765
         Left            =   9030
         Top             =   270
         Width           =   270
      End
      Begin VB.Image BarritaConsola 
         Height          =   120
         Left            =   9030
         Picture         =   "frmMain.frx":79E9E
         Top             =   915
         Width           =   270
      End
      Begin VB.Image BarraConsolaUp 
         Height          =   270
         Left            =   9030
         Top             =   0
         Width           =   270
      End
      Begin VB.Image BarraConsolaDown 
         Height          =   270
         Left            =   9030
         Top             =   1035
         Width           =   270
      End
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   0
      Left            =   14790
      MouseIcon       =   "frmMain.frx":7A205
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":7A357
      Top             =   3960
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   1
      Left            =   14790
      MouseIcon       =   "frmMain.frx":7A69B
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":7A7ED
      Top             =   3705
      Visible         =   0   'False
      Width           =   225
   End
   Begin AOLibre.uAOProgress uAOProgressDownloadFfmpeg 
      Height          =   255
      Left            =   2160
      TabIndex        =   45
      ToolTipText     =   "Descarga ffmpeg"
      Top             =   2040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      Min             =   1
      Value           =   1
      UseBackground   =   0   'False
      BackgroundColor =   65280
      ForeColor       =   12632319
      BackColor       =   16512
      BorderColor     =   0
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOProgress uAOProgressExperienceLevel 
      Height          =   330
      Left            =   11490
      TabIndex        =   37
      ToolTipText     =   "Experiencia necesaria para pasar de nivel"
      Top             =   1500
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      Max             =   999999
      Value           =   0
      UseBackground   =   0   'False
      ForeColor       =   -2147483624
      BackColor       =   49152
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
   Begin VB.Image imgInvLock 
      Height          =   210
      Index           =   2
      Left            =   11520
      Top             =   6570
      Width           =   210
   End
   Begin VB.Image imgInvLock 
      Height          =   210
      Index           =   1
      Left            =   11520
      Top             =   6045
      Width           =   210
   End
   Begin VB.Image imgInvLock 
      Height          =   210
      Index           =   0
      Left            =   11520
      Top             =   5535
      Width           =   210
   End
   Begin VB.Image ImgQuests 
      Height          =   495
      Left            =   14640
      Top             =   9960
      Width           =   495
   End
   Begin VB.Image ImgRetos 
      Height          =   495
      Left            =   14040
      Top             =   9960
      Width           =   495
   End
   Begin VB.Image ImgAmigos 
      Height          =   495
      Left            =   13440
      Top             =   9960
      Width           =   495
   End
   Begin VB.Image ImgClan 
      Height          =   495
      Left            =   12840
      Top             =   9960
      Width           =   495
   End
   Begin VB.Image ImgGrupo 
      Height          =   615
      Left            =   12120
      Top             =   9960
      Width           =   615
   End
   Begin VB.Image ImgMapa 
      Height          =   615
      Left            =   11520
      Top             =   9960
      Width           =   615
   End
   Begin VB.Label lblEstadisticas 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   14760
      TabIndex        =   47
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblMapName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del mapa"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   6600
      TabIndex        =   36
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label lblOpciones 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   11400
      TabIndex        =   46
      Top             =   0
      Width           =   375
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
      Height          =   345
      Left            =   14040
      TabIndex        =   39
      Top             =   1560
      Width           =   555
   End
   Begin VB.Image imgAsignarSkill 
      Height          =   420
      Left            =   14760
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":7AB31
      Top             =   960
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13800
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   14640
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   15000
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblFPS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "101"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   8640
      TabIndex        =   17
      Top             =   120
      Width           =   795
   End
   Begin VB.Image xz 
      Height          =   255
      Index           =   0
      Left            =   14280
      Top             =   120
      Width           =   255
   End
   Begin VB.Image xzz 
      Height          =   195
      Index           =   1
      Left            =   13920
      Top             =   120
      Width           =   225
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del pj"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   11400
      TabIndex        =   38
      Top             =   600
      Width           =   3825
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   11520
      TabIndex        =   16
      Top             =   1080
      Width           =   3525
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0080FFFF&
      Height          =   210
      Left            =   11880
      TabIndex        =   14
      Top             =   7350
      Width           =   90
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   15000
      TabIndex        =   8
      Top             =   7350
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   14400
      TabIndex        =   7
      Top             =   7350
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
      Left            =   9675
      TabIndex        =   6
      Top             =   180
      Width           =   1455
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Left            =   11550
      TabIndex        =   5
      Top             =   9030
      Width           =   975
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Left            =   12600
      TabIndex        =   4
      Top             =   9030
      Width           =   975
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Left            =   13485
      TabIndex        =   3
      Top             =   9030
      Width           =   975
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Left            =   14475
      TabIndex        =   2
      Top             =   9030
      Width           =   975
   End
   Begin VB.Image imgScroll 
      Height          =   240
      Index           =   1000
      Left            =   14760
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   225
   End
   Begin VB.Image InvEqu 
      Height          =   4455
      Left            =   11880
      Top             =   2880
      Width           =   2970
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   11955
      TabIndex        =   10
      Top             =   8340
      Width           =   1215
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   11955
      TabIndex        =   9
      Top             =   8625
      Width           =   1215
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   11955
      TabIndex        =   11
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   13920
      TabIndex        =   12
      Top             =   8340
      Width           =   1095
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   13920
      TabIndex        =   13
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Shape shpEnergia 
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   11910
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Shape shpMana 
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   11910
      Top             =   8355
      Width           =   1335
   End
   Begin VB.Shape shpVida 
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   11910
      Top             =   8070
      Width           =   1335
   End
   Begin VB.Shape shpHambre 
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   13785
      Top             =   8355
      Width           =   1335
   End
   Begin VB.Shape shpSed 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   13785
      Top             =   8055
      Width           =   1335
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
   Begin VB.Menu mnu_ShowConsole 
      Caption         =   "Consola"
      Visible         =   0   'False
      Begin VB.Menu mnu_SetConsolaGeneral 
         Caption         =   "General"
      End
      Begin VB.Menu mnu_SetConsolaCombate 
         Caption         =   "Combate"
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

Dim BoldX As Long, BoldY As Long, BisMoving As Boolean

Public SendTxtHasFocus As Boolean
Public SendCMSTXTHasFocus As Boolean

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

Private bIsRecordingVideo  As Boolean

Private sFfmpegTaskId As String

'Peso del archivo ffmpeg
Dim lSizeInBytes As Long

'Para la descarga de ffmpeg
Dim Directory As String, bDone As Boolean, dError As Boolean

'Usado para controlar que no se dispare el binding de la tecla CTRL cuando se usa CTRL+Tecla.
Dim CtrlMaskOn             As Boolean
Dim SkinSeleccionado       As String

'Para cuando se necesite enviar un mensaje a la consola dentro de un bucle al que no debas parar
Public MsgTimeadoOn        As Boolean
Public MsgTimeado          As String

Private Const NEWBIE_USER_GOLD_COLOR As Long = vbCyan
Private Const USER_GOLD_COLOR As Long = vbYellow

Private Type tConsola
    Texto As String
    Color As Long
    bold As Byte
    italic As Byte
End Type
Private OffSetConsola As Integer
Private LineasConsola As Integer

Private Const CONSOLE_LINE_HEIGHT As Integer = 14
Private Const MAX_CONSOLE_LINES As Integer = 600
Private Const CONSOLE_REMOVED_LINES As Integer = 100
Private Const CONSOLE_ARROWS_DISPLACEMENT As Integer = 1
Private Const CONSOLE_PADDING As Integer = 4
Private CONSOLE_VISIBLE_LINES As Integer

Private Consola(MAX_CONSOLE_LINES) As tConsola

Public hlst As clsGraphicalList
Private Const SPELLS_ARROWS_DISPLACEMENT As Integer = 1
Private Const SPELLS_PADDING As Integer = 4
Private SPELLS_VISIBLE_LINES As Integer

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call Protocol.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
End Sub

Private Sub DownloadFfmpeg()
    If MsgBox(JsonLanguage.item("BTN_RECORD_VIDEO_DESCARGAR_APLICACION").item("TEXTO"), vbYesNo) = vbYes Then
        Dim sFfmpegExeFilePath As String
        sFfmpegExeFilePath = App.path & "\ffmpeg.exe"

        btnGrabarVideo.Enabled = False
        btnGrabarVideo.Visible = False
        uAOProgressDownloadFfmpeg.Visible = True

        lSizeInBytes = 53521905
        uAOProgressDownloadFfmpeg.max = lSizeInBytes

        InetDownloadFfmpeg.AccessType = icUseDefault
        InetDownloadFfmpeg.URL = "https://github.com/ao-libre/ao-website/releases/download/v1.0/ffmpeg.exe"
        Directory = sFfmpegExeFilePath
        bDone = False
        dError = False
            
        InetDownloadFfmpeg.Execute , "GET"
        
        Do While bDone = False
            DoEvents
        Loop
        
        uAOProgressDownloadFfmpeg.Visible = False
        btnGrabarVideo.Visible = True
        btnGrabarVideo.Enabled = True

        If dError Then
            Call MsgBox(JsonLanguage.item("FFMPEG_ERROR_DESCARGA_INSTRUCCIONES").item("TEXTO"))
            Exit Sub
        End If

        Exit Sub
    End If
End Sub

Private Sub BarraConsolaCentro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewOffset As Integer
    
    If LineasConsola <= CONSOLE_VISIBLE_LINES Then
        NewOffset = 0
    Else
        ' El 15 es porque convierto de twip a pixel
        NewOffset = Round((Y \ 15) * (LineasConsola - CONSOLE_VISIBLE_LINES) / BarraConsolaCentro.Height)
    End If
    
    If NewOffset <> OffSetConsola Then
        OffSetConsola = NewOffset
        ReDrawConsola
    End If
End Sub

Private Sub BarraConsolaCentro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        If Y < 0 Then Y = 0
        If (Y \ 15) > BarraConsolaCentro.Height Then Y = BarraConsolaCentro.Height * 15
        Call BarraConsolaCentro_MouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub BarraConsolaUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewOffset As Integer
    
    If OffSetConsola > 0 Then
        NewOffset = IIf(OffSetConsola > CONSOLE_ARROWS_DISPLACEMENT, OffSetConsola - CONSOLE_ARROWS_DISPLACEMENT, 0)
    
        If NewOffset <> OffSetConsola Then
            OffSetConsola = NewOffset
            ReDrawConsola
        End If
    End If
End Sub

Private Sub BarraConsolaDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewOffset As Integer
    
    If OffSetConsola < LineasConsola - CONSOLE_VISIBLE_LINES Then
        NewOffset = IIf(LineasConsola - CONSOLE_VISIBLE_LINES - OffSetConsola > CONSOLE_ARROWS_DISPLACEMENT, OffSetConsola + CONSOLE_ARROWS_DISPLACEMENT, LineasConsola - CONSOLE_VISIBLE_LINES)
    
        If NewOffset <> OffSetConsola Then
            OffSetConsola = NewOffset
            ReDrawConsola
        End If
    End If
End Sub


Private Sub BarraHechizosCentro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If hlst.ListCount <= SPELLS_VISIBLE_LINES Then
        hlst.Scroll = 0
    Else
        ' El 15 es porque convierto de twip a pixel
        hlst.Scroll = Round((Y \ 15) * (hlst.ListCount - SPELLS_VISIBLE_LINES) / BarraHechizosCentro.Height)
    End If
    
    If hlst.ListCount <= SPELLS_VISIBLE_LINES Then
        BarritaHechizos.Top = BarraHechizosCentro.Top
    Else
        BarritaHechizos.Top = BarraHechizosCentro.Top + hlst.Scroll * (BarraHechizosCentro.Height - BarritaHechizos.Height) \ (hlst.ListCount - SPELLS_VISIBLE_LINES)
    End If
End Sub

Private Sub BarraHechizosCentro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        If Y < 0 Then Y = 0
        If (Y \ 15) > BarraHechizosCentro.Height Then Y = BarraHechizosCentro.Height * 15
        Call BarraHechizosCentro_MouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub BarraHechizosUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If hlst.Scroll > 0 Then
        hlst.Scroll = IIf(hlst.Scroll > SPELLS_ARROWS_DISPLACEMENT, hlst.Scroll - SPELLS_ARROWS_DISPLACEMENT, 0)
        
        If hlst.ListCount <= SPELLS_VISIBLE_LINES Then
            BarritaHechizos.Top = BarraHechizosCentro.Top
        Else
            BarritaHechizos.Top = BarraHechizosCentro.Top + hlst.Scroll * (BarraHechizosCentro.Height - BarritaHechizos.Height) \ (hlst.ListCount - SPELLS_VISIBLE_LINES)
        End If
    End If
End Sub

Private Sub BarraHechizosDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If hlst.Scroll < hlst.ListCount - SPELLS_VISIBLE_LINES Then
        hlst.Scroll = IIf(hlst.ListCount - SPELLS_VISIBLE_LINES - hlst.Scroll > SPELLS_ARROWS_DISPLACEMENT, hlst.Scroll + SPELLS_ARROWS_DISPLACEMENT, hlst.ListCount - SPELLS_VISIBLE_LINES)
        
        If hlst.ListCount <= SPELLS_VISIBLE_LINES Then
            BarritaHechizos.Top = BarraHechizosCentro.Top
        Else
            BarritaHechizos.Top = BarraHechizosCentro.Top + hlst.Scroll * (BarraHechizosCentro.Height - BarritaHechizos.Height) \ (hlst.ListCount - SPELLS_VISIBLE_LINES)
        End If
    End If
End Sub

Private Sub ImgAmigos_Click()
    Call frmAmigos.Show(vbModeless, frmMain)
End Sub

Private Sub ImgClan_Click()
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub ImgGrupo_Click()
    Call WriteRequestPartyForm
End Sub

Private Sub imgMapa_Click()
    Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub ImgQuests_Click()
    Call WriteQuestListRequest
End Sub

Private Sub ImgRetos_Click()
    Call FrmRetos.Show(vbModeless, frmMain)
End Sub

Private Sub InetDownloadFfmpeg_StateChanged(ByVal State As Integer)
    Dim Percentage As Long
    Select Case State
        Case icError
            Call MsgBox(JsonLanguage.item("FFMPEG_ERROR_DESCARGA_INSTRUCCIONES").item("TEXTO"))
            bDone = True
            dError = True
            uAOProgressDownloadFfmpeg.Visible = False
            btnGrabarVideo.Visible = True
            btnGrabarVideo.Enabled = True
        Case icResponseCompleted
            Dim vtData As Variant
            Dim tempArray() As Byte
            
            Dim G_Num As Integer
            G_Num = FreeFile
            Open Directory For Binary Access Write As #G_Num
                vtData = InetDownloadFfmpeg.GetChunk(1024, icByteArray)
                DoEvents
                
                Do While Not Len(vtData) = 0
                    tempArray = vtData
                    Put #G_Num, , tempArray
                    
                    vtData = InetDownloadFfmpeg.GetChunk(1024, icByteArray)

                    uAOProgressDownloadFfmpeg.min = uAOProgressDownloadFfmpeg.min + Len(vtData) * 2
                    'Percentage = (uAOProgressDownloadFfmpeg.Value / uAOProgressDownloadFfmpeg.max) * 100
                    'uAOProgressDownloadFfmpeg.Text = "[" & Percentage & "% de " & lSizeInBytes & " MBs.]"
                    
                    DoEvents
                Loop
            Close #G_Num
            
            Call MsgBox(JsonLanguage.item("FFMPEG_DESCARGA_FINALIZADA").item("TEXTO"))

            bDone = True
    End Select
End Sub

Private Sub btnGrabarVideo_Click()
    Dim sFfmpegExeFilePath As String
    sFfmpegExeFilePath = App.path & "\ffmpeg.exe"

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject

    'Comprobamos si existe ffmpeg, sino existe lo bajamos
    If Not fso.FileExists(sFfmpegExeFilePath) Then
        Call DownloadFfmpeg
        Exit Sub
    End If

    If Not bIsRecordingVideo Then
        bIsRecordingVideo = True
        btnGrabarVideo.Caption = JsonLanguage.item("BTN_RECORD_VIDEO_FINALIZAR").item("TEXTO")


        Dim FileName As String
        FileName = Format$(Now, "DD-MM-YYYY-hh.mm.ss") & "_ao-libre.mkv"

        Call MsgBox(JsonLanguage.item("BTN_RECORD_VIDEO_MESSAGE").item("TEXTO"))

        Dim sFfmpegCommand As String
        sFfmpegCommand = sFfmpegExeFilePath & " -f gdigrab -framerate 30 -i title=""Argentum Online Libre"" " & Game.path(Videos) & FileName

        sFfmpegTaskId = Shell(sFfmpegCommand)
    Else
        'Matamos ffmpeg por lo cual se guarda el video :)
        Shell ("taskkill /PID " & sFfmpegTaskId)
        bIsRecordingVideo = False
        btnGrabarVideo.Caption = JsonLanguage.item("BTN_RECORD_VIDEO").item("TEXTO")
        Call MsgBox(JsonLanguage.item("BTN_RECORD_VIDEO_MESSAGE_FINISH").item("TEXTO"))
        Shell ("explorer " & Game.path(Videos))
    End If

End Sub

Private Sub btnReportarBug_Click()
    Call MsgBox(JsonLanguage.item("BTN_REPORTAR_BUG_MESSAGE").item("TEXTO"))
    Call ShellExecute(0, "Open", "https://github.com/ao-libre/ao-cliente/issues", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub btnAmigos_Click()
    Call frmAmigos.Show(vbModeless, frmMain)
End Sub

Private Sub btnQuests_Click()
    Call WriteQuestListRequest
End Sub

Private Sub Form_Activate()
    Call Inventario.DrawInventory
End Sub

Private Sub BarritaMover_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not ResolucionCambiada Then
        BoldX = X
        BoldY = Y
        BisMoving = True
    End If
End Sub
Private Sub BarritaMover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If BisMoving Then
        Me.Top = Me.Top - (BoldY - Y)
        Me.Left = Me.Left - (BoldX - X)
    End If
End Sub
Private Sub BarritaMover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BisMoving = False
End Sub

Private Sub Form_Load()
    SkinSeleccionado = GetVar(Game.path(INIT) & "Config.ini", "Parameters", "SkinSelected")
    
    Me.Picture = LoadPicture(Game.path(Skins) & SkinSeleccionado & "\VentanaPrincipal.jpg")

    If Not ResolucionCambiada Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        Call clsFormulario.Initialize(Me, 120)
    End If
        
    Call LoadButtons

    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)
        
    ' Detect links in console
    Call EnableURLDetect(pConsola.hwnd, Me.hwnd)
    
    ' Hacer las consolas transparentes
    'Call SetWindowLong(RecTxt.hwnd, -20, &H20&)
    
    ' Seteamos el caption
    Me.Caption = "Argentum Online Libre"
    
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(Me)
    
    ' Reseteamos el tamanio de la ventana para que no queden bordes blancos
    Me.Width = 15360
    Me.Height = 11520
    
    CtrlMaskOn = False
    
    FirstTimeChat = True
    FirstTimeClanChat = True
    bIsRecordingVideo = False
    uAOProgressDownloadFfmpeg.Visible = False
    
    CONSOLE_VISIBLE_LINES = (pConsola.Height + CONSOLE_PADDING * 2) \ CONSOLE_LINE_HEIGHT
    
    Set hlst = New clsGraphicalList
    Call hlst.Initialize(pHechizos, pHechizos.ForeColor, SPELLS_PADDING, BarraHechizosCentro.Width)
    
    SPELLS_VISIBLE_LINES = (pHechizos.Height + SPELLS_PADDING * 2) \ hlst.Pixel_Alto
    
End Sub

Private Sub LoadTextsForm()
    
    btnLanzar.Caption = JsonLanguage.item("LBL_LANZAR").item("TEXTO")
    btnInventario.Caption = JsonLanguage.item("LBL_INVENTARIO").item("TEXTO")
    btnHechizos.Caption = JsonLanguage.item("LBL_HECHIZOS").item("TEXTO")
    btnInfo.Caption = JsonLanguage.item("LBL_INFO").item("TEXTO")
    btnMapa.Caption = JsonLanguage.item("LBL_MAPA").item("TEXTO")
    btnGrupo.Caption = JsonLanguage.item("LBL_GRUPO").item("TEXTO")
    btnOpciones.Caption = JsonLanguage.item("LBL_OPCIONES").item("TEXTO")
    btnEstadisticas.Caption = JsonLanguage.item("LBL_ESTADISTICAS").item("TEXTO")
    btnClanes.Caption = JsonLanguage.item("LBL_CLANES").item("TEXTO")
    btnAmigos.Caption = JsonLanguage.item("LBL_AMIGOS").item("TEXTO")
    btnRetos.Caption = JsonLanguage.item("LBL_RETOS").item("TEXTO")
    btnQuests.Caption = JsonLanguage.item("LBL_QUESTS").item("TEXTO")
    btnReportarBug.Caption = JsonLanguage.item("LBL_REPORTAR_BUG").item("TEXTO")
    btnGrabarVideo.Caption = JsonLanguage.item("BTN_RECORD_VIDEO").item("TEXTO")
    
End Sub

Private Sub LoadButtons()
    Dim i As Integer

    Set cBotonDiamArriba = New clsGraphicalButton
    Set cBotonDiamAbajo = New clsGraphicalButton
    Set cBotonAsignarSkill = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton

    'Set picSkillStar = LoadPicture(Game.path(Skins) & SkinSeleccionado & "\BotonAsignarSkills.bmp")

    'If SkillPoints > 0 Then imgAsignarSkill.Picture = picSkillStar
    
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
        'imgAsignarSkill.Picture = picSkillStar
        imgAsignarSkill.Visible = True
    Else
        'Set imgAsignarSkill.Picture = Nothing
        imgAsignarSkill.Visible = False
    End If
End Sub

Private Sub cmdMoverHechi_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                
                If hlst.ListCount <= SPELLS_VISIBLE_LINES Then
                    BarritaHechizos.Top = BarraHechizosCentro.Top
                Else
                    BarritaHechizos.Top = BarraHechizosCentro.Top + hlst.Scroll * (BarraHechizosCentro.Height - BarritaHechizos.Height) \ (hlst.ListCount - SPELLS_VISIBLE_LINES)
                End If

            Case 0 'bajar
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1
                
                If hlst.ListCount <= SPELLS_VISIBLE_LINES Then
                    BarritaHechizos.Top = BarraHechizosCentro.Top
                Else
                    BarritaHechizos.Top = BarraHechizosCentro.Top + hlst.Scroll * (BarraHechizosCentro.Height - BarritaHechizos.Height) \ (hlst.ListCount - SPELLS_VISIBLE_LINES)
                End If
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
                Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("COLOR").item(3), _
                                     True, False, True)
                                        
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("TEXTO")
            Else
                Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_SEGURO_RESU_OFF").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_OFF").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_OFF").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_RESU_OFF").item("COLOR").item(3), _
                                     True, False, True)
                                        
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_RESU_ON").item("TEXTO")
            End If
            
        Case eSMType.sSafemode
            
            If Mostrar Then
                Call frmMain.AddtoRichPicture(UCase$(JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("TEXTO").item(1)), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(3), _
                                      True, False, True)
                                        
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("TEXTO").item(2)
            Else
                Call frmMain.AddtoRichPicture(UCase$(JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("TEXTO").item(1)), _
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
                
                'Si es CTRL+1..9 .... lo dejo por si queremos poner macros algun dia.
                'ElseIf KeyCode >= vbKey1 And KeyCode <= vbKey9 Then

                                        
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
            If charlist(UserCharIndex).Clan = vbNullString Then Exit Sub
            
            If SendCMSTXT.Visible And Not SendCMSTXTHasFocus Then
                Call SendCMSTXT_SendText
                Exit Sub
            End If
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not MirandoForo) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendCMSTXT.Visible = True
                
                If Not Typing Then
                    Call WriteSetTypingFlagFromUserCharIndex
                    Typing = True
                End If
                
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
            
            If SendTxt.Visible And Not SendTxtHasFocus Then
                Call SendTxt_SendText
                Exit Sub
            End If
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not MirandoForo) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus

                If Not Typing Then
                    Call WriteSetTypingFlagFromUserCharIndex
                    Typing = True
                End If
                
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
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_GOLD_LABEL").item("TEXTO"), _
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

Private Sub btnClanes_Click()
    
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub btnEstadisticas_Click()

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

Private Sub btnGrupo_Click()
    
    Call WriteRequestPartyForm
End Sub

Private Sub btnMapa_Click()
    
    Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub btnOpciones_Click()
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
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_ARMOR_LABEL").item("TEXTO"), _
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
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_DEXT_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_DEXT_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_DEXT_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_DEXT_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblEnergia_Click()
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_ENERGIA_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_ENERGIA_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_ENERGIA_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_ENERGIA_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub



Private Sub lblEstadisticas_Click()
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

Private Sub lblHambre_Click()
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_HAMBRE_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_HAMBRE_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_HAMBRE_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_HAMBRE_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblHelm_Click()
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_HELM_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_HELM_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_HELM_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_HELM_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub


Private Sub lblMana_Click()
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_MANA_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_MANA_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_MANA_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_MANA_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblOpciones_Click()
    Call frmOpciones.Show(vbModeless, frmMain)
End Sub

Private Sub lblSed_Click()
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_SED_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_SED_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_SED_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_SED_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblShielder_Click()
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_SHIELDER_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_SHIELDER_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_SHIELDER_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_SHIELDER_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblStrg_Click()
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_STRG_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_STRG_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_STRG_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_STRG_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblVida_Click()
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_VIDA_LABEL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_VIDA_LABEL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_VIDA_LABEL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_VIDA_LABEL").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub lblMinimizar_Click()
    Me.WindowState = 1
End Sub


Private Sub lblWeapon_Click()
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_WEAPON_LABEL").item("TEXTO"), _
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
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_MACRO_ACTIVADO").item("TEXTO"), 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mWork, True)
End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_MACRO_DESACTIVADO").item("TEXTO"), 0, 200, 200, False, True, True)
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
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(3), _
                        False, False, True)
End Sub

Private Sub Coord_Click()
    Call frmMain.AddtoRichPicture(JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub pHechizos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    hlst.ListIndex = (Y - SPELLS_PADDING) \ hlst.Pixel_Alto + hlst.Scroll
    
    ' Mover hechizos con shift + clic
    With hlst

        If ChangeHechi Then
    
            Dim NewLugar As Integer: NewLugar = .ListIndex
            Dim AntLugar As String: AntLugar = .List(NewLugar)
            
            Call WriteDragAndDropHechizos(ChangeHechiNum + 1, NewLugar + 1)
        
            .ForeColor = vbWhite
            .List(NewLugar) = .List(ChangeHechiNum)
            .List(ChangeHechiNum) = AntLugar
        
            ChangeHechi = False
            ChangeHechiNum = 0
            
        ElseIf Shift <> 0 Then
        
            ChangeHechi = True
            ChangeHechiNum = .ListIndex
            .ForeColor = vbRed

        End If

    End With
End Sub

Private Sub pHechizos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        If Y < SPELLS_PADDING Then
            If hlst.ListIndex > 0 Then
                hlst.ListIndex = hlst.ListIndex - 1
                
                If hlst.ListCount <= SPELLS_VISIBLE_LINES Then
                    BarritaHechizos.Top = BarraHechizosCentro.Top
                Else
                    BarritaHechizos.Top = BarraHechizosCentro.Top + hlst.Scroll * (BarraHechizosCentro.Height - BarritaHechizos.Height) \ (hlst.ListCount - SPELLS_VISIBLE_LINES)
                End If
            End If
        ElseIf Y > pHechizos.Height - SPELLS_PADDING Then
            If hlst.ListIndex < hlst.ListCount Then
                hlst.ListIndex = hlst.ListIndex + 1
                
                If hlst.ListCount <= SPELLS_VISIBLE_LINES Then
                    BarritaHechizos.Top = BarraHechizosCentro.Top
                Else
                    BarritaHechizos.Top = BarraHechizosCentro.Top + hlst.Scroll * (BarraHechizosCentro.Height - BarritaHechizos.Height) \ (hlst.ListCount - SPELLS_VISIBLE_LINES)
                End If
            End If
        Else
            hlst.ListIndex = (Y - SPELLS_PADDING) \ hlst.Pixel_Alto + hlst.Scroll
        End If
    End If
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

Private Sub pConsola_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    StartCheckingLinks
End Sub

Private Sub SendCMSTXT_GotFocus()
    SendCMSTXTHasFocus = True
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

Private Sub SendCMSTXT_LostFocus()
    SendCMSTXTHasFocus = False
End Sub

Private Sub SendTxt_GotFocus()
 SendTxtHasFocus = True
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
            
            'Como es KeyDown, si mantenes apretado el mensaje llena la consola

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

        Call SendTxt_SendText
        KeyCode = 0
    End If
End Sub

Public Sub SendTxt_SendText()
'**************************************************************
'Author: Unknown
'Last Modify Date: 04/01/2020
'08/01/2020: cucsifae - colapse en una funcion el mandar mensaje, en caso de no tener focus y apretar enter el mensaje se manda igual desde el KeyUp del mainform
'**************************************************************
        'Say
        If LenB(stxtbuffer) <> 0 Then
            Call ParseUserCommand(stxtbuffer)
        End If

        stxtbuffer = vbNullString
        SendTxt.Text = vbNullString
        Me.SendTxt.Visible = False
        
        If Typing Then
            Call WriteSetTypingFlagFromUserCharIndex
            Typing = False
        End If
        
End Sub

Private Sub Second_Timer()

    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer

    If MsgTimeadoOn Then
        Call ShowConsoleMsg(MsgTimeado)
        MsgTimeadoOn = False
    End If

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

Private Sub btnLanzar_Click()
    
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

Private Sub btnLanzar_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub btnInfo_Click()
    
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
                            
                            Call frmMain.AddtoRichPicture(MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
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
                                
                                Call frmMain.AddtoRichPicture(MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
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
                                    
                                    Call frmMain.AddtoRichPicture(MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
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
                                    
                                    Call frmMain.AddtoRichPicture(MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
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

Private Sub lblDropGold_Click()

    Inventario.SelectGold

    If UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If
    
End Sub

Private Sub btnInventario_Click()
    Call Audio.PlayWave(SND_CLICK)

    ' Activo controles de inventario
    PicInv.Visible = True

    ' Desactivo controles de hechizo
    hlst.Visible = False
    btnInfo.Visible = False
    btnLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    
    DoEvents
    Call Inventario.DrawInventory
    
End Sub

Private Sub btnHechizos_Click()
    
    Call Audio.PlayWave(SND_CLICK)

    ' Activo controles de hechizos
    hlst.Visible = True
    btnInfo.Visible = True
    btnLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    ' Desactivo controles de inventario
    PicInv.Visible = False

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
        
        Case eObjType.otcasco, eObjType.otAnillo, eObjType.otArmadura, eObjType.otescudo, eObjType.otFlechas, eObjType.otMochilas
            Call EquiparItem
    
        Case eObjType.otWeapon
            'Para los arcos y cuchillas hacemos esta validacion, asi se pueden usar con doble click en ves de andar equipando o desequipando (Recox)
            If InStr(Inventario.ItemName(Inventario.SelectedItem), "Arco") > 0 Or _
               InStr(Inventario.ItemName(Inventario.SelectedItem), "Cuchillas") > 0 Then
                If Inventario.Equipped(Inventario.SelectedItem) Then
                    Call UsarItem
                    Exit Sub
                End If
            End If

            Call EquiparItem
        
        Case Else
            Call UsarItem
            
    End Select
    
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
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
        Call SendCMSTXT_SendText
        KeyCode = 0 'esto no deberia ser necesario no se esta pasando el keycode por ref, no le encuentro sentido ponerlo en 0.
    End If
    
End Sub
Public Sub SendCMSTXT_SendText()
'**************************************************************
'Author: Unknown
'Last Modify Date: 04/01/2020
'08/01/2020: cucsifae - colapse en una funcion el mandar mensaje, en caso de no tener focus y apretar enter el mensaje se manda igual desde el KeyUp del mainform
'**************************************************************
        'Say
        If LenB(stxtbuffercmsg) <> 0 Then
            Call WriteGuildMessage(stxtbuffercmsg)
        End If

        stxtbuffercmsg = vbNullString
        SendCMSTXT.Text = vbNullString
        Me.SendCMSTXT.Visible = False
        
        If Typing Then
            Call WriteSetTypingFlagFromUserCharIndex
            Typing = False
        End If
        
        If PicInv.Visible Then
            PicInv.SetFocus
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
    'TODO: No usar variable de compilacion y acceder a esto desde el config.ini
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

Private Sub SendTxt_LostFocus()
    SendTxtHasFocus = False
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
    Security.Redundance = 13
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

'***************************************************
'Incorporado por ReyarB
'Last Modify Date: 21/05/2020 (ReyarB)
'Ajustadas las coordenadas (ReyarB)
'***************************************************
Private Sub Minimapa_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
   If X > 87 Then X = 86
   If X < 14 Then X = 15
   If Y > 90 Then Y = 89
   If Y < 11 Then Y = 12

   If Button = vbRightButton Then
      Call WriteWarpChar("YO", UserMap, CByte(X - 1), CByte(Y - 1))
      Call ActualizarMiniMapa
   End If
End Sub

Public Sub ActualizarMiniMapa()
    '***************************************************
    'Author: Martin Gomez (Samke)
    'Last Modify Date: 21/03/2020 (ReyarB)
    'Integrado por Reyarb
    'Se agrego campo de vision del render (Recox)
    'Ajustadas las coordenadas para centrarlo (WyroX)
    'Ajuste de coordenadas y tamaño del visor (ReyarB)
    '***************************************************
    Me.UserM.Left = UserPos.X - 2
    Me.UserM.Top = UserPos.Y - 2
    Me.UserAreaMinimap.Left = UserPos.X - 13
    Me.UserAreaMinimap.Top = UserPos.Y - 11
    Me.MiniMapa.Refresh
End Sub

Public Sub ActivarMacroHechizos()

    If Not hlst.Visible Then
        Call frmMain.AddtoRichPicture("Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, True)
        Exit Sub
    End If
    
    trainingMacro.Interval = INT_MACRO_HECHIS
    trainingMacro.Enabled = True
    Call frmMain.AddtoRichPicture("Auto lanzar hechizos activado", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mSpells, True)
End Sub

Public Sub DesactivarMacroHechizos()
    trainingMacro.Enabled = False
    Call frmMain.AddtoRichPicture("Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True)
    Call ControlSM(eSMType.mSpells, False)
End Sub

Private Sub timerPasarSegundo_Timer()
    
    If UserEstado = 0 Then
    
        If UserInvisible And TiempoInvi > 0 Then
            TiempoInvi = TiempoInvi - 1
        End If
        
        If TiempoDopas > 0 Then
            TiempoDopas = TiempoDopas - 1
        End If
    
        If UserParalizado And UserParalizadoSegundosRestantes > 0 Then
            UserParalizadoSegundosRestantes = UserParalizadoSegundosRestantes - 1
        End If
    
        If Not UserEquitando And UserEquitandoSegundosRestantes > 0 Then
            UserEquitandoSegundosRestantes = UserEquitandoSegundosRestantes - 1
        End If
        
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
        frmMain.uAOProgressExperienceLevel.max = 100
        frmMain.uAOProgressExperienceLevel.Value = 100
    Else
        frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
        frmMain.uAOProgressExperienceLevel.max = UserPasarNivel
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

Private Sub btnRetos_Click()
    Call FrmRetos.Show(vbModeless, frmMain)
End Sub

Private Sub ReDrawConsola()
    pConsola.Cls
    Dim i As Integer
    Dim Lines As Integer
    
    Lines = IIf(LineasConsola > CONSOLE_VISIBLE_LINES, OffSetConsola + CONSOLE_VISIBLE_LINES - 1, LineasConsola - 1)
    
    For i = OffSetConsola To Lines
        pConsola.CurrentX = CONSOLE_PADDING
        pConsola.CurrentY = (i - OffSetConsola) * CONSOLE_LINE_HEIGHT + CONSOLE_PADDING
        pConsola.ForeColor = Consola(i).Color
        pConsola.FontBold = CBool(Consola(i).bold)
        pConsola.FontItalic = CBool(Consola(i).italic)
        pConsola.Print Consola(i).Texto
    Next i
    
    If LineasConsola <= CONSOLE_VISIBLE_LINES Then
        BarritaConsola.Top = BarraConsolaCentro.Top + BarraConsolaCentro.Height - BarritaConsola.Height
    Else
        BarritaConsola.Top = BarraConsolaCentro.Top + OffSetConsola * (BarraConsolaCentro.Height - BarritaConsola.Height) \ (LineasConsola - CONSOLE_VISIBLE_LINES)
    End If
End Sub

Public Sub AddtoRichPicture(ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
    Dim AText As String
    Dim curLine As Integer
    Dim Lineas() As String
    Dim i As Integer
    Dim l As Integer
    Dim LastEsp As Integer
    
    curLine = LineasConsola
    LineasConsola = LineasConsola + 1

    Lineas = Split(Text, vbCrLf)

    For l = 0 To UBound(Lineas)
        If LineasConsola = MAX_CONSOLE_LINES Then
            For i = 0 To (MAX_CONSOLE_LINES - CONSOLE_REMOVED_LINES)
                Consola(i) = Consola(i + CONSOLE_REMOVED_LINES)
            Next i
            
            LineasConsola = MAX_CONSOLE_LINES - CONSOLE_REMOVED_LINES
            
            If OffSetConsola >= CONSOLE_REMOVED_LINES Then
                OffSetConsola = OffSetConsola - CONSOLE_REMOVED_LINES
            End If
        End If

        Text = Lineas(l)

        Consola(curLine).Texto = Text
        Consola(curLine).Color = RGB(Red, Green, Blue)
        Consola(curLine).bold = bold
        Consola(curLine).italic = italic

        If LineasConsola > CONSOLE_VISIBLE_LINES And OffSetConsola = curLine - CONSOLE_VISIBLE_LINES Then
            OffSetConsola = LineasConsola - CONSOLE_VISIBLE_LINES
        End If
        
        Dim MaxWidth As Integer
        MaxWidth = pConsola.Width - BarraConsolaCentro.Width - CONSOLE_PADDING * 2

        If pConsola.TextWidth(Text) > MaxWidth Then
            LastEsp = 0
            
            For i = 1 To Len(Text)
                If mid(Text, i, 1) = " " Then LastEsp = i
                If pConsola.TextWidth(Left$(Text, i)) > MaxWidth Then Exit For
            Next i

            If LastEsp = 0 Then LastEsp = i - 1
            
            AText = Right$(Text, Len(Text) - LastEsp)
            Text = Left$(Text, LastEsp)
            Consola(curLine).Texto = Text
            
            Call frmMain.AddtoRichPicture(AText, Red, Green, Blue, bold, italic)
        Else
            ReDrawConsola
        End If
    Next l
End Sub

Public Sub ClearConsole()
    LineasConsola = 0
    OffSetConsola = 0
    
    ReDrawConsola
End Sub
