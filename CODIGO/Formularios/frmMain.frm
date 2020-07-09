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
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":7F6A
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1023
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer timerPasarSegundo 
      Interval        =   1000
      Left            =   960
      Top             =   2880
   End
   Begin AOLibre.uAOProgress uAOProgressExperienceLevel 
      Height          =   255
      Left            =   12240
      TabIndex        =   40
      ToolTipText     =   "Experiencia necesaria para pasar de nivel"
      Top             =   1500
      Width           =   2415
      _ExtentX        =   4260
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
      BackColor       =   &H80000001&
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
      Left            =   9675
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   420
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
      BackColor       =   &H80000001&
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
      BackColor       =   &H80000001&
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3370
      Left            =   12165
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   16
      Top             =   2550
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
      Height          =   360
      Left            =   360
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "frmMain.frx":4E535
      ToolTipText     =   "Chat"
      Top             =   2400
      Visible         =   0   'False
      Width           =   10695
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
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "frmMain.frx":4E565
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
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1665
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   300
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   2937
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":4E59B
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
      Height          =   2790
      Left            =   12012
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2550
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9120
      Left            =   180
      MousePointer    =   99  'Custom
      ScaleHeight     =   608
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   736
      TabIndex        =   30
      Top             =   2325
      Width           =   11040
      Begin VB.Timer tmrCounters 
         Left            =   5760
         Top             =   840
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
      Left            =   11880
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   8520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Mapa"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":4E618
      PICF            =   "frmMain.frx":4F042
      PICH            =   "frmMain.frx":4FD04
      PICV            =   "frmMain.frx":50C96
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
      Left            =   11880
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   8880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Grupo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":51B98
      PICF            =   "frmMain.frx":525C2
      PICH            =   "frmMain.frx":53284
      PICV            =   "frmMain.frx":54216
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
      Left            =   13440
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Opciones"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":55118
      PICF            =   "frmMain.frx":55B42
      PICH            =   "frmMain.frx":56804
      PICV            =   "frmMain.frx":57796
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
      Left            =   11880
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   9240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Estadisticas"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":58698
      PICF            =   "frmMain.frx":590C2
      PICH            =   "frmMain.frx":59D84
      PICV            =   "frmMain.frx":5AD16
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
      Left            =   11880
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Clanes"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":5BC18
      PICF            =   "frmMain.frx":5C642
      PICH            =   "frmMain.frx":5D304
      PICV            =   "frmMain.frx":5E296
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
      Height          =   495
      Left            =   11880
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      TX              =   "Inventario"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":5F198
      PICF            =   "frmMain.frx":5FBC2
      PICH            =   "frmMain.frx":60884
      PICV            =   "frmMain.frx":61816
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
      Height          =   495
      Left            =   13440
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      TX              =   "Hechizos"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":62718
      PICF            =   "frmMain.frx":63142
      PICH            =   "frmMain.frx":63E04
      PICV            =   "frmMain.frx":64D96
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
      Height          =   495
      Left            =   11880
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      TX              =   "Lanzar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":65C98
      PICF            =   "frmMain.frx":666C2
      PICH            =   "frmMain.frx":67384
      PICV            =   "frmMain.frx":68316
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
      Height          =   495
      Left            =   13680
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      TX              =   "Info"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":69218
      PICF            =   "frmMain.frx":69C42
      PICH            =   "frmMain.frx":6A904
      PICV            =   "frmMain.frx":6B896
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
      Left            =   13440
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   8880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Retos"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":6C798
      PICF            =   "frmMain.frx":6D1C2
      PICH            =   "frmMain.frx":6DE84
      PICV            =   "frmMain.frx":6EE16
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
      Left            =   13440
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   8520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Amigos"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":6FD18
      PICF            =   "frmMain.frx":70742
      PICH            =   "frmMain.frx":71404
      PICV            =   "frmMain.frx":72396
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
      Left            =   13440
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      TX              =   "Quests"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmMain.frx":73298
      PICF            =   "frmMain.frx":73CC2
      PICH            =   "frmMain.frx":74984
      PICV            =   "frmMain.frx":75916
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
   Begin VB.Image BarritaMover 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   0
      Top             =   0
      Width           =   14415
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
      Left            =   13800
      TabIndex        =   42
      Top             =   1275
      Width           =   555
   End
   Begin VB.Label lblMapName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del mapa"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   4920
      TabIndex        =   39
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Image imgAsignarSkill 
      Height          =   405
      Left            =   13200
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   405
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13800
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   14460
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   14880
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   -120
      Width           =   375
   End
   Begin VB.Label lblFPS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "101"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   6465
      TabIndex        =   20
      Top             =   30
      Width           =   795
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   0
      Left            =   14790
      MouseIcon       =   "frmMain.frx":76818
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":7696A
      Top             =   3960
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   1
      Left            =   14790
      MouseIcon       =   "frmMain.frx":76CAE
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":76E00
      Top             =   3705
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   12000
      TabIndex        =   41
      Top             =   330
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
      Left            =   13080
      TabIndex        =   19
      Top             =   1245
      Width           =   270
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      ForeColor       =   &H80000018&
      Height          =   225
      Left            =   12480
      TabIndex        =   18
      Top             =   1245
      Width           =   465
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0080FFFF&
      Height          =   210
      Left            =   14160
      TabIndex        =   15
      Top             =   7020
      Width           =   90
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   12855
      TabIndex        =   9
      Top             =   6930
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   12240
      TabIndex        =   8
      Top             =   6930
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
      Left            =   11790
      TabIndex        =   7
      Top             =   11160
      Width           =   1455
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
      Left            =   13920
      TabIndex        =   6
      Top             =   10605
      Width           =   975
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
      Left            =   13920
      TabIndex        =   5
      Top             =   10155
      Width           =   975
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
      Left            =   12360
      TabIndex        =   4
      Top             =   10605
      Width           =   975
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
      Left            =   12360
      TabIndex        =   3
      Top             =   10155
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
      Top             =   1920
      Width           =   2970
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   13560
      TabIndex        =   11
      Top             =   7740
      Width           =   1335
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   12000
      TabIndex        =   10
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   13560
      TabIndex        =   12
      Top             =   8190
      Width           =   1335
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   12000
      TabIndex        =   13
      Top             =   7890
      Width           =   1215
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   12000
      TabIndex        =   14
      Top             =   8205
      Width           =   1215
   End
   Begin VB.Shape shpEnergia 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   11967
      Top             =   7560
      Width           =   1245
   End
   Begin VB.Shape shpMana 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   13543
      Top             =   7725
      Width           =   1350
   End
   Begin VB.Shape shpVida 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   13543
      Top             =   8175
      Width           =   1350
   End
   Begin VB.Shape shpHambre 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   11967
      Top             =   7905
      Width           =   1245
   End
   Begin VB.Shape shpSed 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   11967
      Top             =   8220
      Width           =   1245
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

Dim BoldX As Long, BoldY As Long, BisMoving As Boolean

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

'Para cuando se necesite enviar un mensaje a la consola dentro de un bucle al que no debas parar
Public MsgTimeadoOn        As Boolean
Public MsgTimeado          As String

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

Private Sub btnAmigos_Click()
    Call frmAmigos.Show(vbModeless, frmMain)
End Sub

Private Sub btnQuests_Click()
    Call ParseUserCommand("/INFOQUEST")
End Sub

Private Sub Form_Activate()
    Call Inventario.DrawInventory
End Sub

Private Sub BarritaMover_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not ResolucionCambiada Then
        BoldX = x
        BoldY = y
        BisMoving = True
    End If
End Sub
Private Sub BarritaMover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If BisMoving Then
        Me.Top = Me.Top - (BoldY - y)
        Me.Left = Me.Left - (BoldX - x)
    End If
End Sub
Private Sub BarritaMover_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BisMoving = False
End Sub

Private Sub Form_Load()
    SkinSeleccionado = GetVar(Game.path(INIT) & "Config.ini", "Parameters", "SkinSelected")
    
    Me.Picture = LoadPicture(Game.path(Skins) & SkinSeleccionado & "\VentanaPrincipal.jpg")

    If ResolucionCambiada Then
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    clicX = x
    clicY = y
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
                             x As Single, _
                             y As Single)
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
                             x As Single, _
                             y As Single)
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
                                x As Single, _
                                y As Single)
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
                                  x As Single, _
                                  y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
    MouseX = x
    MouseY = y
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
    clicX = x
    clicY = y
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x - MainViewPic.Left
    MouseY = y - MainViewPic.Top
    
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

Private Sub btnInventario_Click()
    Call Audio.PlayWave(SND_CLICK)

    ' Activo controles de inventario
    picInv.Visible = True

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
    picInv.Visible = False

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
        
        Case eObjType.otcasco, eObjType.otAnillo, eObjType.otArmadura, eObjType.otescudo, eObjType.otFlechas
            Call EquiparItem
    
        Case eObjType.otWeapon
            'Para los arcos hacemos esta validacion, asi se pueden usar con doble click en ves de andar equipando o desequipando (Recox)
            If InStr(Inventario.ItemName(Inventario.SelectedItem), "Arco") > 0 Then
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

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

        If picInv.Visible Then
            picInv.SetFocus
                        
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

'***************************************************
'Incorporado por ReyarB
'Last Modify Date: 21/05/2020 (ReyarB)
'Ajustadas las coordenadas (ReyarB)
'***************************************************
Private Sub Minimapa_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)
   If x > 87 Then x = 86
   If x < 14 Then x = 15
   If y > 90 Then y = 89
   If y < 11 Then y = 12

   If Button = vbRightButton Then
      Call WriteWarpChar("YO", UserMap, CByte(x - 1), CByte(y - 1))
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
    Me.UserM.Left = UserPos.x - 2
    Me.UserM.Top = UserPos.y - 2
    Me.UserAreaMinimap.Left = UserPos.x - 13
    Me.UserAreaMinimap.Top = UserPos.y - 11
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
