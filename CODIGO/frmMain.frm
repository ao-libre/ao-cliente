VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   8700
   ClientLeft      =   360
   ClientTop       =   300
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   580
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6750
      Top             =   1920
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   11325
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   30
      Top             =   8445
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   10950
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   29
      Top             =   8445
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
      Left            =   10575
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   28
      Top             =   8445
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
      Left            =   10200
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   27
      Top             =   8445
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
      Height          =   1920
      Left            =   9000
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   16
      Top             =   2400
      Width           =   2400
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5760
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   150
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1839
      Visible         =   0   'False
      Width           =   8190
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   7080
      Top             =   2520
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   6600
      Top             =   2520
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   8520
      Visible         =   0   'False
      Width           =   8250
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   5760
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6240
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   4920
      Top             =   2520
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4200
      Top             =   2520
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1478
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   360
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   2619
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":030A
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
      Left            =   8880
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Image imgMapa 
      Height          =   300
      Left            =   10230
      Top             =   6675
      Width           =   1395
   End
   Begin VB.Image imgClanes 
      Height          =   390
      Left            =   10245
      Top             =   7980
      Width           =   1380
   End
   Begin VB.Image imgEstadisticas 
      Height          =   360
      Left            =   10215
      Top             =   7605
      Width           =   1425
   End
   Begin VB.Image imgOpciones 
      Height          =   330
      Left            =   10215
      Top             =   7275
      Width           =   1425
   End
   Begin VB.Image imgGrupo 
      Height          =   315
      Left            =   10215
      Top             =   6990
      Width           =   1410
   End
   Begin VB.Image imgAsignarSkill 
      Height          =   450
      Left            =   9975
      MousePointer    =   99  'Custom
      Top             =   690
      Width           =   450
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   10560
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   11280
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   60
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   11550
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10320
      MouseIcon       =   "frmMain.frx":0388
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Label lblFPS 
      BackStyle       =   0  'Transparent
      Caption         =   "101"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   6705
      TabIndex        =   26
      Top             =   60
      Width           =   555
   End
   Begin VB.Image imgInvScrollDown 
      Height          =   795
      Left            =   10020
      Top             =   5595
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgInvScrollUp 
      Height          =   1065
      Left            =   10005
      Top             =   4545
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image cmdInfo 
      Height          =   405
      Left            =   10680
      MouseIcon       =   "frmMain.frx":04DA
      MousePointer    =   99  'Custom
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblMapName 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8700
      TabIndex        =   25
      Top             =   8445
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   0
      Left            =   11430
      MouseIcon       =   "frmMain.frx":062C
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":077E
      Top             =   3480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   1
      Left            =   11430
      MouseIcon       =   "frmMain.frx":0AC2
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":0C14
      Top             =   3225
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image xz 
      Height          =   255
      Index           =   0
      Left            =   11400
      Top             =   0
      Width           =   255
   End
   Begin VB.Image xzz 
      Height          =   195
      Index           =   1
      Left            =   11445
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8760
      TabIndex        =   24
      Top             =   360
      Width           =   2625
   End
   Begin VB.Label lblLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9825
      TabIndex        =   23
      Top             =   1140
      Width           =   270
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9225
      TabIndex        =   22
      Top             =   1140
      Width           =   465
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
      Left            =   9225
      TabIndex        =   21
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label lblExp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp: 999999999/99999999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9225
      TabIndex        =   20
      Top             =   1500
      Width           =   2010
   End
   Begin VB.Image CmdLanzar 
      Height          =   375
      Left            =   8760
      MouseIcon       =   "frmMain.frx":0F58
      MousePointer    =   99  'Custom
      Top             =   5280
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8880
      MouseIcon       =   "frmMain.frx":10AA
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   10920
      TabIndex        =   15
      Top             =   6270
      Width           =   90
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   9720
      TabIndex        =   9
      Top             =   6225
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   9120
      TabIndex        =   8
      Top             =   6225
      Width           =   210
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000 X:00 Y: 00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8700
      TabIndex        =   7
      Top             =   8670
      Width           =   1335
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
      Top             =   8640
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5130
      TabIndex        =   5
      Top             =   8640
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2940
      TabIndex        =   4
      Top             =   8640
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1170
      TabIndex        =   3
      Top             =   8640
      Width           =   855
   End
   Begin VB.Image imgScroll 
      Height          =   240
      Index           =   1000
      Left            =   11400
      MousePointer    =   99  'Custom
      Top             =   3225
      Width           =   225
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      Height          =   6240
      Left            =   180
      Top             =   2235
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Image InvEqu 
      Height          =   4230
      Left            =   8715
      Top             =   1875
      Width           =   2970
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8760
      TabIndex        =   11
      Top             =   7140
      Width           =   1095
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8760
      TabIndex        =   10
      Top             =   6780
      Width           =   1095
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8760
      TabIndex        =   12
      Top             =   7455
      Width           =   1095
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8760
      TabIndex        =   13
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8760
      TabIndex        =   14
      Top             =   8130
      Width           =   1095
   End
   Begin VB.Shape shpEnergia 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   8760
      Top             =   6795
      Width           =   1125
   End
   Begin VB.Shape shpMana 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   8760
      Top             =   7155
      Width           =   1125
   End
   Begin VB.Shape shpVida 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   8760
      Top             =   7470
      Width           =   1125
   End
   Begin VB.Shape shpHambre 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   8760
      Top             =   7815
      Width           =   1125
   End
   Begin VB.Shape shpSed 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   8760
      Top             =   8160
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

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Private clsFormulario As clsFormMovementManager

Private cBotonDiamArriba As clsGraphicalButton
Private cBotonDiamAbajo As clsGraphicalButton
Private cBotonMapa As clsGraphicalButton
Private cBotonGrupo As clsGraphicalButton
Private cBotonOpciones As clsGraphicalButton
Private cBotonEstadisticas As clsGraphicalButton
Private cBotonClanes As clsGraphicalButton
Private cBotonAsignarSkill As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public picSkillStar As Picture

Dim PuedeMacrear As Boolean

Private Sub Form_Load()
    
    If NoRes Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me, 120
    End If

    Me.Picture = LoadPicture(DirGraficos & "VentanaPrincipal.JPG")
    
    InvEqu.Picture = LoadPicture(DirGraficos & "CentroInventario.jpg")
    
    Call LoadButtons
    
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    Dim i As Integer
    
    GrhPath = DirGraficos

    Set cBotonDiamArriba = New clsGraphicalButton
    Set cBotonDiamAbajo = New clsGraphicalButton
    Set cBotonGrupo = New clsGraphicalButton
    Set cBotonOpciones = New clsGraphicalButton
    Set cBotonEstadisticas = New clsGraphicalButton
    Set cBotonClanes = New clsGraphicalButton
    Set cBotonAsignarSkill = New clsGraphicalButton
    Set cBotonMapa = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonDiamArriba.Initialize(imgInvScrollUp, "", _
                                    GrhPath & "BotonDiamArribaF.bmp", _
                                    GrhPath & "BotonDiamArribaF.bmp", Me)

    Call cBotonDiamAbajo.Initialize(imgInvScrollDown, "", _
                                    GrhPath & "BotonDiamAbajoF.bmp", _
                                    GrhPath & "BotonDiamAbajoF.bmp", Me)
    
    Call cBotonMapa.Initialize(imgMapa, "", _
                                    GrhPath & "BotonMapaRollover.jpg", _
                                    GrhPath & "BotonMapaClick.jpg", Me)
                                    
    Call cBotonGrupo.Initialize(imgGrupo, "", _
                                    GrhPath & "BotonGrupoRollover.jpg", _
                                    GrhPath & "BotonGrupoClick.jpg", Me)

    Call cBotonOpciones.Initialize(imgOpciones, "", _
                                    GrhPath & "BotonOpcionesRollover.jpg", _
                                    GrhPath & "BotonOpcionesClick.jpg", Me)

    Call cBotonEstadisticas.Initialize(imgEstadisticas, "", _
                                    GrhPath & "BotonEstadisticasRollover.jpg", _
                                    GrhPath & "BotonEstadisticasClick.jpg", Me)

    Call cBotonClanes.Initialize(imgClanes, "", _
                                    GrhPath & "BotonClanesRollover.jpg", _
                                    GrhPath & "BotonClanesClick.jpg", Me)

    Set picSkillStar = LoadPicture(GrhPath & "BotonAsignarSkills.bmp")

    If SkillPoints > 0 Then imgAsignarSkill.Picture = picSkillStar
    
    imgAsignarSkill.MouseIcon = picMouseIcon
    lblDropGold.MouseIcon = picMouseIcon
    lblCerrar.MouseIcon = picMouseIcon
    lblMinimizar.MouseIcon = picMouseIcon
    
    For i = 0 To 3
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

Public Sub ActivarMacroHechizos()
    If Not hlst.Visible Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, True)
        Exit Sub
    End If
    
    TrainingMacro.Interval = INT_MACRO_HECHIS
    TrainingMacro.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mSpells, True)
End Sub

Public Sub DesactivarMacroHechizos()
    TrainingMacro.Enabled = False
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True)
    Call ControlSM(eSMType.mSpells, False)
End Sub

Public Sub ControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)
Dim GrhIndex As Long
Dim SR As RECT
Dim DR As RECT

GrhIndex = GRH_INI_SM + Index + SM_CANT * (CInt(Mostrar) + 1)

With GrhData(GrhIndex)
    SR.Left = .sX
    SR.Right = SR.Left + .pixelWidth
    SR.Top = .sY
    SR.Bottom = SR.Top + .pixelHeight
    
    DR.Left = 0
    DR.Right = .pixelWidth
    DR.Top = 0
    DR.Bottom = .pixelHeight
End With

Call DrawGrhtoHdc(picSM(Index).hdc, GrhIndex, SR, DR)
picSM(Index).Refresh

Select Case Index
    Case eSMType.sResucitation
        If Mostrar Then
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, True)
            picSM(Index).ToolTipText = "Seguro de resucitación activado."
        Else
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, True)
            picSM(Index).ToolTipText = "Seguro de resucitación desactivado."
        End If
        
    Case eSMType.sSafemode
        If Mostrar Then
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, True)
            picSM(Index).ToolTipText = "Seguro activado."
        Else
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, True)
            picSM(Index).ToolTipText = "Seguro desactivado."
        End If
        
    Case eSMType.mSpells
        If Mostrar Then
            picSM(Index).ToolTipText = "Macro de hechizos activado."
        Else
            picSM(Index).ToolTipText = "Macro de hechizos desactivado."
        End If
        
    Case eSMType.mWork
        If Mostrar Then
            picSM(Index).ToolTipText = "Macro de trabajo activado."
        Else
            picSM(Index).ToolTipText = "Macro de trabajo desactivado."
        End If
End Select

SMStatus(Index) = Mostrar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 18/11/2009
'18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
'***************************************************
#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If
    
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
        
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
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
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
            Select Case KeyCode
                'Custom messages!
                Case vbKey0 To vbKey9
                    Dim CustomMessage As String
                    
                    CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)
                    If LenB(CustomMessage) <> 0 Then
                        ' No se pueden mandar mensajes personalizados de clan o privado!
                        If UCase(Left(CustomMessage, 5)) <> "/CMSG" And _
                            Left(CustomMessage, 1) <> "\" Then
                            
                            Call ParseUserCommand(CustomMessage)
                        End If
                    End If
            End Select
        End If
    End If
    
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
            If SendTxt.Visible Then Exit Sub
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And _
              (Not frmMSG.Visible) And (Not MirandoForo) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
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
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
                
            If Not PuedeMacrear Then
                AddtoRichTextBox frmMain.RecTxt, "No tan rápido..!", 255, 255, 255, False, False, True
            Else
                Call WriteMeditate
                PuedeMacrear = False
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If TrainingMacro.Enabled Then
                DesactivarMacroHechizos
            Else
                ActivarMacroHechizos
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
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
            
            If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
            If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteAttack
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If SendCMSTXT.Visible Then Exit Sub
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And _
              (Not frmMSG.Visible) And (Not MirandoForo) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
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

Private Sub imgInvScrollDown_Click()
    Call Inventario.ScrollInventory(True)
End Sub

Private Sub imgInvScrollUp_Click()
    Call Inventario.ScrollInventory(False)
End Sub

Private Sub imgMapa_Click()
    Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub imgOpciones_Click()
    Call frmOpciones.Show(vbModeless, frmMain)
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub lblScroll_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub lblCerrar_Click()
    prgRun = False
End Sub

Private Sub lblMinimizar_Click()
    Me.WindowState = 1
End Sub

Private Sub Macro_Timer()
    PuedeMacrear = True
End Sub

Private Sub macrotrabajo_Timer()
    If Inventario.SelectedItem = 0 Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    If Not Application.IsAppActive() Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or _
                UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not frmHerrero.Visible) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
     If Not (frmCarp.Visible = True) Then Call UsarItem
End Sub

Public Sub ActivarMacroTrabajo()
    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mWork, True)
End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mWork, False)
End Sub


Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub PicMH_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, True)
End Sub

Private Sub Coord_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, True)
End Sub

Private Sub picSM_DblClick(Index As Integer)
Select Case Index
    Case eSMType.sResucitation
        Call WriteResuscitationToggle
        
    Case eSMType.sSafemode
        Call WriteSafeToggle
        
    Case eSMType.mSpells
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
        If TrainingMacro.Enabled Then
            Call DesactivarMacroHechizos
        Else
            Call ActivarMacroHechizos
        End If
        
    Case eSMType.mWork
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
        If macrotrabajo.Enabled Then
            Call DesactivarMacroTrabajo
        Else
            Call ActivarMacroTrabajo
        End If
End Select
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
        
        If PicInv.Visible Then
            PicInv.SetFocus
        Else
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

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
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
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
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem()
    If pausa Then Exit Sub
    
    If Comerciando Then Exit Sub
    
    If TrainingMacro.Enabled Then DesactivarMacroHechizos
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
End Sub

Private Sub EquiparItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
    End If
End Sub



''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()
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
    
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    If UsingSkill = Magia And Not MainTimer.Check(TimersIndex.CastSpell) Then Exit Sub
    
    If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub
    
    Call WriteWorkLeftClick(tX, tY, UsingSkill)
    UsingSkill = 0
End Sub

Private Sub cmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub cmdINFO_Click()
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
End Sub

Private Sub DespInv_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub Form_Click()
    If Cartel Then Cartel = False

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
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
                    Call WriteLeftClick(tX, tY)
                Else
                
                    If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
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
                                    Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rápido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
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
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
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
        Call WriteDoubleClick(tX, tY)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewShp.Left
    MouseY = Y - MainViewShp.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewShp.Width Then
        MouseX = MainViewShp.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewShp.Height Then
        MouseY = MainViewShp.Height
    End If
    
    LastPressed.ToggleToNormal
    
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

Private Sub Label4_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centroinventario.jpg")

    ' Activo controles de inventario
    PicInv.Visible = True
    imgInvScrollUp.Visible = True
    imgInvScrollDown.Visible = True

    ' Desactivo controles de hechizo
    hlst.Visible = False
    cmdINFO.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centrohechizos.jpg")
    
    ' Activo controles de hechizos
    hlst.Visible = True
    cmdINFO.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    ' Desactivo controles de inventario
    PicInv.Visible = False
    imgInvScrollUp.Visible = False
    imgInvScrollDown.Visible = False

End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
    
    Call UsarItem
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
    ElseIf (Not Comerciando) And (Not MirandoAsignarSkills) And _
        (Not frmMSG.Visible) And (Not MirandoForo) And _
        (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
         
        If PicInv.Visible Then
            PicInv.SetFocus
        ElseIf hlst.Visible Then
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If PicInv.Visible Then
        PicInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
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
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
        
        If PicInv.Visible Then
            PicInv.SetFocus
        Else
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
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


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    
#If SeguridadAlkon Then
    Call ConnectionStablished(Socket1.PeerAddress)
#End If
    
    Second.Enabled = True

    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
        
        Case E_MODO.Normal
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
        
        Case E_MODO.Dados
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Audio.PlayMIDI("7.mid")
            frmCrearPersonaje.Show vbModal
    End Select
End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long
    
    Second.Enabled = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    Do While i < Forms.Count - 1
        i = i + 1
        
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name And Forms(i).Name <> frmCrearPersonaje.Name Then
            Unload Forms(i)
        End If
    Loop
    
    On Local Error GoTo 0
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If
    
    frmMain.Visible = False
    
    pausa = False
    UserMeditar = False
    
#If SeguridadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    For i = 1 To MAX_INVENTORY_SLOTS
        
    Next i
    
    macrotrabajo.Enabled = False

    SkillPoints = 0
    Alocados = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim RD As String
    Dim data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub
    
#If SeguridadAlkon Then
    Call DataReceived(data)
#End If
    
    'Put data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim M As New frmMenuseFashion
            
            Load M
            M.SetCallback Me
            M.SetMenuId 1
            M.ListaInit 2, False
            
            If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
                M.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
            Else
                M.ListaSetItem 0, "<NPC>", True
            End If
            M.ListaSetItem 1, "Comerciar"
            
            M.ListaFin
            M.Show , Me

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
        Call WriteLeftClick(tX, tY)
        
    Case 1 'Comerciar
        Call WriteLeftClick(tX, tY)
        Call WriteCommerceStart
    End Select
End Select
End Sub


'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"
    
    Second.Enabled = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    Do While i < Forms.Count - 1
        i = i + 1
        
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name And Forms(i).Name <> frmCrearPersonaje.Name Then
            Unload Forms(i)
        End If
    Loop
    On Local Error GoTo 0
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Debug.Print "Winsock Connect"
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    
#If SeguridadAlkon Then
    Call ConnectionStablished(Winsock1.RemoteHostIP)
#End If
    
    Second.Enabled = True
    
    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login


        Case E_MODO.Normal
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login

        Case E_MODO.Dados
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Audio.PlayMIDI("7.mid")
            frmCrearPersonaje.Show vbModal
            
#If SeguridadAlkon Then
            Call ProtectForm(frmCrearPersonaje)
#End If
    End Select
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim data() As Byte
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD
    
    data = StrConv(RD, vbFromUnicode)
    
#If SeguridadAlkon Then
    Call DataReceived(data)
#End If
    
    'Set data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close

    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub
#End If

Private Function InGameArea() As Boolean
'***************************************************
'Author: NicoNZ
'Last Modification: 04/07/08
'Checks if last click was performed within or outside the game area.
'***************************************************
    If clicX < MainViewShp.Left Or clicX > MainViewShp.Left + MainViewShp.Width Then Exit Function
    If clicY < MainViewShp.Top Or clicY > MainViewShp.Top + MainViewShp.Height Then Exit Function
    
    InGameArea = True
End Function

Private Sub Winsock2_Connect()
#If SeguridadAlkon = 1 Then
    Call modURL.ProcessRequest
#End If
End Sub
