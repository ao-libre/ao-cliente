VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   990
      Left            =   7080
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   28
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picPJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   7080
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tAnimacion 
      Left            =   840
      Top             =   1080
   End
   Begin VB.ComboBox lstProfesion 
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
      ItemData        =   "frmCrearPersonaje.frx":53C8D
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":53C8F
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4035
      Width           =   2625
   End
   Begin VB.ComboBox lstGenero 
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
      ItemData        =   "frmCrearPersonaje.frx":53C91
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":53C9B
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4560
      Width           =   2625
   End
   Begin VB.ComboBox lstRaza 
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
      ItemData        =   "frmCrearPersonaje.frx":53CAE
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":53CB0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3495
      Width           =   2625
   End
   Begin VB.ComboBox lstHogar 
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
      ItemData        =   "frmCrearPersonaje.frx":53CB2
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":53CB4
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2985
      Width           =   2625
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000012&
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
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1320
      Width           =   5055
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   6795
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   2
      Left            =   7200
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   23
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   3
      Left            =   7605
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   24
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   4
      Left            =   8010
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   6390
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   21
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin AOLibre.uAOButton imgVolver 
      Height          =   495
      Left            =   1200
      TabIndex        =   29
      Top             =   8160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      TX              =   "Volver"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCrearPersonaje.frx":53CB6
      PICF            =   "frmCrearPersonaje.frx":546E0
      PICH            =   "frmCrearPersonaje.frx":553A2
      PICV            =   "frmCrearPersonaje.frx":56334
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
   Begin AOLibre.uAOButton imgCrear 
      Height          =   495
      Left            =   9120
      TabIndex        =   30
      Top             =   8160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      TX              =   "Crear Personaje"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCrearPersonaje.frx":57236
      PICF            =   "frmCrearPersonaje.frx":57C60
      PICH            =   "frmCrearPersonaje.frx":58922
      PICV            =   "frmCrearPersonaje.frx":598B4
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
   Begin AOLibre.uAOButton imgDados 
      Height          =   975
      Left            =   1320
      TabIndex        =   31
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      TX              =   "Tirar Dados"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCrearPersonaje.frx":5A7B6
      PICF            =   "frmCrearPersonaje.frx":5B1E0
      PICH            =   "frmCrearPersonaje.frx":5BEA2
      PICV            =   "frmCrearPersonaje.frx":5CE34
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
   Begin VB.Label imgEspecialidad 
      BackStyle       =   0  'Transparent
      Caption         =   "Especialidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3360
      TabIndex        =   49
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label imgArcos 
      BackStyle       =   0  'Transparent
      Caption         =   "Arcos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3360
      TabIndex        =   48
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label imgArmas 
      BackStyle       =   0  'Transparent
      Caption         =   "Armas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3360
      TabIndex        =   47
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label imgEscudos 
      BackStyle       =   0  'Transparent
      Caption         =   "Escudos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3360
      TabIndex        =   46
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label imgVida 
      BackStyle       =   0  'Transparent
      Caption         =   "Vida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3360
      TabIndex        =   45
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label imgMagia 
      BackStyle       =   0  'Transparent
      Caption         =   "Magia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3360
      TabIndex        =   44
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label imgEvasion 
      BackStyle       =   0  'Transparent
      Caption         =   "Evasion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3360
      TabIndex        =   43
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label imgConstitucion 
      BackStyle       =   0  'Transparent
      Caption         =   "Carisma"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3480
      TabIndex        =   42
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label imgGenero 
      BackStyle       =   0  'Transparent
      Caption         =   "Genero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6480
      TabIndex        =   41
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label imgClase 
      BackStyle       =   0  'Transparent
      Caption         =   "Clase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6480
      TabIndex        =   40
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label imgRaza 
      BackStyle       =   0  'Transparent
      Caption         =   "Raza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6960
      TabIndex        =   39
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label imgCarisma 
      BackStyle       =   0  'Transparent
      Caption         =   "Carisma"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3480
      TabIndex        =   38
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label imgInteligencia 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Inteligencia"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3360
      TabIndex        =   37
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label imgAgilidad 
      BackStyle       =   0  'Transparent
      Caption         =   "Agilidad"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3480
      TabIndex        =   36
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label imgFuerza 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuerza"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3480
      TabIndex        =   35
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label imgAtributos 
      BackStyle       =   0  'Transparent
      Caption         =   "Atributos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3600
      TabIndex        =   34
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label imgPuebloOrigen 
      BackStyle       =   0  'Transparent
      Caption         =   "Pueblo de Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6480
      TabIndex        =   33
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label imgNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   4560
      TabIndex        =   32
      Top             =   960
      Width           =   2895
   End
   Begin VB.Image ImgProfesionDibujo 
      Height          =   885
      Left            =   240
      MouseIcon       =   "frmCrearPersonaje.frx":5DD36
      MousePointer    =   99  'Custom
      Top             =   4680
      Width           =   900
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   7110
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   7110
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   7110
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   7110
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   7110
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   6825
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   6825
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   6825
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   6825
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   6540
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   6540
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   6540
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   6540
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   6255
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   6255
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   6255
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   6255
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   5970
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   5970
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   5970
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   5970
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   6825
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   6540
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   6255
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   5970
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   5
      Left            =   5400
      Top             =   5685
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   4
      Left            =   5175
      Top             =   5685
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   3
      Left            =   4950
      Top             =   5685
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   2
      Left            =   4725
      Top             =   5685
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   1
      Left            =   4500
      Top             =   5685
      Width           =   225
   End
   Begin VB.Label lblEspecialidad 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   26
      Top             =   7395
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   3
      Visible         =   0   'False
      X1              =   479
      X2              =   505
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   2
      Visible         =   0   'False
      X1              =   479
      X2              =   505
      Y1              =   391
      Y2              =   391
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   1
      Visible         =   0   'False
      X1              =   504
      X2              =   504
      Y1              =   392
      Y2              =   416
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   479
      X2              =   479
      Y1              =   392
      Y2              =   416
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   5
      Left            =   5445
      TabIndex        =   20
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   4
      Left            =   5445
      TabIndex        =   19
      Top             =   4470
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   3
      Left            =   5445
      TabIndex        =   18
      Top             =   4125
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   2
      Left            =   5445
      TabIndex        =   17
      Top             =   3780
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   1
      Left            =   5445
      TabIndex        =   16
      Top             =   3450
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   5
      Left            =   4950
      TabIndex        =   15
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   4
      Left            =   4950
      TabIndex        =   14
      Top             =   4470
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   3
      Left            =   4950
      TabIndex        =   13
      Top             =   4125
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   2
      Left            =   4950
      TabIndex        =   12
      Top             =   3780
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   1
      Left            =   4950
      TabIndex        =   11
      Top             =   3450
      Width           =   225
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   9480
      TabIndex        =   10
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Image imgF 
      Height          =   270
      Left            =   5415
      Top             =   3075
      Width           =   270
   End
   Begin VB.Image imgM 
      Height          =   270
      Left            =   4950
      Top             =   3075
      Width           =   270
   End
   Begin VB.Image imgD 
      Height          =   270
      Left            =   4485
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   1
      Left            =   7560
      Picture         =   "frmCrearPersonaje.frx":5DE88
      Top             =   6360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   0
      Left            =   6960
      Picture         =   "frmCrearPersonaje.frx":5E19A
      Top             =   6360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   1
      Left            =   8460
      Picture         =   "frmCrearPersonaje.frx":5E4AC
      Top             =   5925
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   0
      Left            =   6075
      Picture         =   "frmCrearPersonaje.frx":5E7BE
      Top             =   5925
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   8880
      Stretch         =   -1  'True
      Top             =   9120
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   5640
      Top             =   9120
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   4
      Left            =   4500
      TabIndex        =   9
      Top             =   4470
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   3
      Left            =   4500
      TabIndex        =   8
      Top             =   4125
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   5
      Left            =   4500
      TabIndex        =   7
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   2
      Left            =   4500
      TabIndex        =   6
      Top             =   3780
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   1
      Left            =   4500
      TabIndex        =   5
      Top             =   3450
      Width           =   225
   End
End
Attribute VB_Name = "frmCrearPersonaje"
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

Option Explicit

Public LastButtonPressed As clsGraphicalButton

Private picFullStar As Picture
Private picHalfStar As Picture
Private picGlowStar As Picture

Private Enum eHelp
    iePasswd
    ieTirarDados
    ieMail
    ieNombre
    ieConfirmPasswd
    ieAtributos
    ieD
    ieM
    ieF
    ieFuerza
    ieAgilidad
    ieInteligencia
    ieCarisma
    ieConstitucion
    ieEvasion
    ieMagia
    ieVida
    ieEscudos
    ieArmas
    ieArcos
    ieEspecialidad
    iePuebloOrigen
    ieRaza
    ieClase
    ieGenero
End Enum

Private vHelp(25) As String
Private vEspecialidades() As String

Private Type tModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type

Private Type tModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    DanoArmas As Double
    DanoProyectiles As Double
    Escudo As Double
    Magia As Double
    Vida As Double
    Hit As Double
End Type

Private ModRaza() As tModRaza
Private ModClase() As tModClase

Private NroRazas As Integer
Private NroClases As Integer

Private Cargando As Boolean

Private currentGrh As Long
Private Dir As E_Heading

Private Sub Form_Load()
    Cargando = True
    
    Me.Picture = LoadPicture(DirGraficos & "VentanaCrearPersonaje.jpg")
    Me.imgDados.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_TIRAR_DADOS").Item("TEXTO")
    Me.imgCrear.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_CREAR").Item("TEXTO")
    Me.imgVolver.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_VOLVER").Item("TEXTO")
    
    Me.imgEspecialidad.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_ESPECIALIDAD").Item("TEXTO")
    Me.imgNombre.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_NOMBRE").Item("TEXTO")
    Me.imgAtributos.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_ATRIBUTOS").Item("TEXTO")
    Me.imgFuerza.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_FUERZA").Item("TEXTO")
    Me.imgAgilidad.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_AGILIDAD").Item("TEXTO")
    Me.imgCarisma.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_CARISMA").Item("TEXTO")
    Me.imgConstitucion.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_CONSTITUCION").Item("TEXTO")
    Me.imgInteligencia.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_INTELIGENCIA").Item("TEXTO")
    Me.imgArcos.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_ARCOS").Item("TEXTO")
    Me.imgArmas.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_ARMAS").Item("TEXTO")
    Me.imgEscudos.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_ESCUDOS").Item("TEXTO")
    Me.imgEvasion.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_EVASION").Item("TEXTO")
    Me.imgMagia.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_MAGIA").Item("TEXTO")
    Me.imgVida.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_VIDA").Item("TEXTO")
    Me.imgPuebloOrigen.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_PUEBLO_ORIGEN").Item("TEXTO")
    Me.imgRaza.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_RAZA").Item("TEXTO")
    Me.imgClase.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_CLASE").Item("TEXTO")
    Me.imgGenero.Caption = JsonLanguage.Item("FRM_CREARPERSONAJE_GENERO").Item("TEXTO")
    
    Call LoadCharInfo
    Call CargarEspecialidades
    
    Call IniciarGraficos
    Call CargarCombos
    
    Call LoadHelp
    
    Call DrawImageInPicture(picPJ, Me.Picture, 0, 0, , , picPJ.Left, picPJ.Top)
    Dir = SOUTH
    
    Call TirarDados
    
    Cargando = False
    
    'UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = vbNullString
    UserHead = 0

End Sub

Private Sub CargarEspecialidades()

    ReDim vEspecialidades(1 To NroClases)
    
    vEspecialidades(eClass.Hunter) = JsonLanguage.Item("HABILIDADES").Item("OCULTARSE").Item("TEXTO")
    vEspecialidades(eClass.Thief) = JsonLanguage.Item("HABILIDADES").Item("ROBAR").Item("TEXTO") & JsonLanguage.Item("LETRA_Y").Item("TEXTO") & JsonLanguage.Item("HABILIDADES").Item("OCULTARSE").Item("TEXTO")
    vEspecialidades(eClass.Assasin) = JsonLanguage.Item("HABILIDADES").Item("APUNALAR").Item("TEXTO")
    vEspecialidades(eClass.Bandit) = JsonLanguage.Item("HABILIDADES").Item("COMBATE_CUERPO_A_CUERPO").Item("TEXTO")
    vEspecialidades(eClass.Druid) = JsonLanguage.Item("HABILIDADES").Item("DOMAR_ANIMALES").Item("TEXTO")
    vEspecialidades(eClass.Pirat) = JsonLanguage.Item("HABILIDADES").Item("NAVEGACION").Item("TEXTO")
    vEspecialidades(eClass.Worker) = JsonLanguage.Item("HABILIDADES").Item("MINERIA").Item("TEXTO") & "," _
                                    & JsonLanguage.Item("HABILIDADES").Item("CARPINTERIA").Item("TEXTO") & JsonLanguage.Item("LETRA_Y").Item("TEXTO") _
                                    & JsonLanguage.Item("HABILIDADES").Item("TALAR").Item("TEXTO")
End Sub
Private Sub IniciarGraficos()

    Dim GrhPath As String
    GrhPath = DirGraficos

    Set picFullStar = LoadPicture(GrhPath & "EstrellaSimple.jpg")
    Set picHalfStar = LoadPicture(GrhPath & "EstrellaMitad.jpg")
    Set picGlowStar = LoadPicture(GrhPath & "EstrellaBrillante.jpg")

End Sub

Private Sub CargarCombos()
Dim i As Integer
    Dim Lower_ciudades As Long, Lower_listaClases As Long, Lower_listaRazas As Long
    Dim Upper_ciudades As Long
    
    lstProfesion.Clear
    
    Lower_listaClases = LBound(ListaClases)
    
    For i = Lower_listaClases To NroClases
        lstProfesion.AddItem ListaClases(i)
    Next i
    
    lstHogar.Clear
    
    Lower_ciudades = LBound(Ciudades())
    Upper_ciudades = UBound(Ciudades())
    
    For i = Lower_ciudades To Upper_ciudades
        lstHogar.AddItem Ciudades(i)
    Next i
    
    lstRaza.Clear
    
    Lower_listaRazas = LBound(ListaRazas())
    
    For i = Lower_listaRazas To NroRazas
        lstRaza.AddItem ListaRazas(i)
    Next i
    
    lstProfesion.ListIndex = 1
End Sub

Function CheckData() As Boolean
    
    If LenB(txtNombre.Text) = 0 Then
        MsgBox JsonLanguage.Item("VALIDACION_NOMBRE_PJ").Item("TEXTO")
        txtNombre.SetFocus
        Exit Function
    End If

    If UserRaza = 0 Then
        MsgBox JsonLanguage.Item("VALIDACION_RAZA").Item("TEXTO")
        Exit Function
    End If
    
    If UserSexo = 0 Then
        MsgBox JsonLanguage.Item("VALIDACION_SEXO").Item("TEXTO")
        Exit Function
    End If
    
    If UserClase = 0 Then
        MsgBox JsonLanguage.Item("VALIDACION_CLASE").Item("TEXTO")
        Exit Function
    End If
    
    If UserHogar = 0 Then
        MsgBox JsonLanguage.Item("VALIDACION_HOGAR").Item("TEXTO")
        Exit Function
    End If
    
    If Len(AccountHash) = 0 Then
        MsgBox JsonLanguage.Item("VALIDACION_HASH").Item("TEXTO")
        Exit Function
    End If

    'Toqueteado x Salvito
    Dim i As Integer
    For i = 1 To NUMATRIBUTOS
        If Val(lblAtributos(i).Caption) = 0 Then
            MsgBox JsonLanguage.Item("VALIDACION_ATRIBUTOS").Item("TEXTO")
            Exit Function
        End If
    Next i
    
    If LenB(UserName) > 30 Then
        MsgBox JsonLanguage.Item("VALIDACION_BAD_NOMBRE_PJ").Item("TEXTO").Item(1)
        Exit Function
    End If
    
    CheckData = True

End Function

Private Sub TirarDados()
    Call WriteThrowDices
    Call FlushBuffer
End Sub

Private Sub DirPJ_Click(index As Integer)
    Select Case index
        Case 0
            Dir = CheckDir(Dir + 1)
        Case 1
            Dir = CheckDir(Dir - 1)
    End Select
    
    Call UpdateHeadSelection
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearLabel
End Sub

Private Sub HeadPJ_Click(index As Integer)
    Select Case index
        Case 0
            UserHead = CheckCabeza(UserHead + 1)
        Case 1
            UserHead = CheckCabeza(UserHead - 1)
    End Select
    
    Call UpdateHeadSelection
    
End Sub

Private Sub UpdateHeadSelection()
    Dim Head As Integer
    
    Head = UserHead
    Call DrawHead(Head, 2)
    
    Head = Head + 1
    Call DrawHead(CheckCabeza(Head), 3)
    
    Head = Head + 1
    Call DrawHead(CheckCabeza(Head), 4)
    
    Head = UserHead
    
    Head = Head - 1
    Call DrawHead(CheckCabeza(Head), 1)
    
    Head = Head - 1
    Call DrawHead(CheckCabeza(Head), 0)
End Sub

Private Sub ImgCrear_Click()

    Dim i As Integer
    
    UserName = txtNombre.Text
            
    If Right$(UserName, 1) = " " Then
        UserName = RTrim$(UserName)
        MsgBox JsonLanguage.Item("VALIDACION_BAD_NOMBRE_PJ").Item("TEXTO").Item(2)

    End If
    
    UserRaza = lstRaza.ListIndex + 1
    UserSexo = lstGenero.ListIndex + 1
    UserClase = lstProfesion.ListIndex + 1
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = Val(lblAtributos(i).Caption)
    Next i
         
    UserHogar = lstHogar.ListIndex + 1
    
    If Not CheckData Then Exit Sub
    
    EstadoLogin = E_MODO.CrearNuevoPj
    
    'Clear spell list
    frmMain.hlst.Clear
    
    #If UsarWrench = 1 Then
        If Not frmMain.Socket1.Connected Then
        
    #ElseIf UsarWrench = 2 Then
        If frmMain.Winsock1.State <> sckConnected Then
        
    #ElseIf UsarWrench = 3 Then
        If Not frmMain.Client.State = sckConnected Then
        
    #End If
            
            MsgBox JsonLanguage.Item("ERROR_CONN_LOST").Item("TEXTO")
            Unload Me
        
        Else
            Call Login

        End If
    
        bShowTutorial = True

    End Sub

Private Sub imgDados_Click()
    Call Audio.PlayWave(SND_DICE)
    Call TirarDados
End Sub

Private Sub imgEspecialidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEspecialidad)
End Sub

Private Sub imgNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub imgAtributos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAtributos)
End Sub

Private Sub imgD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieD)
End Sub

Private Sub imgM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieM)
End Sub

Private Sub imgF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieF)
End Sub

Private Sub imgFuerza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieFuerza)
End Sub

Private Sub imgAgilidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAgilidad)
End Sub

Private Sub imgInteligencia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieInteligencia)
End Sub

Private Sub imgCarisma_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieCarisma)
End Sub

Private Sub imgConstitucion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConstitucion)
End Sub

Private Sub imgArcos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArcos)
End Sub

Private Sub imgArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArmas)
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEscudos)
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEvasion)
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMagia)
End Sub

Private Sub imgVida_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieVida)
End Sub

Private Sub imgDados_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieTirarDados)
End Sub

Private Sub imgPuebloOrigen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.iePuebloOrigen)
End Sub

Private Sub imgRaza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieRaza)
End Sub

Private Sub imgClase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieClase)
End Sub

Private Sub imgGenero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieGenero)
End Sub

Private Sub imgVolver_Click()
    Call Audio.PlayMIDI("2.mid")
    
    bShowTutorial = False
    
    Unload Me
End Sub

Private Sub lstGenero_Click()
    UserSexo = lstGenero.ListIndex + 1
    Call DarCuerpoYCabeza
End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
    If lstProfesion.Text = "Trabajador" Then
        'Agarramos un numero aleatorio del 0 al 5 por que hay 5 imagenes de trabajador
        ImgProfesionDibujo.Picture = LoadPicture(App.path & "\graficos\" & lstProfesion.Text & (CInt(Rnd * 5)) & ".jpg")
    Else
        ImgProfesionDibujo.Picture = LoadPicture(App.path & "\graficos\" & lstProfesion.Text & ".jpg")
    End If
    
    UserClase = lstProfesion.ListIndex + 1
    
    Call UpdateStats
    Call UpdateEspecialidad(UserClase)
End Sub

Private Sub UpdateEspecialidad(ByVal eClase As eClass)
    lblEspecialidad.Caption = vEspecialidades(eClase)
End Sub

Private Sub lstRaza_Click()
    UserRaza = lstRaza.ListIndex + 1
    Call DarCuerpoYCabeza
    
    Call UpdateStats
End Sub

Private Sub picHead_Click(index As Integer)
    ' No se mueve si clickea al medio
    If index = 2 Then Exit Sub
    
    Dim Counter As Integer, Count_index As Long
    Dim Head As Integer
    
    Head = UserHead
    
    If index > 2 Then

        Count_index = index - 2
        For Counter = Count_index To 1 Step -1
            Head = CheckCabeza(Head + 1)
        Next Counter
    Else
        Count_index = 2 - index
        For Counter = Count_index To 1 Step -1
            Head = CheckCabeza(Head - 1)
        Next Counter
    End If
    
    UserHead = Head
    
    Call UpdateHeadSelection
    
End Sub

Private Sub tAnimacion_Timer()
On Error Resume Next
    Dim SR As RECT
    Dim DR As RECT
    Dim Grh As Long
    Static Frame As Byte
    
    If frmMain.Visible = False Then Exit Sub
    If currentGrh = 0 Then Exit Sub
    UserHead = CheckCabeza(UserHead)
    
    Frame = Frame + 1
    If Frame >= GrhData(currentGrh).NumFrames Then Frame = 1
    Call DrawImageInPicture(picPJ, Me.Picture, 0, 0, , , picPJ.Left, picPJ.Top)
    
    Grh = GrhData(currentGrh).Frames(Frame)
    
    With GrhData(Grh)
        SR.Left = .SX
        SR.Top = .SY
        SR.Right = SR.Left + .pixelWidth
        SR.Bottom = SR.Top + .pixelHeight
        
        DR.Left = (picPJ.Width - .pixelWidth) \ 2 - 2
        DR.Top = ((picPJ.Height - .pixelHeight) \ 2) + 5
        DR.Right = DR.Left + .pixelWidth
        DR.Bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picPJ.hdc, picTemp.hdc, DR, DR, vbBlack)
    End With
    
    Grh = HeadData(UserHead).Head(Dir).GrhIndex
    
    With GrhData(Grh)
        SR.Left = .SX
        SR.Top = .SY
        SR.Right = SR.Left + .pixelWidth
        SR.Bottom = SR.Top + .pixelHeight
        
        DR.Left = (picPJ.Width - .pixelWidth) \ 2 - 2
        DR.Top = DR.Bottom + BodyData(UserBody).HeadOffset.Y - .pixelHeight
        DR.Right = DR.Left + .pixelWidth
        DR.Bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picPJ.hdc, picTemp.hdc, DR, DR, vbBlack)
    End With
End Sub

Private Sub DrawHead(ByVal Head As Integer, ByVal PicIndex As Integer)

    Dim SR As RECT
    Dim DR As RECT
    Dim Grh As Long

    Call DrawImageInPicture(picHead(PicIndex), Me.Picture, 0, 0, , , picHead(PicIndex).Left, picHead(PicIndex).Top)
    
    Grh = HeadData(Head).Head(Dir).GrhIndex

    With GrhData(Grh)
        SR.Left = .SX
        SR.Top = .SY
        SR.Right = SR.Left + .pixelWidth
        SR.Bottom = SR.Top + .pixelHeight
        
        DR.Left = (picHead(0).Width - .pixelWidth) \ 2 + 1
        DR.Top = 5
        DR.Right = DR.Left + .pixelWidth
        DR.Bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picHead(PicIndex).hdc, picTemp.hdc, DR, DR, vbBlack)
    End With
    
End Sub

Private Sub txtConfirmPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)
End Sub

Private Sub txtNombre_Change()
    txtNombre.Text = LTrim$(txtNombre.Text)
End Sub

Private Sub DarCuerpoYCabeza()

    Dim bVisible As Boolean
    Dim PicIndex As Integer
    Dim LineIndex As Integer
    
    Select Case UserSexo
        Case eGenero.Hombre
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = eCabezas.HUMANO_H_PRIMER_CABEZA
                    UserBody = eCabezas.HUMANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = eCabezas.ELFO_H_PRIMER_CABEZA
                    UserBody = eCabezas.ELFO_H_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = eCabezas.DROW_H_PRIMER_CABEZA
                    UserBody = eCabezas.DROW_H_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = eCabezas.ENANO_H_PRIMER_CABEZA
                    UserBody = eCabezas.ENANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = eCabezas.GNOMO_H_PRIMER_CABEZA
                    UserBody = eCabezas.GNOMO_H_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case eGenero.Mujer
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = eCabezas.HUMANO_M_PRIMER_CABEZA
                    UserBody = eCabezas.HUMANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = eCabezas.ELFO_M_PRIMER_CABEZA
                    UserBody = eCabezas.ELFO_M_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = eCabezas.DROW_M_PRIMER_CABEZA
                    UserBody = eCabezas.DROW_M_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = eCabezas.ENANO_M_PRIMER_CABEZA
                    UserBody = eCabezas.ENANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = eCabezas.GNOMO_M_PRIMER_CABEZA
                    UserBody = eCabezas.GNOMO_M_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
        Case Else
            UserHead = 0
            UserBody = 0
    End Select
    
    bVisible = UserHead <> 0 And UserBody <> 0
    
    HeadPJ(0).Visible = bVisible
    HeadPJ(1).Visible = bVisible
    DirPJ(0).Visible = bVisible
    DirPJ(1).Visible = bVisible
    
    For PicIndex = 0 To 4
        picHead(PicIndex).Visible = bVisible
    Next PicIndex
    
    For LineIndex = 0 To 3
        Line1(LineIndex).Visible = bVisible
    Next LineIndex
    
    If bVisible Then Call UpdateHeadSelection
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    If currentGrh > 0 Then _
        tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)
End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer

Select Case UserSexo
    Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                If Head > eCabezas.HUMANO_H_ULTIMA_CABEZA Then
                    CheckCabeza = eCabezas.HUMANO_H_PRIMER_CABEZA + (Head - eCabezas.HUMANO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < eCabezas.HUMANO_H_PRIMER_CABEZA Then
                    CheckCabeza = eCabezas.HUMANO_H_ULTIMA_CABEZA - (eCabezas.HUMANO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Elfo
                If Head > eCabezas.ELFO_H_ULTIMA_CABEZA Then
                    CheckCabeza = eCabezas.ELFO_H_PRIMER_CABEZA + (Head - eCabezas.ELFO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < eCabezas.ELFO_H_PRIMER_CABEZA Then
                    CheckCabeza = eCabezas.ELFO_H_ULTIMA_CABEZA - (eCabezas.ELFO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.ElfoOscuro
                If Head > eCabezas.DROW_H_ULTIMA_CABEZA Then
                    CheckCabeza = eCabezas.DROW_H_PRIMER_CABEZA + (Head - eCabezas.DROW_H_ULTIMA_CABEZA) - 1
                ElseIf Head < eCabezas.DROW_H_PRIMER_CABEZA Then
                    CheckCabeza = eCabezas.DROW_H_ULTIMA_CABEZA - (eCabezas.DROW_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Enano
                If Head > eCabezas.ENANO_H_ULTIMA_CABEZA Then
                    CheckCabeza = eCabezas.ENANO_H_PRIMER_CABEZA + (Head - eCabezas.ENANO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < eCabezas.ENANO_H_PRIMER_CABEZA Then
                    CheckCabeza = eCabezas.ENANO_H_ULTIMA_CABEZA - (eCabezas.ENANO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Gnomo
                If Head > eCabezas.GNOMO_H_ULTIMA_CABEZA Then
                    CheckCabeza = eCabezas.GNOMO_H_PRIMER_CABEZA + (Head - eCabezas.GNOMO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < eCabezas.GNOMO_H_PRIMER_CABEZA Then
                    CheckCabeza = eCabezas.GNOMO_H_ULTIMA_CABEZA - (eCabezas.GNOMO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case Else
                UserRaza = lstRaza.ListIndex + 1
                CheckCabeza = CheckCabeza(Head)
        End Select
        
    Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                If Head > eCabezas.HUMANO_M_ULTIMA_CABEZA Then
                    CheckCabeza = eCabezas.HUMANO_M_PRIMER_CABEZA + (Head - eCabezas.HUMANO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < eCabezas.HUMANO_M_PRIMER_CABEZA Then
                    CheckCabeza = eCabezas.HUMANO_M_ULTIMA_CABEZA - (eCabezas.HUMANO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Elfo
                If Head > eCabezas.ELFO_M_ULTIMA_CABEZA Then
                    CheckCabeza = eCabezas.ELFO_M_PRIMER_CABEZA + (Head - eCabezas.ELFO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < eCabezas.ELFO_M_PRIMER_CABEZA Then
                    CheckCabeza = eCabezas.ELFO_M_ULTIMA_CABEZA - (eCabezas.ELFO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.ElfoOscuro
                If Head > eCabezas.DROW_M_ULTIMA_CABEZA Then
                    CheckCabeza = eCabezas.DROW_M_PRIMER_CABEZA + (Head - eCabezas.DROW_M_ULTIMA_CABEZA) - 1
                ElseIf Head < eCabezas.DROW_M_PRIMER_CABEZA Then
                    CheckCabeza = eCabezas.DROW_M_ULTIMA_CABEZA - (eCabezas.DROW_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Enano
                If Head > eCabezas.ENANO_M_ULTIMA_CABEZA Then
                    CheckCabeza = eCabezas.ENANO_M_PRIMER_CABEZA + (Head - eCabezas.ENANO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < eCabezas.ENANO_M_PRIMER_CABEZA Then
                    CheckCabeza = eCabezas.ENANO_M_ULTIMA_CABEZA - (eCabezas.ENANO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Gnomo
                If Head > eCabezas.GNOMO_M_ULTIMA_CABEZA Then
                    CheckCabeza = eCabezas.GNOMO_M_PRIMER_CABEZA + (Head - eCabezas.GNOMO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < eCabezas.GNOMO_M_PRIMER_CABEZA Then
                    CheckCabeza = eCabezas.GNOMO_M_ULTIMA_CABEZA - (eCabezas.GNOMO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case Else
                UserRaza = lstRaza.ListIndex + 1
                CheckCabeza = CheckCabeza(Head)
        End Select
    Case Else
        UserSexo = lstGenero.ListIndex + 1
        CheckCabeza = CheckCabeza(Head)
End Select
End Function

Private Function CheckDir(ByRef Dir As E_Heading) As E_Heading

    If Dir > E_Heading.WEST Then Dir = E_Heading.NORTH
    If Dir < E_Heading.NORTH Then Dir = E_Heading.WEST
    
    CheckDir = Dir
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    If currentGrh > 0 Then _
        tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)

End Function

Private Sub LoadHelp()
    vHelp(eHelp.ieTirarDados) = JsonLanguage.Item("AYUDA_CREARPJ_DADOS").Item("TEXTO")
    vHelp(eHelp.ieNombre) = JsonLanguage.Item("AYUDA_CREARPJ_NOMBREPJ").Item("TEXTO")
    vHelp(eHelp.ieAtributos) = JsonLanguage.Item("AYUDA_CREARPJ_ATRIBUTOS").Item("TEXTO")
    vHelp(eHelp.ieD) = JsonLanguage.Item("AYUDA_CREARPJ_IED").Item("TEXTO")
    vHelp(eHelp.ieM) = JsonLanguage.Item("AYUDA_CREARPJ_IEM").Item("TEXTO")
    vHelp(eHelp.ieF) = JsonLanguage.Item("AYUDA_CREARPJ_IEF").Item("TEXTO")
    vHelp(eHelp.ieFuerza) = JsonLanguage.Item("AYUDA_CREARPJ_FUERZA").Item("TEXTO")
    vHelp(eHelp.ieAgilidad) = JsonLanguage.Item("AYUDA_CREARPJ_AGILIDAD").Item("TEXTO")
    vHelp(eHelp.ieInteligencia) = JsonLanguage.Item("AYUDA_CREARPJ_INTELIGENCIA").Item("TEXTO")
    vHelp(eHelp.ieCarisma) = JsonLanguage.Item("AYUDA_CREARPJ_CARISMA").Item("TEXTO")
    vHelp(eHelp.ieConstitucion) = JsonLanguage.Item("AYUDA_CREARPJ_CONSTITUCION").Item("TEXTO")
    vHelp(eHelp.ieEvasion) = JsonLanguage.Item("AYUDA_CREARPJ_EVASION").Item("TEXTO")
    vHelp(eHelp.ieMagia) = JsonLanguage.Item("AYUDA_CREARPJ_MAGIA").Item("TEXTO")
    vHelp(eHelp.ieVida) = JsonLanguage.Item("AYUDA_CREARPJ_VIDA").Item("TEXTO")
    vHelp(eHelp.ieEscudos) = JsonLanguage.Item("AYUDA_CREARPJ_ESCUDOS").Item("TEXTO")
    vHelp(eHelp.ieArmas) = JsonLanguage.Item("AYUDA_CREARPJ_ARMAS").Item("TEXTO")
    vHelp(eHelp.ieArcos) = JsonLanguage.Item("AYUDA_CREARPJ_ARCOS").Item("TEXTO")
    vHelp(eHelp.iePuebloOrigen) = JsonLanguage.Item("AYUDA_CREARPJ_HOGAR").Item("TEXTO")
    vHelp(eHelp.ieRaza) = JsonLanguage.Item("AYUDA_CREARPJ_RAZA").Item("TEXTO")
    vHelp(eHelp.ieClase) = JsonLanguage.Item("AYUDA_CREARPJ_CLASE").Item("TEXTO")
    vHelp(eHelp.ieGenero) = JsonLanguage.Item("AYUDA_CREARPJ_GENERO").Item("TEXTO")
End Sub

Private Sub ClearLabel()
    lblHelp = vbNullString
End Sub

Private Sub txtNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub txtPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.iePasswd)
End Sub

Public Sub UpdateStats()
    Call UpdateRazaMod
    Call UpdateStars
End Sub

Private Sub UpdateRazaMod()
    Dim SelRaza As Integer
    Dim i As Integer
    
    
    If lstRaza.ListIndex > -1 Then
    
        SelRaza = lstRaza.ListIndex + 1
        
        With ModRaza(SelRaza)
            lblModRaza(eAtributos.Fuerza).Caption = IIf(.Fuerza >= 0, "+", vbNullString) & .Fuerza
            lblModRaza(eAtributos.Agilidad).Caption = IIf(.Agilidad >= 0, "+", vbNullString) & .Agilidad
            lblModRaza(eAtributos.Inteligencia).Caption = IIf(.Inteligencia >= 0, "+", vbNullString) & .Inteligencia
            lblModRaza(eAtributos.Carisma).Caption = IIf(.Carisma >= 0, "+", "") & .Carisma
            lblModRaza(eAtributos.Constitucion).Caption = IIf(.Constitucion >= 0, "+", vbNullString) & .Constitucion
        End With
    End If
    
    ' Atributo total
    For i = 1 To NUMATRIBUTES
        lblAtributoFinal(i).Caption = Val(lblAtributos(i).Caption) + Val(lblModRaza(i))
    Next i
    
End Sub

Private Sub UpdateStars()
    Dim NumStars As Double
    
    If UserClase = 0 Then Exit Sub
    
    ' Estrellas de evasion
    NumStars = (2.454 + 0.073 * Val(lblAtributoFinal(eAtributos.Agilidad).Caption)) * ModClase(UserClase).Evasion
    Call SetStars(imgEvasionStar, NumStars * 2)
    
    ' Estrellas de magia
    NumStars = ModClase(UserClase).Magia * Val(lblAtributoFinal(eAtributos.Inteligencia).Caption) * 0.085
    Call SetStars(imgMagiaStar, NumStars * 2)
    
    ' Estrellas de vida
    NumStars = 0.24 + (Val(lblAtributoFinal(eAtributos.Constitucion).Caption) * 0.5 - ModClase(UserClase).Vida) * 0.475
    Call SetStars(imgVidaStar, NumStars * 2)
    
    ' Estrellas de escudo
    NumStars = 4 * ModClase(UserClase).Escudo
    Call SetStars(imgEscudosStar, NumStars * 2)
    
    ' Estrellas de armas
    NumStars = (0.509 + 0.01185 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * ModClase(UserClase).Hit * _
                ModClase(UserClase).DanoArmas + 0.119 * ModClase(UserClase).AtaqueArmas * _
                Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArmasStar, NumStars * 2)
    
    ' Estrellas de arcos
    NumStars = (0.4915 + 0.01265 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * _
                ModClase(UserClase).DanoProyectiles * ModClase(UserClase).Hit + 0.119 * ModClase(UserClase).AtaqueProyectiles * _
                Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArcoStar, NumStars * 2)
End Sub

Private Sub SetStars(ByRef ImgContainer As Object, ByVal NumStars As Integer)
    Dim FullStars As Integer
    Dim HasHalfStar As Boolean
    Dim index As Integer
    Dim Counter As Integer

    If NumStars > 0 Then
        
        If NumStars > 10 Then NumStars = 10
        
        FullStars = Int(NumStars / 2)
        
        ' Tienen brillo extra si estan todas
        If FullStars = 5 Then
            For index = 1 To FullStars
                ImgContainer(index).Picture = picGlowStar
            Next index
        Else
            ' Numero impar? Entonces hay que poner "media estrella"
            If (NumStars Mod 2) > 0 Then HasHalfStar = True
            
            ' Muestro las estrellas enteras
            If FullStars > 0 Then
                For index = 1 To FullStars
                    ImgContainer(index).Picture = picFullStar
                Next index
                
                Counter = FullStars
            End If
            
            ' Muestro la mitad de la estrella (si tiene)
            If HasHalfStar Then
                Counter = Counter + 1
                
                ImgContainer(Counter).Picture = picHalfStar
            End If
            
            ' Si estan completos los espacios, no borro nada
            If Counter <> 5 Then
                
                'pre-calculo el index para mejorar el rendimiento
                Dim Count_index As Long
                    Count_index = Counter + 1
                
                ' Limpio las que queden vacias
                For index = Count_index To 5
                    Set ImgContainer(index).Picture = Nothing
                Next index
            End If
            
        End If
    Else
        ' Limpio todo
        For index = 1 To 5
            Set ImgContainer(index).Picture = Nothing
        Next index
    End If

End Sub

Private Sub LoadCharInfo()
    Dim SearchVar As String
    Dim i As Integer
    
    NroRazas = UBound(ListaRazas())
    NroClases = UBound(ListaClases())

    ReDim ModRaza(1 To NroRazas)
    ReDim ModClase(1 To NroClases)
    
    'Modificadores de Clase
    For i = 1 To NroClases
        With ModClase(i)
            SearchVar = ListaClases(i)
            
            .Evasion = Val(GetVar(IniPath & "CharInfo.dat", "MODEVASION", SearchVar))
            .AtaqueArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODATAQUEARMAS", SearchVar))
            .AtaqueProyectiles = Val(GetVar(IniPath & "CharInfo.dat", "MODATAQUEPROYECTILES", SearchVar))
            .DanoArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODDANOARMAS", SearchVar))
            .DanoProyectiles = Val(GetVar(IniPath & "CharInfo.dat", "MODDANOPROYECTILES", SearchVar))
            .Escudo = Val(GetVar(IniPath & "CharInfo.dat", "MODESCUDO", SearchVar))
            .Hit = Val(GetVar(IniPath & "CharInfo.dat", "HIT", SearchVar))
            .Magia = Val(GetVar(IniPath & "CharInfo.dat", "MODMAGIA", SearchVar))
            .Vida = Val(GetVar(IniPath & "CharInfo.dat", "MODVIDA", SearchVar))
        End With
    Next i
    
    'Modificadores de Raza
    For i = 1 To NroRazas
        With ModRaza(i)
            SearchVar = Replace(ListaRazas(i), " ", "")
        
            .Fuerza = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Inteligencia"))
            .Carisma = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Carisma"))
            .Constitucion = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Constitucion"))
        End With
    Next i

End Sub
