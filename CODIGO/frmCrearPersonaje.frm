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
   KeyPreview      =   -1  'True
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
      Left            =   6840
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   32
      Top             =   6840
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
      Left            =   6840
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   31
      Top             =   6840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox lstAlienacion 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
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
      Left            =   6120
      List            =   "frmCrearPersonaje.frx":53C97
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.TextBox txtMail 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      TabIndex        =   3
      Text            =   "Deshabilitado"
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox txtConfirmPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   6120
      TabIndex        =   2
      Text            =   "Deshabilitado"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   3480
      TabIndex        =   1
      Text            =   "Deshabilitado"
      Top             =   1800
      Width           =   2415
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
      ItemData        =   "frmCrearPersonaje.frx":53CAA
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":53CAC
      Style           =   2  'Dropdown List
      TabIndex        =   6
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
      ItemData        =   "frmCrearPersonaje.frx":53CAE
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":53CB8
      Style           =   2  'Dropdown List
      TabIndex        =   7
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
      ItemData        =   "frmCrearPersonaje.frx":53CCB
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":53CCD
      Style           =   2  'Dropdown List
      TabIndex        =   5
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
      ItemData        =   "frmCrearPersonaje.frx":53CCF
      Left            =   6060
      List            =   "frmCrearPersonaje.frx":53CD1
      Style           =   2  'Dropdown List
      TabIndex        =   4
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
      TabIndex        =   26
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
      TabIndex        =   27
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
      TabIndex        =   28
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
      TabIndex        =   29
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
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
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
      TabIndex        =   30
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   3450
      Width           =   225
   End
   Begin VB.Image imgAtributos 
      Height          =   270
      Left            =   3960
      Top             =   2745
      Width           =   975
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
      TabIndex        =   14
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Image imgVolver 
      Height          =   450
      Left            =   1335
      Top             =   8190
      Width           =   1290
   End
   Begin VB.Image imgCrear 
      Height          =   435
      Left            =   9090
      Top             =   8190
      Width           =   2610
   End
   Begin VB.Image imgalineacion 
      Height          =   240
      Left            =   6855
      Top             =   4830
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image imgGenero 
      Height          =   240
      Left            =   6960
      Top             =   4335
      Width           =   705
   End
   Begin VB.Image imgClase 
      Height          =   240
      Left            =   7020
      Top             =   3795
      Width           =   555
   End
   Begin VB.Image imgRaza 
      Height          =   255
      Left            =   7035
      Top             =   3270
      Width           =   570
   End
   Begin VB.Image imgPuebloOrigen 
      Height          =   225
      Left            =   6600
      Top             =   2760
      Width           =   1425
   End
   Begin VB.Image imgEspecialidad 
      Height          =   240
      Left            =   3330
      Top             =   7410
      Width           =   1065
   End
   Begin VB.Image imgArcos 
      Height          =   225
      Left            =   3345
      Top             =   7140
      Width           =   555
   End
   Begin VB.Image imgArmas 
      Height          =   240
      Left            =   3330
      Top             =   6840
      Width           =   615
   End
   Begin VB.Image imgEscudos 
      Height          =   255
      Left            =   3315
      Top             =   6540
      Width           =   735
   End
   Begin VB.Image imgVida 
      Height          =   225
      Left            =   3330
      Top             =   6270
      Width           =   465
   End
   Begin VB.Image imgMagia 
      Height          =   255
      Left            =   3285
      Top             =   5955
      Width           =   660
   End
   Begin VB.Image imgEvasion 
      Height          =   255
      Left            =   3285
      Top             =   5670
      Width           =   735
   End
   Begin VB.Image imgConstitucion 
      Height          =   255
      Left            =   3285
      Top             =   4785
      Width           =   1080
   End
   Begin VB.Image imgCarisma 
      Height          =   240
      Left            =   3435
      Top             =   4440
      Width           =   765
   End
   Begin VB.Image imgInteligencia 
      Height          =   240
      Left            =   3330
      Top             =   4110
      Width           =   1005
   End
   Begin VB.Image imgAgilidad 
      Height          =   240
      Left            =   3420
      Top             =   3765
      Width           =   735
   End
   Begin VB.Image imgFuerza 
      Height          =   240
      Left            =   3450
      Top             =   3420
      Width           =   675
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
   Begin VB.Image imgConfirmPasswd 
      Height          =   255
      Left            =   6585
      Top             =   1545
      Width           =   1440
   End
   Begin VB.Image imgPasswd 
      Height          =   255
      Left            =   4350
      Top             =   1545
      Width           =   930
   End
   Begin VB.Image imgNombre 
      Height          =   240
      Left            =   5205
      Top             =   1065
      Width           =   1635
   End
   Begin VB.Image imgMail 
      Height          =   240
      Left            =   5310
      Top             =   2055
      Width           =   1395
   End
   Begin VB.Image imgTirarDados 
      Height          =   765
      Left            =   1380
      Top             =   3105
      Width           =   1200
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   1
      Left            =   7560
      Picture         =   "frmCrearPersonaje.frx":53CD3
      Top             =   6360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   0
      Left            =   6960
      Picture         =   "frmCrearPersonaje.frx":53FE5
      Top             =   6360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   1
      Left            =   8460
      Picture         =   "frmCrearPersonaje.frx":542F7
      Top             =   5925
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   0
      Left            =   6075
      Picture         =   "frmCrearPersonaje.frx":54609
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
   Begin VB.Image imgDados 
      Height          =   885
      Left            =   195
      MouseIcon       =   "frmCrearPersonaje.frx":5491B
      MousePointer    =   99  'Custom
      Top             =   2775
      Width           =   900
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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

Private cBotonPasswd        As clsGraphicalButton
Private cBotonTirarDados    As clsGraphicalButton
Private cBotonMail          As clsGraphicalButton
Private cBotonNombre        As clsGraphicalButton
Private cBotonConfirmPasswd As clsGraphicalButton
Private cBotonAtributos     As clsGraphicalButton
Private cBotonD             As clsGraphicalButton
Private cBotonM             As clsGraphicalButton
Private cBotonF             As clsGraphicalButton
Private cBotonFuerza        As clsGraphicalButton
Private cBotonAgilidad      As clsGraphicalButton
Private cBotonInteligencia  As clsGraphicalButton
Private cBotonCarisma       As clsGraphicalButton
Private cBotonConstitucion  As clsGraphicalButton
Private cBotonEvasion       As clsGraphicalButton
Private cBotonMagia         As clsGraphicalButton
Private cBotonVida          As clsGraphicalButton
Private cBotonEscudos       As clsGraphicalButton
Private cBotonArmas         As clsGraphicalButton
Private cBotonArcos         As clsGraphicalButton
Private cBotonEspecialidad  As clsGraphicalButton
Private cBotonPuebloOrigen  As clsGraphicalButton
Private cBotonRaza          As clsGraphicalButton
Private cBotonClase         As clsGraphicalButton
Private cBotonGenero        As clsGraphicalButton
Private cBotonAlineacion    As clsGraphicalButton
Private cBotonVolver        As clsGraphicalButton
Private cBotonCrear         As clsGraphicalButton

Public LastButtonPressed    As clsGraphicalButton

Private picFullStar         As Picture
Private picHalfStar         As Picture
Private picGlowStar         As Picture

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
    ieAlineacion

End Enum

Private vHelp(25)         As String
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
    Da�oArmas As Double
    Da�oProyectiles As Double
    Escudo As Double
    Magia As Double
    Vida As Double
    Hit As Double

End Type

Private ModRaza()  As tModRaza
Private ModClase() As tModClase

Private NroRazas   As Integer
Private NroClases  As Integer

Private Cargando   As Boolean

Private currentGrh As Long
Private Dir        As E_Heading

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Me.Picture = LoadPicture(DirGraficos & "VentanaCrearPersonaje.jpg")
    
    Cargando = True
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

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub CargarEspecialidades()
    
    On Error GoTo CargarEspecialidades_Err
    

    ReDim vEspecialidades(1 To NroClases)
    
<<<<<<< HEAD
    vEspecialidades(eClass.Hunter) = "Ocultarse"
    vEspecialidades(eClass.Thief) = "Robar y Ocultarse"
    vEspecialidades(eClass.Assasin) = "Apu�alar"
    vEspecialidades(eClass.Bandit) = "Combate Sin Armas"
    vEspecialidades(eClass.Druid) = "Domar"
    vEspecialidades(eClass.Pirat) = "Navegar"
    vEspecialidades(eClass.Worker) = "Extracci�n y Construcci�n"

    
    Exit Sub

CargarEspecialidades_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "CargarEspecialidades"
    End If
Resume Next
    
=======
    vEspecialidades(eClass.Hunter) = JsonLanguage.Item("HABILIDADES").Item("OCULTARSE").Item("TEXTO")
    vEspecialidades(eClass.Thief) = JsonLanguage.Item("HABILIDADES").Item("ROBAR").Item("TEXTO") & JsonLanguage.Item("LETRA_Y").Item("TEXTO") & JsonLanguage.Item("HABILIDADES").Item("OCULTARSE").Item("TEXTO")
    vEspecialidades(eClass.Assasin) = JsonLanguage.Item("HABILIDADES").Item("APUNALAR").Item("TEXTO")
    vEspecialidades(eClass.Bandit) = JsonLanguage.Item("HABILIDADES").Item("COMBATE_CUERPO_A_CUERPO").Item("TEXTO")
    vEspecialidades(eClass.Druid) = JsonLanguage.Item("HABILIDADES").Item("DOMAR_ANIMALES").Item("TEXTO")
    vEspecialidades(eClass.Pirat) = JsonLanguage.Item("HABILIDADES").Item("NAVEGACION").Item("TEXTO")
    vEspecialidades(eClass.Worker) = JsonLanguage.Item("HABILIDADES").Item("MINERIA").Item("TEXTO") & "," _
                                    & JsonLanguage.Item("HABILIDADES").Item("CARPINTERIA").Item("TEXTO") & JsonLanguage.Item("LETRA_Y").Item("TEXTO") _
                                    & JsonLanguage.Item("HABILIDADES").Item("TALAR").Item("TEXTO")
>>>>>>> origin/master
End Sub
Private Sub IniciarGraficos()
    
    On Error GoTo IniciarGraficos_Err
    

    Dim GrhPath As String
    GrhPath = DirGraficos
    
    Set cBotonPasswd = New clsGraphicalButton
    Set cBotonTirarDados = New clsGraphicalButton
    Set cBotonMail = New clsGraphicalButton
    Set cBotonNombre = New clsGraphicalButton
    Set cBotonConfirmPasswd = New clsGraphicalButton
    Set cBotonAtributos = New clsGraphicalButton
    Set cBotonD = New clsGraphicalButton
    Set cBotonM = New clsGraphicalButton
    Set cBotonF = New clsGraphicalButton
    Set cBotonFuerza = New clsGraphicalButton
    Set cBotonAgilidad = New clsGraphicalButton
    Set cBotonInteligencia = New clsGraphicalButton
    Set cBotonCarisma = New clsGraphicalButton
    Set cBotonConstitucion = New clsGraphicalButton
    Set cBotonEvasion = New clsGraphicalButton
    Set cBotonMagia = New clsGraphicalButton
    Set cBotonVida = New clsGraphicalButton
    Set cBotonEscudos = New clsGraphicalButton
    Set cBotonArmas = New clsGraphicalButton
    Set cBotonArcos = New clsGraphicalButton
    Set cBotonEspecialidad = New clsGraphicalButton
    Set cBotonPuebloOrigen = New clsGraphicalButton
    Set cBotonRaza = New clsGraphicalButton
    Set cBotonClase = New clsGraphicalButton
    Set cBotonGenero = New clsGraphicalButton
    Set cBotonAlineacion = New clsGraphicalButton
    Set cBotonVolver = New clsGraphicalButton
    Set cBotonCrear = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonPasswd.Initialize(imgPasswd, "", GrhPath & "BotonPassword.jpg", GrhPath & "BotonPassword.jpg", Me, , , False, False)
                                    
    Call cBotonTirarDados.Initialize(imgTirarDados, "", GrhPath & "BotonTirarDados.jpg", GrhPath & "BotonTirarDados.jpg", Me, , , False, False)
                                    
    Call cBotonMail.Initialize(imgMail, "", GrhPath & "BotonMailPj.jpg", GrhPath & "BotonMailPj.jpg", Me, , , False, False)
                                    
    Call cBotonNombre.Initialize(imgNombre, "", GrhPath & "BotonNombrePJ.jpg", GrhPath & "BotonNombrePJ.jpg", Me, , , False, False)
                                    
    Call cBotonConfirmPasswd.Initialize(imgConfirmPasswd, "", GrhPath & "BotonRepetirPassword.jpg", GrhPath & "BotonRepetirPassword.jpg", Me, , , False, False)
                                    
    Call cBotonAtributos.Initialize(imgAtributos, "", GrhPath & "BotonAtributos.jpg", GrhPath & "BotonAtributos.jpg", Me, , , False, False)
                                    
    Call cBotonD.Initialize(imgD, "", GrhPath & "BotonD.jpg", GrhPath & "BotonD.jpg", Me, , , False, False)
                                    
    Call cBotonM.Initialize(imgM, "", GrhPath & "BotonM.jpg", GrhPath & "BotonM.jpg", Me, , , False, False)
                                    
    Call cBotonF.Initialize(imgF, "", GrhPath & "BotonF.jpg", GrhPath & "BotonF.jpg", Me, , , False, False)
                                    
    Call cBotonFuerza.Initialize(imgFuerza, "", GrhPath & "BotonFuerza.jpg", GrhPath & "BotonFuerza.jpg", Me, , , False, False)
                                    
    Call cBotonAgilidad.Initialize(imgAgilidad, "", GrhPath & "BotonAgilidad.jpg", GrhPath & "BotonAgilidad.jpg", Me, , , False, False)
                                    
    Call cBotonInteligencia.Initialize(imgInteligencia, "", GrhPath & "BotonInteligencia.jpg", GrhPath & "BotonInteligencia.jpg", Me, , , False, False)
                                    
    Call cBotonCarisma.Initialize(imgCarisma, "", GrhPath & "BotonCarisma.jpg", GrhPath & "BotonCarisma.jpg", Me, , , False, False)
                                    
    Call cBotonConstitucion.Initialize(imgConstitucion, "", GrhPath & "BotonConstitucion.jpg", GrhPath & "BotonConstitucion.jpg", Me, , , False, False)
                                    
    Call cBotonEvasion.Initialize(imgEvasion, "", GrhPath & "BotonEvasion.jpg", GrhPath & "BotonEvasion.jpg", Me, , , False, False)
                                    
    Call cBotonMagia.Initialize(imgMagia, "", GrhPath & "BotonMagia.jpg", GrhPath & "BotonMagia.jpg", Me, , , False, False)
                                    
    Call cBotonVida.Initialize(imgVida, "", GrhPath & "BotonVida.jpg", GrhPath & "BotonVida.jpg", Me, , , False, False)
                                    
    Call cBotonEscudos.Initialize(imgEscudos, "", GrhPath & "BotonEscudos.jpg", GrhPath & "BotonEscudos.jpg", Me, , , False, False)
                                    
    Call cBotonArmas.Initialize(imgArmas, "", GrhPath & "BotonArmas.jpg", GrhPath & "BotonArmas.jpg", Me, , , False, False)
                                    
    Call cBotonArcos.Initialize(imgArcos, "", GrhPath & "BotonArcos.jpg", GrhPath & "BotonArcos.jpg", Me, , , False, False)
                                    
    Call cBotonEspecialidad.Initialize(imgEspecialidad, "", GrhPath & "BotonEspecialidad.jpg", GrhPath & "BotonEspecialidad.jpg", Me, , , False, False)
                                    
    Call cBotonPuebloOrigen.Initialize(imgPuebloOrigen, "", GrhPath & "BotonPuebloOrigen.jpg", GrhPath & "BotonPuebloOrigen.jpg", Me, , , False, False)
                                    
    Call cBotonRaza.Initialize(imgRaza, "", GrhPath & "BotonRaza.jpg", GrhPath & "BotonRaza.jpg", Me, , , False, False)
                                    
    Call cBotonClase.Initialize(imgClase, "", GrhPath & "BotonClase.jpg", GrhPath & "BotonClase.jpg", Me, , , False, False)
                                    
    Call cBotonGenero.Initialize(imgGenero, "", GrhPath & "BotonGenero.jpg", GrhPath & "BotonGenero.jpg", Me, , , False, False)
                                    
    Call cBotonAlineacion.Initialize(imgalineacion, "", GrhPath & "BotonAlineacion.jpg", GrhPath & "BotonAlineacion.jpg", Me, , , False, False)
                                    
    Call cBotonVolver.Initialize(imgVolver, "", GrhPath & "BotonVolverRollover.jpg", GrhPath & "BotonVolverClick.jpg", Me)
                                    
    Call cBotonCrear.Initialize(imgCrear, "", GrhPath & "BotonCrearPersonajeRollover.jpg", GrhPath & "BotonCrearPersonajeClick.jpg", Me)

    Set picFullStar = LoadPicture(GrhPath & "EstrellaSimple.jpg")
    Set picHalfStar = LoadPicture(GrhPath & "EstrellaMitad.jpg")
    Set picGlowStar = LoadPicture(GrhPath & "EstrellaBrillante.jpg")

    
    Exit Sub

IniciarGraficos_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "IniciarGraficos"
    End If
Resume Next
    
End Sub

Private Sub CargarCombos()
    
    On Error GoTo CargarCombos_Err
    
    Dim i              As Integer
    Dim Lower_ciudades As Long
    Dim Upper_ciudades As Long
    Dim Lower_clases   As Long
    Dim Lower_razas    As Long
    
    lstProfesion.Clear
    
    Lower_ciudades = LBound(Ciudades())
    Upper_ciudades = UBound(Ciudades())
    
    Lower_clases = LBound(ListaClases)
    
    Lower_razas = LBound(ListaRazas())
    
    For i = Lower_clases To NroClases
        lstProfesion.AddItem ListaClases(i)
    Next i
    
    lstHogar.Clear
    
    For i = Lower_ciudades To Upper_ciudades
        lstHogar.AddItem Ciudades(i)
    Next i
    
    lstRaza.Clear
    
    For i = Lower_razas To NroRazas
        lstRaza.AddItem ListaRazas(i)
    Next i
    
    lstProfesion.ListIndex = 1
    
    Exit Sub

CargarCombos_Err:

    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "CargarCombos"

    End If

    Resume Next
    
End Sub

Function CheckData() As Boolean
    
    On Error GoTo CheckData_Err
    
    
    If LenB(txtNombre.Text) = 0 Then
        MsgBox JsonLanguage.Item("VALIDACION_NOMBRE_PJ").Item("TEXTO")
        txtNombre.SetFocus
        Exit Function

    End If
<<<<<<< HEAD
    
    If LenB(txtPasswd.Text) = 0 Then
        MsgBox "Ingresa una contrase�a."
        txtPasswd.SetFocus
        Exit Function

    End If
=======
>>>>>>> origin/master

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

    
    Exit Function

CheckData_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "CheckData"
    End If
Resume Next
    
End Function

Private Sub TirarDados()
    
    On Error GoTo TirarDados_Err
    
    Call WriteThrowDices
    Call FlushBuffer

    
    Exit Sub

TirarDados_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "TirarDados"
    End If
Resume Next
    
End Sub

Private Sub DirPJ_Click(Index As Integer)
    
    On Error GoTo DirPJ_Click_Err
    

    Select Case Index

        Case 0
            Dir = CheckDir(Dir + 1)

        Case 1
            Dir = CheckDir(Dir - 1)

    End Select
    
    Call UpdateHeadSelection

    
    Exit Sub

DirPJ_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "DirPJ_Click"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    ClearLabel

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub HeadPJ_Click(Index As Integer)
    
    On Error GoTo HeadPJ_Click_Err
    

    Select Case Index

        Case 0
            UserHead = CheckCabeza(UserHead + 1)

        Case 1
            UserHead = CheckCabeza(UserHead - 1)

    End Select
    
    Call UpdateHeadSelection
    
    
    Exit Sub

HeadPJ_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "HeadPJ_Click"
    End If
Resume Next
    
End Sub

Private Sub UpdateHeadSelection()
    
    On Error GoTo UpdateHeadSelection_Err
    
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

    
    Exit Sub

UpdateHeadSelection_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "UpdateHeadSelection"
    End If
Resume Next
    
End Sub

Private Sub ImgCrear_Click()
    
    On Error GoTo ImgCrear_Click_Err
    

    Dim i         As Integer
    Dim CharAscii As Byte
    
    UserName = txtNombre.Text
            
    If Right$(UserName, 1) = " " Then
        UserName = RTrim$(UserName)
<<<<<<< HEAD
        MsgBox "Nombre invalido, se han removido los espacios al final del nombre"

=======
        MsgBox JsonLanguage.Item("VALIDACION_BAD_NOMBRE_PJ").Item("TEXTO").Item(2)
>>>>>>> origin/master
    End If
    
    UserRaza = lstRaza.ListIndex + 1
    UserSexo = lstGenero.ListIndex + 1
    UserClase = lstProfesion.ListIndex + 1
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = Val(lblAtributos(i).Caption)
    Next i
         
    UserHogar = lstHogar.ListIndex + 1
    
    If Not CheckData Then Exit Sub
    
<<<<<<< HEAD
    #If UsarWrench = 1 Then
        frmMain.Socket1.hostname = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
    #End If
=======
#If UsarWrench = 1 Then
    frmMain.Socket1.hostname = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
#End If
>>>>>>> origin/master
    
    EstadoLogin = E_MODO.CrearNuevoPj
    
    'Clear spell list
    frmMain.hlst.Clear
    
<<<<<<< HEAD
    #If UsarWrench = 1 Then

        If Not frmMain.Socket1.Connected Then
        #Else

            If frmMain.Winsock1.State <> sckConnected Then
            #End If
            MsgBox "Error: Se ha perdido la conexion con el server."
            Unload Me
=======
#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State <> sckConnected Then
#End If
        MsgBox JsonLanguage.Item("ERROR_CONN_LOST").Item("TEXTO")
        Unload Me
>>>>>>> origin/master
        
        Else
            Call Login

        End If
    
        bShowTutorial = True

    
    Exit Sub

ImgCrear_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "ImgCrear_Click"
    End If
Resume Next
    
    End Sub

Private Sub imgDados_Click()
    
    On Error GoTo imgDados_Click_Err
    
    Call Audio.PlayWave(SND_DICE)
    Call TirarDados

    
    Exit Sub

imgDados_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgDados_Click"
    End If
Resume Next
    
End Sub

Private Sub imgEspecialidad_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    
    On Error GoTo imgEspecialidad_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieEspecialidad)

    
    Exit Sub

imgEspecialidad_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgEspecialidad_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgNombre_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    On Error GoTo imgNombre_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieNombre)

    
    Exit Sub

imgNombre_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgNombre_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgPasswd_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    On Error GoTo imgPasswd_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.iePasswd)

    
    Exit Sub

imgPasswd_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgPasswd_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgConfirmPasswd_MouseMove(Button As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    
    On Error GoTo imgConfirmPasswd_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)

    
    Exit Sub

imgConfirmPasswd_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgConfirmPasswd_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgAtributos_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    
    On Error GoTo imgAtributos_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieAtributos)

    
    Exit Sub

imgAtributos_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgAtributos_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo imgD_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieD)

    
    Exit Sub

imgD_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgD_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo imgM_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieM)

    
    Exit Sub

imgM_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgM_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo imgF_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieF)

    
    Exit Sub

imgF_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgF_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgFuerza_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    On Error GoTo imgFuerza_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieFuerza)

    
    Exit Sub

imgFuerza_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgFuerza_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgAgilidad_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    
    On Error GoTo imgAgilidad_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieAgilidad)

    
    Exit Sub

imgAgilidad_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgAgilidad_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgInteligencia_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    
    On Error GoTo imgInteligencia_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieInteligencia)

    
    Exit Sub

imgInteligencia_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgInteligencia_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgCarisma_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo imgCarisma_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieCarisma)

    
    Exit Sub

imgCarisma_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgCarisma_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgConstitucion_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    
    On Error GoTo imgConstitucion_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieConstitucion)

    
    Exit Sub

imgConstitucion_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgConstitucion_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgArcos_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    On Error GoTo imgArcos_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieArcos)

    
    Exit Sub

imgArcos_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgArcos_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgArmas_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    On Error GoTo imgArmas_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieArmas)

    
    Exit Sub

imgArmas_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgArmas_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo imgEscudos_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieEscudos)

    
    Exit Sub

imgEscudos_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgEscudos_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo imgEvasion_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieEvasion)

    
    Exit Sub

imgEvasion_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgEvasion_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    On Error GoTo imgMagia_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieMagia)

    
    Exit Sub

imgMagia_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgMagia_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgMail_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    
    On Error GoTo imgMail_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieMail)

    
    Exit Sub

imgMail_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgMail_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgVida_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    
    On Error GoTo imgVida_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieVida)

    
    Exit Sub

imgVida_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgVida_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgTirarDados_MouseMove(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)
    
    On Error GoTo imgTirarDados_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieTirarDados)

    
    Exit Sub

imgTirarDados_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgTirarDados_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgPuebloOrigen_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    
    On Error GoTo imgPuebloOrigen_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.iePuebloOrigen)

    
    Exit Sub

imgPuebloOrigen_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgPuebloOrigen_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgRaza_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    
    On Error GoTo imgRaza_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieRaza)

    
    Exit Sub

imgRaza_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgRaza_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgClase_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    On Error GoTo imgClase_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieClase)

    
    Exit Sub

imgClase_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgClase_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgGenero_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    On Error GoTo imgGenero_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieGenero)

    
    Exit Sub

imgGenero_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgGenero_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgalineacion_MouseMove(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)
    
    On Error GoTo imgalineacion_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieAlineacion)

    
    Exit Sub

imgalineacion_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgalineacion_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgVolver_Click()
    
    On Error GoTo imgVolver_Click_Err
    
    Call Audio.PlayMIDI("2.mid")
    
    bShowTutorial = False
    
    Unload Me

    
    Exit Sub

imgVolver_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "imgVolver_Click"
    End If
Resume Next
    
End Sub

Private Sub lstGenero_Click()
    
    On Error GoTo lstGenero_Click_Err
    
    UserSexo = lstGenero.ListIndex + 1
    Call DarCuerpoYCabeza

    
    Exit Sub

lstGenero_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "lstGenero_Click"
    End If
Resume Next
    
End Sub

Private Sub lstProfesion_Click()

    On Error Resume Next

    '    Image1.Picture = LoadPicture(App.path & "\graficos\" & lstProfesion.Text & ".jpg")
    '
    UserClase = lstProfesion.ListIndex + 1
    
    Call UpdateStats
    Call UpdateEspecialidad(UserClase)

End Sub

Private Sub UpdateEspecialidad(ByVal eClase As eClass)
    
    On Error GoTo UpdateEspecialidad_Err
    
    lblEspecialidad.Caption = vEspecialidades(eClase)

    
    Exit Sub

UpdateEspecialidad_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "UpdateEspecialidad"
    End If
Resume Next
    
End Sub

Private Sub lstRaza_Click()
    
    On Error GoTo lstRaza_Click_Err
    
    UserRaza = lstRaza.ListIndex + 1
    Call DarCuerpoYCabeza
    
    Call UpdateStats

    
    Exit Sub

lstRaza_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "lstRaza_Click"
    End If
Resume Next
    
End Sub

Private Sub picHead_Click(Index As Integer)
    
    On Error GoTo picHead_Click_Err

    ' No se mueve si clickea al medio
    If Index = 2 Then Exit Sub
    
    Dim Counter             As Integer
    Dim Head                As Integer
    Dim Count_index         As Long
    Dim Count_index_reverse As Long
    
    Head = UserHead
    
    Count_index = Index - 2
    Count_index_reverse = 2 - Index
    
    If Index > 2 Then

        For Counter = Count_index To 1 Step -1
            Head = CheckCabeza(Head + 1)
        Next Counter

    Else

        For Counter = Count_index_reverse To 1 Step -1
            Head = CheckCabeza(Head - 1)
        Next Counter

    End If
    
    UserHead = Head
    
    Call UpdateHeadSelection
    
    Exit Sub

picHead_Click_Err:

    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "picHead_Click"

    End If

    Resume Next
    
End Sub

Private Sub tAnimacion_Timer()

    On Error Resume Next

    Dim SR       As RECT
    Dim DR       As RECT
    Dim Grh      As Long
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
        SR.bottom = SR.Top + .pixelHeight
        
        DR.Left = (picPJ.Width - .pixelWidth) \ 2 - 2
        DR.Top = ((picPJ.Height - .pixelHeight) \ 2) + 5
        DR.Right = DR.Left + .pixelWidth
        DR.bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picPJ.hdc, picTemp.hdc, DR, DR, vbBlack)

    End With
    
    Grh = HeadData(UserHead).Head(Dir).GrhIndex
    
    With GrhData(Grh)
        SR.Left = .SX
        SR.Top = .SY
        SR.Right = SR.Left + .pixelWidth
        SR.bottom = SR.Top + .pixelHeight
        
        DR.Left = (picPJ.Width - .pixelWidth) \ 2 - 2
        DR.Top = DR.bottom + BodyData(UserBody).HeadOffset.Y - .pixelHeight
        DR.Right = DR.Left + .pixelWidth
        DR.bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picPJ.hdc, picTemp.hdc, DR, DR, vbBlack)

    End With

End Sub

Private Sub DrawHead(ByVal Head As Integer, ByVal PicIndex As Integer)
    
    On Error GoTo DrawHead_Err
    

    Dim SR  As RECT
    Dim DR  As RECT
    Dim Grh As Long

    Call DrawImageInPicture(picHead(PicIndex), Me.Picture, 0, 0, , , picHead(PicIndex).Left, picHead(PicIndex).Top)
    
    Grh = HeadData(Head).Head(Dir).GrhIndex

    With GrhData(Grh)
        SR.Left = .SX
        SR.Top = .SY
        SR.Right = SR.Left + .pixelWidth
        SR.bottom = SR.Top + .pixelHeight
        
        DR.Left = (picHead(0).Width - .pixelWidth) \ 2 + 1
        DR.Top = 5
        DR.Right = DR.Left + .pixelWidth
        DR.bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        Call DrawTransparentGrhtoHdc(picHead(PicIndex).hdc, picTemp.hdc, DR, DR, vbBlack)

    End With
    
    
    Exit Sub

DrawHead_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "DrawHead"
    End If
Resume Next
    
End Sub

Private Sub txtConfirmPasswd_MouseMove(Button As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    
    On Error GoTo txtConfirmPasswd_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)

    
    Exit Sub

txtConfirmPasswd_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "txtConfirmPasswd_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub txtMail_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    
    On Error GoTo txtMail_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieMail)

    
    Exit Sub

txtMail_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "txtMail_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub txtNombre_Change()
    
    On Error GoTo txtNombre_Change_Err
    
    txtNombre.Text = LTrim$(txtNombre.Text)

    
    Exit Sub

txtNombre_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "txtNombre_Change"
    End If
Resume Next
    
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    
    On Error GoTo txtNombre_KeyPress_Err
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

    
    Exit Sub

txtNombre_KeyPress_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "txtNombre_KeyPress"
    End If
Resume Next
    
End Sub

Private Sub DarCuerpoYCabeza()
    
    On Error GoTo DarCuerpoYCabeza_Err
    

    Dim bVisible  As Boolean
    Dim PicIndex  As Integer
    Dim LineIndex As Integer
    
    Select Case UserSexo

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    UserHead = HUMANO_H_PRIMER_CABEZA
                    UserBody = HUMANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_H_PRIMER_CABEZA
                    UserBody = ELFO_H_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_H_PRIMER_CABEZA
                    UserBody = DROW_H_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_H_PRIMER_CABEZA
                    UserBody = ENANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_H_PRIMER_CABEZA
                    UserBody = GNOMO_H_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0

            End Select
            
        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    UserHead = HUMANO_M_PRIMER_CABEZA
                    UserBody = HUMANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_M_PRIMER_CABEZA
                    UserBody = ELFO_M_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_M_PRIMER_CABEZA
                    UserBody = DROW_M_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_M_PRIMER_CABEZA
                    UserBody = ENANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_M_PRIMER_CABEZA
                    UserBody = GNOMO_M_CUERPO_DESNUDO
                    
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

    If currentGrh > 0 Then tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)

    
    Exit Sub

DarCuerpoYCabeza_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "DarCuerpoYCabeza"
    End If
Resume Next
    
End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer
    
    On Error GoTo CheckCabeza_Err
    

    Select Case UserSexo

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano

                    If Head > HUMANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = HUMANO_H_PRIMER_CABEZA + (Head - HUMANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < HUMANO_H_PRIMER_CABEZA Then
                        CheckCabeza = HUMANO_H_ULTIMA_CABEZA - (HUMANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head

                    End If
                
                Case eRaza.Elfo

                    If Head > ELFO_H_ULTIMA_CABEZA Then
                        CheckCabeza = ELFO_H_PRIMER_CABEZA + (Head - ELFO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < ELFO_H_PRIMER_CABEZA Then
                        CheckCabeza = ELFO_H_ULTIMA_CABEZA - (ELFO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head

                    End If
                
                Case eRaza.ElfoOscuro

                    If Head > DROW_H_ULTIMA_CABEZA Then
                        CheckCabeza = DROW_H_PRIMER_CABEZA + (Head - DROW_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < DROW_H_PRIMER_CABEZA Then
                        CheckCabeza = DROW_H_ULTIMA_CABEZA - (DROW_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head

                    End If
                
                Case eRaza.Enano

                    If Head > ENANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = ENANO_H_PRIMER_CABEZA + (Head - ENANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < ENANO_H_PRIMER_CABEZA Then
                        CheckCabeza = ENANO_H_ULTIMA_CABEZA - (ENANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head

                    End If
                
                Case eRaza.Gnomo

                    If Head > GNOMO_H_ULTIMA_CABEZA Then
                        CheckCabeza = GNOMO_H_PRIMER_CABEZA + (Head - GNOMO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < GNOMO_H_PRIMER_CABEZA Then
                        CheckCabeza = GNOMO_H_ULTIMA_CABEZA - (GNOMO_H_PRIMER_CABEZA - Head) + 1
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

                    If Head > HUMANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = HUMANO_M_PRIMER_CABEZA + (Head - HUMANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < HUMANO_M_PRIMER_CABEZA Then
                        CheckCabeza = HUMANO_M_ULTIMA_CABEZA - (HUMANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head

                    End If
                
                Case eRaza.Elfo

                    If Head > ELFO_M_ULTIMA_CABEZA Then
                        CheckCabeza = ELFO_M_PRIMER_CABEZA + (Head - ELFO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < ELFO_M_PRIMER_CABEZA Then
                        CheckCabeza = ELFO_M_ULTIMA_CABEZA - (ELFO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head

                    End If
                
                Case eRaza.ElfoOscuro

                    If Head > DROW_M_ULTIMA_CABEZA Then
                        CheckCabeza = DROW_M_PRIMER_CABEZA + (Head - DROW_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < DROW_M_PRIMER_CABEZA Then
                        CheckCabeza = DROW_M_ULTIMA_CABEZA - (DROW_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head

                    End If
                
                Case eRaza.Enano

                    If Head > ENANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = ENANO_M_PRIMER_CABEZA + (Head - ENANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < ENANO_M_PRIMER_CABEZA Then
                        CheckCabeza = ENANO_M_ULTIMA_CABEZA - (ENANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head

                    End If
                
                Case eRaza.Gnomo

                    If Head > GNOMO_M_ULTIMA_CABEZA Then
                        CheckCabeza = GNOMO_M_PRIMER_CABEZA + (Head - GNOMO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < GNOMO_M_PRIMER_CABEZA Then
                        CheckCabeza = GNOMO_M_ULTIMA_CABEZA - (GNOMO_M_PRIMER_CABEZA - Head) + 1
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

    
    Exit Function

CheckCabeza_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "CheckCabeza"
    End If
Resume Next
    
End Function

Private Function CheckDir(ByRef Dir As E_Heading) As E_Heading
    
    On Error GoTo CheckDir_Err
    

    If Dir > E_Heading.WEST Then Dir = E_Heading.NORTH
    If Dir < E_Heading.NORTH Then Dir = E_Heading.WEST
    
    CheckDir = Dir
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex

    If currentGrh > 0 Then tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)

    
    Exit Function

CheckDir_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "CheckDir"
    End If
Resume Next
    
End Function

Private Sub LoadHelp()
<<<<<<< HEAD
    
    On Error GoTo LoadHelp_Err
    
    vHelp(eHelp.iePasswd) = "La contrase�a que utilizar�s para conectar tu personaje al juego."
    vHelp(eHelp.ieTirarDados) = "Presionando sobre la Esfera Roja, se modificar�n al azar los atributos de tu personaje, de esta manera puedes elegir los que m�s te parezcan para definir a tu personaje."
    vHelp(eHelp.ieMail) = "Es sumamente importante que ingreses una direcci�n de correo electr�nico v�lida, ya que en el caso de perder la contrase�a de tu personaje, se te enviar� cuando lo requieras, a esa direcci�n."
    vHelp(eHelp.ieNombre) = "S� cuidadoso al seleccionar el nombre de tu personaje. Argentum es un juego de rol, un mundo m�gico y fant�stico, y si seleccion�s un nombre obsceno o con connotaci�n pol�tica, los administradores borrar�n tu personaje y no habr� ninguna posibilidad de recuperarlo."
    vHelp(eHelp.ieConfirmPasswd) = "La contrase�a que utilizar�s para conectar tu personaje al juego."
    vHelp(eHelp.ieAtributos) = "Son las cualidades que definen tu personaje. Generalmente se los llama ""Dados"". (Ver Tirar Dados)"
    vHelp(eHelp.ieD) = "Son los atributos que obtuviste al azar. Presion� la esfera roja para volver a tirarlos."
    vHelp(eHelp.ieM) = "Son los modificadores por raza que influyen en los atributos de tu personaje."
    vHelp(eHelp.ieF) = "Los atributos finales de tu personaje, de acuerdo a la raza que elegiste."
    vHelp(eHelp.ieFuerza) = "De ella depender� qu� tan potentes ser�n tus golpes, tanto con armas de cuerpo a cuerpo, a distancia o sin armas."
    vHelp(eHelp.ieAgilidad) = "Este atributo intervendr� en qu� tan bueno seas, tanto evadiendo como acertando golpes, respecto de otros personajes como de las criaturas a las q te enfrentes."
    vHelp(eHelp.ieInteligencia) = "Influir� de manera directa en cu�nto man� ganar�s por nivel."
    vHelp(eHelp.ieCarisma) = "Ser� necesario tanto para la relaci�n con otros personajes (entrenamiento en parties) como con las criaturas (domar animales)."
    vHelp(eHelp.ieConstitucion) = "Afectar� a la cantidad de vida que podr�s ganar por nivel."
    vHelp(eHelp.ieEvasion) = "Eval�a la habilidad esquivando ataques f�sicos."
    vHelp(eHelp.ieMagia) = "Punt�a la cantidad de man� que se tendr�."
    vHelp(eHelp.ieVida) = "Valora la cantidad de salud que se podr� llegar a tener."
    vHelp(eHelp.ieEscudos) = "Estima la habilidad para rechazar golpes con escudos."
    vHelp(eHelp.ieArmas) = "Eval�a la habilidad en el combate cuerpo a cuerpo con armas."
    vHelp(eHelp.ieArcos) = "Eval�a la habilidad en el combate a distancia con arcos. "
    vHelp(eHelp.ieEspecialidad) = vbNullString
    vHelp(eHelp.iePuebloOrigen) = "Define el hogar de tu personaje. Sin embargo, el personaje nacer� en Nemahuak, la ciudad de los novatos."
    vHelp(eHelp.ieRaza) = "De la raza que elijas depender� c�mo se modifiquen los dados que saques. Pod�s cambiar de raza para poder visualizar c�mo se modifican los distintos atributos."
    vHelp(eHelp.ieClase) = "La clase influir� en las caracter�sticas principales que tenga tu personaje, asi como en las magias e items que podr� utilizar. Las estrellas que ves abajo te mostrar�n en qu� habilidades se destaca la misma."
    vHelp(eHelp.ieGenero) = "Indica si el personaje ser� masculino o femenino. Esto influye en los items que podr� equipar."
    vHelp(eHelp.ieAlineacion) = "Indica si el personaje seguir� la senda del mal o del bien. (Actualmente deshabilitado)"

    
    Exit Sub

LoadHelp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "LoadHelp"
    End If
Resume Next
    
=======
    vHelp(eHelp.ieTirarDados) = JsonLanguage.Item("AYUDA_CREARPJ_DADOS").Item("TEXTO")
    vHelp(eHelp.ieMail) = JsonLanguage.Item("AYUDA_CREARPJ_CORREO").Item("TEXTO")
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
    vHelp(eHelp.ieAlineacion) = JsonLanguage.Item("AYUDA_CREARPJ_ALINEACION").Item("TEXTO")
>>>>>>> origin/master
End Sub

Private Sub ClearLabel()
    
    On Error GoTo ClearLabel_Err
    
    LastButtonPressed.ToggleToNormal
    lblHelp = vbNullString

    
    Exit Sub

ClearLabel_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "ClearLabel"
    End If
Resume Next
    
End Sub

Private Sub txtNombre_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    On Error GoTo txtNombre_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.ieNombre)

    
    Exit Sub

txtNombre_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "txtNombre_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub txtPasswd_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    On Error GoTo txtPasswd_MouseMove_Err
    
    lblHelp.Caption = vHelp(eHelp.iePasswd)

    
    Exit Sub

txtPasswd_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "txtPasswd_MouseMove"
    End If
Resume Next
    
End Sub

Public Sub UpdateStats()
    
    On Error GoTo UpdateStats_Err
    
    Call UpdateRazaMod
    Call UpdateStars

    
    Exit Sub

UpdateStats_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "UpdateStats"
    End If
Resume Next
    
End Sub

Private Sub UpdateRazaMod()
    
    On Error GoTo UpdateRazaMod_Err
    
    Dim SelRaza As Integer
    Dim i       As Integer
    
    If lstRaza.ListIndex > -1 Then
    
        SelRaza = lstRaza.ListIndex + 1
        
        With ModRaza(SelRaza)
            lblModRaza(eAtributos.Fuerza).Caption = IIf(.Fuerza >= 0, "+", vbNullString) & .Fuerza
            lblModRaza(eAtributos.Agilidad).Caption = IIf(.Agilidad >= 0, "+", vbNullString) & .Agilidad
            lblModRaza(eAtributos.Inteligencia).Caption = IIf(.Inteligencia >= 0, "+", vbNullString) & .Inteligencia
            lblModRaza(eAtributos.Carisma).Caption = IIf(.Carisma >= 0, "+", "") & .Carisma
<<<<<<< HEAD
            lblModRaza(eAtributos.Constitucion).Caption = IIf(.Constitucion >= 0, "+", "") & .Constitucion

=======
            lblModRaza(eAtributos.Constitucion).Caption = IIf(.Constitucion >= 0, "+", vbNullString) & .Constitucion
>>>>>>> origin/master
        End With

    End If
    
    ' Atributo total
    For i = 1 To NUMATRIBUTES
        lblAtributoFinal(i).Caption = Val(lblAtributos(i).Caption) + Val(lblModRaza(i))
    Next i
    
    
    Exit Sub

UpdateRazaMod_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "UpdateRazaMod"
    End If
Resume Next
    
End Sub

Private Sub UpdateStars()
    
    On Error GoTo UpdateStars_Err
    
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
    NumStars = (0.509 + 0.01185 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * ModClase(UserClase).Hit * ModClase(UserClase).Da�oArmas + 0.119 * ModClase(UserClase).AtaqueArmas * Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArmasStar, NumStars * 2)
    
    ' Estrellas de arcos
    NumStars = (0.4915 + 0.01265 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * ModClase(UserClase).Da�oProyectiles * ModClase(UserClase).Hit + 0.119 * ModClase(UserClase).AtaqueProyectiles * Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
    Call SetStars(imgArcoStar, NumStars * 2)

    
    Exit Sub

UpdateStars_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "UpdateStars"
    End If
Resume Next
    
End Sub

Private Sub SetStars(ByRef ImgContainer As Object, ByVal NumStars As Integer)
    
    On Error GoTo SetStars_Err
    
    Dim FullStars   As Integer
    Dim HasHalfStar As Boolean
    Dim Index       As Integer
    Dim Counter     As Integer

    If NumStars > 0 Then
        
        If NumStars > 10 Then NumStars = 10
        
        FullStars = Int(NumStars / 2)
        
        ' Tienen brillo extra si estan todas
        If FullStars = 5 Then

            For Index = 1 To FullStars
                ImgContainer(Index).Picture = picGlowStar
            Next Index

        Else

            ' Numero impar? Entonces hay que poner "media estrella"
            If (NumStars Mod 2) > 0 Then HasHalfStar = True
            
            ' Muestro las estrellas enteras
            If FullStars > 0 Then

                For Index = 1 To FullStars
                    ImgContainer(Index).Picture = picFullStar
                Next Index
                
                Counter = FullStars

            End If
            
            ' Muestro la mitad de la estrella (si tiene)
            If HasHalfStar Then
                Counter = Counter + 1
                
                ImgContainer(Counter).Picture = picHalfStar

            End If
            
            ' Si estan completos los espacios, no borro nada
            If Counter <> 5 Then
                Counter = Counter + 1
                
                ' Limpio las que queden vacias
                For Index = Counter To 5
                    Set ImgContainer(Index).Picture = Nothing
                Next Index

            End If
            
        End If

    Else

        ' Limpio todo
        For Index = 1 To 5
            Set ImgContainer(Index).Picture = Nothing
        Next Index

    End If

    
    Exit Sub

SetStars_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "SetStars"
    End If
Resume Next
    
End Sub

Private Sub LoadCharInfo()
    
    On Error GoTo LoadCharInfo_Err
    
    Dim SearchVar As String
    Dim i         As Integer
    
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
            .Da�oArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODDA�OARMAS", SearchVar))
            .Da�oProyectiles = Val(GetVar(IniPath & "CharInfo.dat", "MODDA�OPROYECTILES", SearchVar))
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

    
    Exit Sub

LoadCharInfo_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "LoadCharInfo"
    End If
Resume Next
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyUp_Err
    

    If KeyCode = vbKeyEscape Then
        Call Audio.PlayMIDI("2.mid")
        bShowTutorial = False
        Unload Me
    End If


    Exit Sub

Form_KeyUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearPersonaje" & "->" & "Form_KeyUp"
    End If
Resume Next
    
End Sub
