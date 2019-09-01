VERSION 5.00
Begin VB.Form frmRecuperarCuenta 
   Caption         =   "Recuperar Cuenta"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmRecuperarCuenta.frx":0000
   ScaleHeight     =   3300
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtContrasena 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox txtCorreo 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin AOLibre.uAOButton cmdProcesar 
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      TX              =   "Enviar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmRecuperarCuenta.frx":DB91
      PICF            =   "frmRecuperarCuenta.frx":E5BB
      PICH            =   "frmRecuperarCuenta.frx":F27D
      PICV            =   "frmRecuperarCuenta.frx":1020F
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
   Begin VB.Label lblPass 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   2385
   End
   Begin VB.Label lblCorreo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Correo Electronico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   2385
   End
End
Attribute VB_Name = "frmRecuperarCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaRecuperarCuenta.jpg")
    Me.lblCorreo.Caption = JsonLanguage.Item("FRM_RECUPERAR_CUENTA_LBLCORREO").Item("TEXTO")
    Me.lblPass.Caption = JsonLanguage.Item("FRM_RECUPERAR_CUENTA_LBLPASS").Item("TEXTO")
    Me.cmdProcesar.Caption = JsonLanguage.Item("FRM_RECUPERAR_CUENTA_CMDENVIAR").Item("TEXTO")
End Sub

Private Sub cmdProcesar_Click()
    
    AccountMailToRecover = txtCorreo.Text
    AccountNewPassword = txtContrasena.Text
    
    If LenB(AccountMailToRecover) <> 0 And LenB(AccountNewPassword) <> 0 Then
         If CheckMailString(AccountMailToRecover) Then
               Call Login
         End If
    End If

    Unload Me
End Sub
