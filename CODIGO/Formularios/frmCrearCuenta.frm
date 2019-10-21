VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BorderStyle     =   0  'None
   Caption         =   "Crear Cuenta"
   ClientHeight    =   4530
   ClientLeft      =   5115
   ClientTop       =   4125
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   Picture         =   "frmCrearCuenta.frx":0000
   ScaleHeight     =   4530
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCuentaRepite 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   200
      IMEMode         =   3  'DISABLE
      Left            =   2680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3060
      Width           =   2480
   End
   Begin VB.TextBox txtCuentaPassword 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   200
      IMEMode         =   3  'DISABLE
      Left            =   2680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2550
      Width           =   2480
   End
   Begin VB.TextBox txtCuentaEmail 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   200
      Left            =   2680
      MaxLength       =   50
      TabIndex        =   0
      Top             =   2020
      Width           =   2480
   End
   Begin AOLibre.uAOButton imgSalir 
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCrearCuenta.frx":20F13
      PICF            =   "frmCrearCuenta.frx":2193D
      PICH            =   "frmCrearCuenta.frx":225FF
      PICV            =   "frmCrearCuenta.frx":23591
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
   Begin AOLibre.uAOButton imgCrearCuenta 
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      TX              =   "Crear Cuenta"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCrearCuenta.frx":24493
      PICF            =   "frmCrearCuenta.frx":24EBD
      PICH            =   "frmCrearCuenta.frx":25B7F
      PICV            =   "frmCrearCuenta.frx":26B11
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
   Begin VB.Label lblMensaje 
      BackStyle       =   0  'Transparent
      Caption         =   "Utilize un email real para recibir correctamente los correos"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   1440
      Width           =   3495
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    Me.Picture = LoadPicture(Game.path(Interfaces) & "frmCuentaNueva.jpg")
    txtCuentaEmail.Text = vbNullString
    txtCuentaPassword.Text = vbNullString
    txtCuentaRepite.Text = vbNullString

   Call LoadTextsForm
End Sub

Private Sub LoadTextsForm()
    imgCrearCuenta.Caption = JsonLanguage.Item("FRM_CREARCUENTA_CREARCUENTA").Item("TEXTO")
    imgSalir.Caption = JsonLanguage.Item("FRM_CREARCUENTA_SALIR").Item("TEXTO")
    lblMensaje.Caption = JsonLanguage.Item("FRM_CREARCUENTA_MENSAJE").Item("TEXTO")
End Sub

Private Sub imgCrearCuenta_Click()

    If Not IsFormValid Then Exit Sub
    
    AccountName = txtCuentaEmail.Text
    AccountPassword = txtCuentaPassword.Text
    
    Call Login

    Unload frmCrearCuenta
End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub

Private Function IsFormValid() As Boolean

    If Len(txtCuentaEmail.Text) = 0 Then
        MsgBox JsonLanguage.Item("VALIDACION_EMAIL").Item("TEXTO")
        Exit Function
    End If
    
    If Len(txtCuentaPassword.Text) = 0 Then
        MsgBox JsonLanguage.Item("VALIDACION_PASSWORD").Item("TEXTO")
        Exit Function
    End If

    If Not CheckMailString(txtCuentaEmail.Text) Then
        MsgBox JsonLanguage.Item("VALIDACION_BAD_EMAIL").Item("TEXTO").Item(1)
        Exit Function
    End If
    
    If Len(txtCuentaEmail.Text) > 30 Then
        MsgBox JsonLanguage.Item("VALIDACION_BAD_EMAIL").Item("TEXTO").Item(2)
        Exit Function
    End If
    
    If Not txtCuentaPassword.Text = txtCuentaRepite.Text Then
        MsgBox JsonLanguage.Item("VALIDACION_BAD_PASSWORD").Item("TEXTO").Item(1)
        Exit Function
    End If
    
    IsFormValid = True
End Function
