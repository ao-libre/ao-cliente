VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BorderStyle     =   0  'None
   Caption         =   "Crear Cuenta"
   ClientHeight    =   4530
   ClientLeft      =   5115
   ClientTop       =   4125
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
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
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   560
      Top             =   3800
      Width           =   1335
   End
   Begin VB.Image imgCrearCuenta 
      Height          =   375
      Left            =   4520
      Top             =   3800
      Width           =   1335
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cBotonCrearCuenta As clsGraphicalButton
Private cBotonSalir       As clsGraphicalButton

Public LastButtonPressed  As clsGraphicalButton

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.path & "\graficos\frmCuentaNueva.jpg")
    txtCuentaEmail.Text = ""
    txtCuentaPassword.Text = ""
    txtCuentaRepite.Text = ""
    LoadButtons

End Sub

Private Sub imgCrearCuenta_Click()

    If Not IsFormValid Then Exit Sub

    EstadoLogin = E_MODO.CrearCuenta
    
    AccountName = txtCuentaEmail.Text
    AccountPassword = txtCuentaPassword.Text
    
    #If UsarWrench = 1 Then

        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents

        End If

        frmMain.Socket1.hostname = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
    #Else

        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents

        End If

        frmMain.Winsock1.Connect CurServerIp, CurServerPort
    #End If

    Unload frmCrearCuenta

End Sub

Private Sub imgSalir_Click()
    Unload Me

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal

End Sub

Private Function IsFormValid() As Boolean

    If txtCuentaEmail.Text = "" Then
        MsgBox "Ingrese un e-mail."
        Exit Function

    End If
    
    If txtCuentaPassword.Text = "" Then
        MsgBox "Ingrese un password."
        Exit Function

    End If

    If Not CheckMailString(txtCuentaEmail.Text) Then
        MsgBox "Direccion de e-mail invalida."
        Exit Function

    End If
    
    If Len(txtCuentaEmail.Text) > 30 Then
        MsgBox "El e-mail debe tener menos de 30 letras."
        Exit Function

    End If
    
    If Not txtCuentaPassword.Text = txtCuentaRepite.Text Then
        MsgBox "Los passwords no coinciden."
        Exit Function

    End If
    
    IsFormValid = True

End Function

Private Sub LoadButtons()
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos
    
    Set cBotonCrearCuenta = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

    Call cBotonCrearCuenta.Initialize(Me.imgCrearCuenta, GrhPath & "BotonCrearCuenta.jpg", GrhPath & "BotonCrearCuentaRollover.jpg", GrhPath & "BotonCrearCuentaClick.jpg", Me)

    Call cBotonSalir.Initialize(Me.imgSalir, GrhPath & "BotonSalirConnect.jpg", GrhPath & "BotonBotonSalirRolloverConnect.jpg", GrhPath & "BotonSalirClickConnect.jpg", Me)

End Sub
