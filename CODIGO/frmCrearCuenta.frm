VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BorderStyle     =   0  'None
   Caption         =   "Crear Cuenta"
   ClientHeight    =   4530
   ClientLeft      =   5115
   ClientTop       =   4125
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmCrearCuenta.frx":0000
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
      Picture         =   "frmCrearCuenta.frx":20F13
      Top             =   3800
      Width           =   1335
   End
   Begin VB.Image imgCrearCuenta 
      Height          =   375
      Left            =   4520
      Picture         =   "frmCrearCuenta.frx":24A8B
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
    
    On Error GoTo Form_Load_Err
    
    Me.Picture = LoadPicture(App.path & "\graficos\frmCuentaNueva.jpg")
    txtCuentaEmail.Text = vbNullString
    txtCuentaPassword.Text = vbNullString
    txtCuentaRepite.Text = vbNullString
    
    LoadButtons
    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearCuenta" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub imgCrearCuenta_Click()
    
    On Error GoTo imgCrearCuenta_Click_Err
    

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

    
    Exit Sub

imgCrearCuenta_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearCuenta" & "->" & "imgCrearCuenta_Click"
    End If
Resume Next
    
End Sub

Private Sub imgSalir_Click()
    
    On Error GoTo imgSalir_Click_Err
    
    Unload Me

    
    Exit Sub

imgSalir_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearCuenta" & "->" & "imgSalir_Click"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearCuenta" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Function IsFormValid() As Boolean
    
    On Error GoTo IsFormValid_Err
    

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

    
    Exit Function

IsFormValid_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearCuenta" & "->" & "IsFormValid"
    End If
Resume Next
    
End Function

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos
    
    Set cBotonCrearCuenta = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

    Call cBotonCrearCuenta.Initialize(Me.imgCrearCuenta, GrhPath & "BotonCrearCuenta.jpg", GrhPath & "BotonCrearCuentaRollover.jpg", GrhPath & "BotonCrearCuentaClick.jpg", Me)

    Call cBotonSalir.Initialize(Me.imgSalir, GrhPath & "BotonSalirConnect.jpg", GrhPath & "BotonBotonSalirRolloverConnect.jpg", GrhPath & "BotonSalirClickConnect.jpg", Me)

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearCuenta" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyUp_Err
    
    If KeyCode = vbKeyEscape Then
        Set clsFormulario = Nothing
        Unload Me
    End If

    Exit Sub

Form_KeyUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCrearCuenta" & "->" & "Form_KeyUp"
    End If
Resume Next
    
End Sub
