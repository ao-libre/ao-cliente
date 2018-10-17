VERSION 5.00
Begin VB.Form frmCrearCuenta 
   Caption         =   "Crear Cuenta"
   ClientHeight    =   3015
   ClientLeft      =   5235
   ClientTop       =   4590
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.TextBox txtCuentaRepite 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "X"
      TabIndex        =   7
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtCuentaPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "X"
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtCuentaEmail 
      Height          =   375
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmdCrearCuenta 
      Caption         =   "Crear"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Repite"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Pass"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Email"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Crear nueva cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCrearCuenta_Click()
EstadoLogin = E_MODO.CrearCuenta

AccountName = txtCuentaEmail.Text
AccountPassword = txtCuentaPassword.Text
'CHOTS | @TODO validar mail y password

#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

End Sub

Private Sub Form_Load()
txtCuentaEmail.Text = ""
txtCuentaPassword.Text = ""
txtCuentaRepite.Text = ""
End Sub
