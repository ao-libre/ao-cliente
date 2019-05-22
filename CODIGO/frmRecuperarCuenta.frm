VERSION 5.00
Begin VB.Form frmRecuperarCuenta 
   Caption         =   "Recuperar Cuenta"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4305
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
   ScaleHeight     =   3030
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtContrasena 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox txtCorreo 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lblPass 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva Contraseña"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   2
      Top             =   1320
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
      Height          =   330
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2385
   End
End
Attribute VB_Name = "frmRecuperarCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdProcesar_Click()
    
    #If UsarWrench = 1 Then

        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents

        End If

    #Else

        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents

        End If

    #End If
    
    If CheckMailString(txtCorreo.Text) Then
    
        EstadoLogin = E_MODO.CambiarContrasena
        
        #If UsarWrench = 1 Then
            frmMain.Socket1.hostname = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
        #Else
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
        #End If
    
    Else
        
        Call MsgBox("Formato de correo electronico invalido.")
    
    End If
    
End Sub
