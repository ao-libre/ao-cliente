VERSION 5.00
Begin VB.Form frmCuenta 
   Caption         =   "Mi Cuenta"
   ClientHeight    =   3015
   ClientLeft      =   5010
   ClientTop       =   4155
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear PJ"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox lstPersonajes 
      Height          =   1620
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdConectar 
      Caption         =   "Conectar"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblCuentaNombre 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
AccountName = vbNullString
AccountHash = vbNullString
NumberOfCharacters = 0
Unload Me
End Sub

Private Sub cmdConectar_Click()
If lstPersonajes.ListIndex >= 0 Then
    #If UsarWrench = 1 Then
        If Not frmMain.Socket1.Connected Then
    #Else
        If frmMain.Winsock1.State <> sckConnected Then
    #End If
        MsgBox "Error: Se ha perdido la conexion con el server."
        AccountName = vbNullString
        AccountHash = vbNullString
        NumberOfCharacters = 0
        Unload Me
    Else
        UserName = lstPersonajes.List(lstPersonajes.ListIndex)
        Call WriteLoginExistingChar
    End If
End If
End Sub

Private Sub cmdCrear_Click()
    frmCrearPersonaje.Show
End Sub

Private Sub Form_Load()
    Dim i As Byte
    lblCuentaNombre.Caption = AccountName

    If NumberOfCharacters > 0 Then
        For i = 1 To NumberOfCharacters
            Call lstPersonajes.AddItem(UserCharacters(i))
        Next i
    End If
End Sub
