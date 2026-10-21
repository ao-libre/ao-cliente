VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.1#0"; "MSWINSCK.OCX"
Begin VB.Form frmObserver
   Caption         =   "[OBSERVE MODE] aobench Observer"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock ObsSocket
      Left            =   5280
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConnect
      Caption         =   "Observar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtCharname
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   ""
      Top             =   120
      Width           =   3495
   End
   Begin VB.ListBox lstState
      Height          =   3600
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5775
   End
   Begin VB.Label lblStatus
      Caption         =   "Desconectado"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   5775
   End
End
Attribute VB_Name = "frmObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' aobench: observer mode form
' Connects to BotAPI port 4200, sends loginObserver, displays incoming state JSON
Option Explicit

Private m_buffer As String

Private Sub Form_Load()
    lblStatus.Caption = "Desconectado"
    txtCharname.SetFocus
End Sub

Private Sub cmdConnect_Click()
    Dim charname As String
    charname = Trim(txtCharname.Text)
    If Len(charname) = 0 Then
        MsgBox "Ingresa el nombre del personaje a observar.", vbInformation
        Exit Sub
    End If

    If ObsSocket.State <> sckClosed Then ObsSocket.Close

    ObsSocket.RemoteHost = "127.0.0.1"
    ObsSocket.RemotePort = 4200
    ObsSocket.Connect

    lblStatus.Caption = "Conectando..."
End Sub

Private Sub ObsSocket_Connect()
    Dim charname As String
    charname = Trim(txtCharname.Text)
    Dim msg As String
    msg = "{""action"":""loginObserver"",""charname"":""" & charname & """}"
    ObsSocket.SendData msg & Chr(10)

    ObserveMode = True
    lblStatus.Caption = "[OBSERVE MODE] Observando: " & charname
End Sub

Private Sub ObsSocket_DataArrival(ByVal bytesTotal As Long)
    Dim chunk As String
    ObsSocket.GetData chunk, vbString
    m_buffer = m_buffer & chunk

    Dim pos As Integer
    pos = InStr(m_buffer, Chr(10))
    Do While pos > 0
        Dim line As String
        line = Left$(m_buffer, pos - 1)
        m_buffer = Mid$(m_buffer, pos + 1)
        lstState.AddItem line
        If lstState.ListCount > 100 Then lstState.RemoveItem 0
        lstState.ListIndex = lstState.ListCount - 1
        pos = InStr(m_buffer, Chr(10))
    Loop
End Sub

Private Sub ObsSocket_Close()
    ObserveMode = False
    lblStatus.Caption = "Desconectado"
End Sub

Private Sub ObsSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    ObserveMode = False
    lblStatus.Caption = "Error: " & Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ObsSocket.State <> sckClosed Then ObsSocket.Close
    ObserveMode = False
End Sub
