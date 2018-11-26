VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   5.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet InetReddit 
      Left            =   120
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox lstRedditPosts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   4320
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.ListBox lstServers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4905
      ItemData        =   "frmConnect.frx":000C
      Left            =   8685
      List            =   "frmConnect.frx":000E
      TabIndex        =   6
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
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
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3720
      Width           =   2460
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
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
      Left            =   4905
      TabIndex        =   4
      Top             =   3210
      Width           =   2460
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4890
      TabIndex        =   2
      Text            =   "7666"
      Top             =   2760
      Width           =   825
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5760
      TabIndex        =   3
      Text            =   "localhost"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Image imgTeclas 
      Height          =   375
      Left            =   6120
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Image imgConectarse 
      Height          =   375
      Left            =   4800
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Image imgVerForo 
      Height          =   465
      Left            =   450
      Top             =   6120
      Width           =   2835
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   9960
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgBorrarPj 
      Height          =   375
      Left            =   8400
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgCodigoFuente 
      Height          =   375
      Left            =   6840
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgReglamento 
      Height          =   375
      Left            =   5280
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgManual 
      Height          =   375
      Left            =   3720
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgRecuperar 
      Height          =   375
      Left            =   2160
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgCrearPj 
      Height          =   375
      Left            =   600
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   555
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

Private cBotonCrearPj As clsGraphicalButton
Private cBotonRecuperarPass As clsGraphicalButton
Private cBotonManual As clsGraphicalButton
Private cBotonReglamento As clsGraphicalButton
Private cBotonCodigoFuente As clsGraphicalButton
Private cBotonBorrarPj As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton
Private cBotonLeerMas As clsGraphicalButton
Private cBotonForo As clsGraphicalButton
Private cBotonConectarse As clsGraphicalButton
Private cBotonTeclas As clsGraphicalButton

Private Type tRedditPost
    Title As String
    URL As String
End Type

Dim Posts() As tRedditPost

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Activate()
'On Error Resume Next
    Call CargarServidores
    
    If ServersRecibidos Then
        If CurServer <> 0 Then
            IPTxt = ServersLst(1).Ip
            PortTxt = ServersLst(1).Puerto
        Else
            IPTxt = IPdelServidor
            PortTxt = PuertoDelServidor
        End If
    End If
    
    Call GetPostsFromReddit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        prgRun = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    'Port
    PortTxt.Visible = True
    'Label4.Visible = True
    
    'Server IP
    PortTxt.Text = "7666"
    IPTxt.Text = "192.168.0.2"
    IPTxt.Visible = True
    'Label5.Visible = True
    
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]

    PortTxt.Text = Config_Inicio.Puerto
 
     '[CODE]:MatuX
    '
    '  El código para mostrar la versión se genera acá para
    ' evitar que por X razones luego desaparezca, como suele
    ' pasar a veces :)
       version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    '[END]'
    
    Me.Picture = LoadPicture(App.path & "\graficos\VentanaConectar.jpg")
    
    Call LoadButtons

    Call CheckLicenseAgreement
        
End Sub

Private Sub CheckLicenseAgreement()
    'Recordatorio para cumplir la licencia, por si borrás el Boton sin leer el code...
    Dim i As Long
    
    For i = 0 To Me.Controls.Count - 1
        If Me.Controls(i).Name = "imgCodigoFuente" Then
            Exit For
        End If
    Next i
    
    If i = Me.Controls.Count Then
        MsgBox "No debe eliminarse la posibilidad de bajar el código de sus servidor. Caso contrario estarían violando la licencia Affero GPL y con ella derechos de autor, incurriendo de esta forma en un delito punible por ley." & vbCrLf & vbCrLf & vbCrLf & _
                "Argentum Online es libre, es de todos. Mantengamoslo así. Si tanto te gusta el juego y querés los cambios que hacemos nosotros, compartí los tuyos. Es un cambio justo. Si no estás de acuerdo, no uses nuestro código, pues nadie te obliga o bien utiliza una versión anterior a la 0.12.0.", vbCritical Or vbApplicationModal
    End If

End Sub

Private Sub LoadButtons()
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos
    
    Set cBotonCrearPj = New clsGraphicalButton
    Set cBotonRecuperarPass = New clsGraphicalButton
    Set cBotonManual = New clsGraphicalButton
    Set cBotonReglamento = New clsGraphicalButton
    Set cBotonCodigoFuente = New clsGraphicalButton
    Set cBotonBorrarPj = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    Set cBotonLeerMas = New clsGraphicalButton
    Set cBotonForo = New clsGraphicalButton
    Set cBotonConectarse = New clsGraphicalButton
    Set cBotonTeclas = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

        
    Call cBotonCrearPj.Initialize(imgCrearPj, GrhPath & "BotonCrearPersonajeConectar.jpg", _
                                    GrhPath & "BotonCrearPersonajeRolloverConectar.jpg", _
                                    GrhPath & "BotonCrearPersonajeClickConectar.jpg", Me)
                                    
    Call cBotonRecuperarPass.Initialize(imgRecuperar, GrhPath & "BotonRecuperarPass.jpg", _
                                    GrhPath & "BotonRecuperarPassRollover.jpg", _
                                    GrhPath & "BotonRecuperarPassClick.jpg", Me)
                                    
    Call cBotonManual.Initialize(imgManual, GrhPath & "BotonManual.jpg", _
                                    GrhPath & "BotonManualRollover.jpg", _
                                    GrhPath & "BotonManualClick.jpg", Me)
                                    
    Call cBotonReglamento.Initialize(imgReglamento, GrhPath & "BotonReglamento.jpg", _
                                    GrhPath & "BotonReglamentoRollover.jpg", _
                                    GrhPath & "BotonReglamentoClick.jpg", Me)
                                    
    Call cBotonCodigoFuente.Initialize(imgCodigoFuente, GrhPath & "BotonCodigoFuente.jpg", _
                                    GrhPath & "BotonCodigoFuenteRollover.jpg", _
                                    GrhPath & "BotonCodigoFuenteClick.jpg", Me)
                                    
    Call cBotonBorrarPj.Initialize(imgBorrarPj, GrhPath & "BotonBorrarPersonaje.jpg", _
                                    GrhPath & "BotonBorrarPersonajeRollover.jpg", _
                                    GrhPath & "BotonBorrarPersonajeClick.jpg", Me)
                                    
    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonSalirConnect.jpg", _
                                    GrhPath & "BotonBotonSalirRolloverConnect.jpg", _
                                    GrhPath & "BotonSalirClickConnect.jpg", Me)
                                    
    Call cBotonForo.Initialize(imgVerForo, GrhPath & "BotonVerForo.jpg", _
                                    GrhPath & "BotonVerForoRollover.jpg", _
                                    GrhPath & "BotonVerForoClick.jpg", Me)
                                    
    Call cBotonConectarse.Initialize(imgConectarse, GrhPath & "BotonConectarse.jpg", _
                                    GrhPath & "BotonConectarseRollover.jpg", _
                                    GrhPath & "BotonConectarseClick.jpg", Me)
                                    
    Call cBotonTeclas.Initialize(imgTeclas, GrhPath & "BotonTeclas.jpg", _
                                    GrhPath & "BotonTeclasRollover.jpg", _
                                    GrhPath & "BotonTeclasClick.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub CheckServers()
    If ServersRecibidos Then
        If Not IsIp(IPTxt) And CurServer <> 0 Then
            If MsgBox("Atencion, está intentando conectarse a un servidor no oficial, NoLand Studios no se hace responsable de los posibles problemas que estos servidores presenten. ¿Desea continuar?", vbYesNo) = vbNo Then
                If CurServer <> 0 Then
                    IPTxt = ServersLst(CurServer).Ip
                    PortTxt = ServersLst(CurServer).Puerto
                Else
                    IPTxt = IPdelServidor
                    PortTxt = PuertoDelServidor
                End If
                Exit Sub
            End If
        End If
    End If
    CurServer = 0
    IPdelServidor = IPTxt
    PuertoDelServidor = PortTxt
End Sub

Private Sub imgBorrarPj_Click()

On Error GoTo errH
    Call Shell(App.path & "\RECUPERAR.EXE", vbNormalFocus)

    Exit Sub

errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub

Private Sub imgCodigoFuente_Click()
'***********************************
'IMPORTANTE!
'
'No debe eliminarse la posibilidad de bajar el código de sus servidor de esta forma.
'Caso contrario estarían violando la licencia Affero GPL y con ella derechos de autor,
'incurriendo de esta forma en un delito punible por ley.
'
'Argentum Online es libre, es de todos. Mantengamoslo así. Si tanto te gusta el juego y querés los
'cambios que hacemos nosotros, compartí los tuyos. Es un cambio justo. Si no estás de acuerdo,
'no uses nuestro código, pues nadie te obliga o bien utiliza una versión anterior a la 0.12.0.
'***********************************
    Call ShellExecute(0, "Open", "https://github.com/ao-libre", "", App.path, SW_SHOWNORMAL)

End Sub

Private Sub imgConectarse_Click()
    Call CheckServers
    
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
    
    'update user info
    UserName = txtNombre.Text
    
    Dim aux As String
    aux = txtPasswd.Text
    UserPassword = aux
    'Clear spell list
    frmMain.hlst.Clear

    If CheckUserData(False) = True Then
        EstadoLogin = Normal
        
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

    End If
    
End Sub

Private Sub imgCrearPj_Click()
    
    Call CheckServers
    
    EstadoLogin = E_MODO.Dados
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

Private Sub imgLeerMas_Click()
    Call ShellExecute(0, "Open", "http://www.argentumonline.org", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgManual_Click()
    Call ShellExecute(0, "Open", "http://www.argentumonline.org", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgRecuperar_Click()
On Error GoTo errH

    Call Audio.PlayWave(SND_CLICK)
    Call Shell(App.path & "\RECUPERAR.EXE", vbNormalFocus)
    Exit Sub
errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub

Private Sub imgReglamento_Click()
    Call ShellExecute(0, "Open", "http://www.argentumonline.org", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgSalir_Click()
    prgRun = False
End Sub

Private Sub imgServArgentina_Click()
    Call Audio.PlayWave(SND_CLICK)
    IPTxt.Text = IPdelServidor
    PortTxt.Text = PuertoDelServidor
End Sub

Private Sub imgTeclas_Click()
    Load frmKeypad
    frmKeypad.Show vbModal
    Unload frmKeypad
    txtPasswd.SetFocus
End Sub

Private Sub imgVerForo_Click()
    Call ShellExecute(0, "Open", "https://www.reddit.com/r/argentumonlineoficial/", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub lstRedditPosts_Click()
    Call ShellExecute(0, "Open", Posts(lstRedditPosts.ListIndex + 1).URL, "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub lstServers_Click()
    IPTxt.Text = ServersLst(lstServers.ListIndex + 1).Ip
    PortTxt.Text = ServersLst(lstServers.ListIndex + 1).Puerto
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then imgConectarse_Click
End Sub

Private Sub GetPostsFromReddit()
    Dim ResponseGithub As String
    Dim JsonObject As Object
    Dim Endpoint As String
    
    Endpoint = GetVar(App.path & "\INIT\Config.ini", "Parameters", "SubRedditEndpoint")
    ResponseGithub = InetReddit.OpenURL(Endpoint)
    Set JsonObject = JSON.parse(ResponseGithub)
    
    Dim qtyPostsOnReddit As Integer
    qtyPostsOnReddit = JsonObject.Item("data").Item("children").Count
    ReDim Posts(qtyPostsOnReddit)
    
    Dim i As Integer
    i = 1
    Do While i <= qtyPostsOnReddit
        Posts(i).Title = JsonObject.Item("data").Item("children").Item(i).Item("data").Item("title")
        Posts(i).URL = JsonObject.Item("data").Item("children").Item(i).Item("data").Item("url")
        
        lstRedditPosts.AddItem JsonObject.Item("data").Item("children").Item(i).Item("data").Item("title")
        
        i = i + 1
    Loop
End Sub


