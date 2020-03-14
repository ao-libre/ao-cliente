VERSION 5.00
Begin VB.Form frmPanelAccount 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Panel de Cuenta"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPanelAccount.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin AOLibre.uAOButton uAOBorrarPersonaje 
      Height          =   615
      Left            =   840
      TabIndex        =   27
      Top             =   6360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Borrar Personaje"
      ENAB            =   -1  'True
      FCOL            =   255
      OCOL            =   16777215
      PICE            =   "frmPanelAccount.frx":678C1
      PICF            =   "frmPanelAccount.frx":678DD
      PICH            =   "frmPanelAccount.frx":678F9
      PICV            =   "frmPanelAccount.frx":67915
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   9
      Left            =   8760
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   9
      Top             =   3570
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   8
      Left            =   7005
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   8
      Top             =   3570
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   7
      Left            =   5355
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   7
      Top             =   3570
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   6
      Left            =   3660
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   6
      Top             =   3570
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   5
      Left            =   1965
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   5
      Top             =   3570
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   4
      Left            =   8760
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   4
      Top             =   1695
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   3
      Left            =   7005
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   3
      Top             =   1695
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   2
      Left            =   5325
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   2
      Top             =   1695
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   1
      Left            =   3675
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   1
      Top             =   1695
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1191
      Index           =   0
      Left            =   1920
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   0
      Top             =   1695
      Width           =   1140
   End
   Begin AOLibre.uAOButton uAOConectar 
      Height          =   615
      Left            =   9120
      TabIndex        =   28
      Top             =   7920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Conectar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmPanelAccount.frx":67931
      PICF            =   "frmPanelAccount.frx":6794D
      PICH            =   "frmPanelAccount.frx":67969
      PICV            =   "frmPanelAccount.frx":67985
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton uAOCrearPersonaje 
      Height          =   615
      Left            =   5040
      TabIndex        =   29
      Top             =   7920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Crear Personaje"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmPanelAccount.frx":679A1
      PICF            =   "frmPanelAccount.frx":679BD
      PICH            =   "frmPanelAccount.frx":679D9
      PICV            =   "frmPanelAccount.frx":679F5
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton uAOSalir 
      Height          =   615
      Left            =   960
      TabIndex        =   30
      Top             =   7920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmPanelAccount.frx":67A11
      PICF            =   "frmPanelAccount.frx":67A2D
      PICH            =   "frmPanelAccount.frx":67A49
      PICV            =   "frmPanelAccount.frx":67A65
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   1890
      TabIndex        =   11
      Top             =   3094
      Width           =   1245
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   5
      Left            =   3240
      TabIndex        =   26
      Top             =   7215
      Width           =   45
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   3240
      TabIndex        =   25
      Top             =   6885
      Width           =   45
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   3240
      TabIndex        =   24
      Top             =   6540
      Width           =   45
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   3240
      TabIndex        =   23
      Top             =   6180
      Width           =   45
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   3240
      TabIndex        =   22
      Top             =   5835
      Width           =   45
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   3240
      TabIndex        =   21
      Top             =   5520
      Width           =   45
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   8700
      TabIndex        =   20
      Top             =   4939
      Width           =   1245
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   6930
      TabIndex        =   19
      Top             =   4939
      Width           =   1245
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   5280
      TabIndex        =   18
      Top             =   4939
      Width           =   1245
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   3600
      TabIndex        =   17
      Top             =   4939
      Width           =   1245
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   1890
      TabIndex        =   16
      Top             =   4939
      Width           =   1245
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   8730
      TabIndex        =   10
      Top             =   3094
      Width           =   1245
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   6960
      TabIndex        =   12
      Top             =   3094
      Width           =   1245
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   5310
      TabIndex        =   13
      Top             =   3094
      Width           =   1245
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   3630
      TabIndex        =   14
      Top             =   3094
      Width           =   1245
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   15
      Top             =   876
      Width           =   6465
   End
End
Attribute VB_Name = "frmPanelAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Seleccionado As Byte

Private Sub Form_Load()

    Unload frmConnect

    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)

    Me.Picture = LoadPicture(Game.path(Interfaces) & "frmPanelAccount.jpg")

    Dim i As Long

    Me.Icon = frmMain.Icon

    For i = 1 To 10
        lblAccData(i).Caption = vbNullString
    Next i

    Me.lblAccData(0).Caption = AccountName

    'If Curper Then
        'Call FormParser.Parse_Form(Me)
    'End If

End Sub

Private Sub LoadTextsForm()
   Me.uAOBorrarPersonaje.Caption = JsonLanguage.item("FRMPANELACCOUNT_BTN_BORRAR_PERSONAJE").item("TEXTO")
   Me.uAOConectar.Caption = JsonLanguage.item("FRMPANELACCOUNT_BTN_CONECTAR").item("TEXTO")
   Me.uAOCrearPersonaje.Caption = JsonLanguage.item("FRMPANELACCOUNT_BTN_CREAR_PERSONAJE").item("TEXTO")
   Me.uAOSalir.Caption = JsonLanguage.item("FRMPANELACCOUNT_BTN_SALIR").item("TEXTO")
End Sub

Private Sub lblName_Click(Index As Integer)
    Seleccionado = Index
End Sub

Private Sub uAOBorrarPersonaje_Click()
    If LenB(lblAccData(Seleccionado).Caption) = 0 Then
        MsgBox JsonLanguage.item("ERROR_PERSONAJE_NO_SELECCIONADO").item("TEXTO")
        Exit Sub
    End If

   If MsgBox(JsonLanguage.item("FRMPANELACCOUNT_CONFIRMAR_BORRAR_PJ").item("TEXTO"), vbYesNo, JsonLanguage.item("FRMPANELACCOUNT_CONFIRMAR_BORRAR_PJ_TITULO").item("TEXTO")) = vbYes Then
    
      If Not frmMain.Client.State = sckConnected Then
         MsgBox JsonLanguage.item("ERROR_CONN_LOST").item("TEXTO")
         AccountName = vbNullString
         AccountHash = vbNullString
         NumberOfCharacters = 0
         Unload Me
      Else
         UserName = cPJ(Seleccionado).Nombre
         Call WriteDeleteChar
         
      End If

   End If
End Sub

Private Sub uAOConectar_Click()

    If LenB(lblAccData(Seleccionado).Caption) = 0 Then
        MsgBox JsonLanguage.item("ERROR_PERSONAJE_NO_SELECCIONADO").item("TEXTO")
        Exit Sub
    End If

    If Not frmMain.Client.State = sckConnected Then
        MsgBox JsonLanguage.item("ERROR_CONN_LOST").item("TEXTO")
        AccountName = vbNullString
        AccountHash = vbNullString
        NumberOfCharacters = 0
        Unload Me
    Else
        UserName = lblAccData(Seleccionado).Caption
        Call WriteLoginExistingChar
    End If

End Sub

Private Sub uAOCrearPersonaje_Click()

    If NumberOfCharacters > 9 Then
        MsgBox JsonLanguage.item("ERROR_DEMASIADOS_PJS").item("TEXTO")
        Exit Sub
    End If
    
    Dim LoopC As Long

    For LoopC = 1 To 10
        If LenB(lblAccData(LoopC).Caption) = 0 Then
            frmCrearPersonaje.Show
            Exit Sub
        End If
    Next LoopC

End Sub

Private Sub uAOSalir_Click()
    frmMain.Client.CloseSck
    Unload Me
    frmConnect.Show
End Sub

Private Sub picChar_Click(Index As Integer)

    Seleccionado = Index + 1
    
    If Seleccionado > NumberOfCharacters Then Exit Sub

    With cPJ(Seleccionado)

        If LenB(.Nombre) <> 0 Then
            lblCharData(0) = JsonLanguage.item("NOMBRE").item("TEXTO") & ": " & .Nombre
            lblCharData(1) = JsonLanguage.item("CLASE").item("TEXTO") & ": " & ListaClases(.Class)
            lblCharData(2) = JsonLanguage.item("RAZA").item("TEXTO") & ": " & ListaRazas(.Race)
            lblCharData(3) = JsonLanguage.item("NIVEL").item("TEXTO") & ": " & .Level
            lblCharData(4) = JsonLanguage.item("ORO").item("TEXTO") & ": " & .Gold
            lblCharData(5) = JsonLanguage.item("MAPA").item("TEXTO") & ": " & .Map
        Else
            lblCharData(0) = vbNullString
            lblCharData(1) = vbNullString
            lblCharData(2) = vbNullString
            lblCharData(3) = vbNullString
            lblCharData(4) = vbNullString
            lblCharData(5) = vbNullString
        End If

    End With

End Sub

Private Sub picChar_DblClick(Index As Integer)

    Seleccionado = Index + 1
    
    If LenB(lblAccData(Seleccionado).Caption) <> 0 Then
        UserName = lblAccData(Seleccionado).Caption
        Call WriteLoginExistingChar
    Else
        frmCrearPersonaje.Show
    End If

End Sub
