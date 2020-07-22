VERSION 5.00
Begin VB.Form FrmRetos 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Retos"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9570
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmRetos.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCompa 
      Alignment       =   2  'Center
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
      Height          =   285
      Index           =   1
      Left            =   6240
      TabIndex        =   11
      Top             =   2810
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtCompa 
      Alignment       =   2  'Center
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
      Height          =   285
      Index           =   0
      Left            =   6240
      TabIndex        =   10
      Top             =   2450
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtOponente 
      Alignment       =   2  'Center
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
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtOponente 
      Alignment       =   2  'Center
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
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtGld 
      Alignment       =   2  'Center
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
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Text            =   "0"
      Top             =   4290
      Width           =   1695
   End
   Begin VB.TextBox txtOponente 
      Alignment       =   2  'Center
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
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin AOLibre.uAOButton Comenzar 
      Height          =   615
      Left            =   5400
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Comenzar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "FrmRetos.frx":27B23
      PICF            =   "FrmRetos.frx":2854D
      PICH            =   "FrmRetos.frx":2920F
      PICV            =   "FrmRetos.frx":2A1A1
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton Salir 
      Height          =   615
      Left            =   7200
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4320
      Width           =   975
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "FrmRetos.frx":2B0A3
      PICF            =   "FrmRetos.frx":2BACD
      PICH            =   "FrmRetos.frx":2C78F
      PICV            =   "FrmRetos.frx":2D721
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label tresvstres 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6600
      TabIndex        =   17
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label dosvsdos 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4080
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label unovsuno 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1680
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona el tipo de reto que quieres jugar."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   420
      Width           =   5400
   End
   Begin VB.Label lblCompa2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aliado 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   4920
      TabIndex        =   5
      Top             =   2810
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblOponente3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oponente 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   1440
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblCompa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aliado 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   4920
      TabIndex        =   3
      Top             =   2450
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblOponente2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oponente 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   1440
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblOro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monedas de Oro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   1575
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblOponente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oponente 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   1440
      TabIndex        =   0
      Top             =   2280
      Width           =   1125
   End
End
Attribute VB_Name = "FrmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RetoModo As Byte

Private Sub dosvsdos_Click()

    txtOponente(1).Visible = True
    txtOponente(2).Visible = False
    lblOponente.Visible = True
    lblOponente2.Visible = True
    lblOponente3.Visible = False
    txtCompa(0).Visible = True
    txtCompa(1).Visible = False
    lblCompa.Visible = True
    lblCompa2.Visible = False
    RetoModo = 2
    
End Sub

Private Sub Form_Load()

    RetoModo = 1
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    Call clsFormulario.Initialize(Me)
    
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaRetos.jpg")
    
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)

    Set picNegrita = LoadPicture(Game.path(Interfaces) & "OpcionPrendidaN.jpg")
    Set picCursiva = LoadPicture(Game.path(Interfaces) & "OpcionPrendidaC.jpg")

End Sub

Private Sub LoadTextsForm()
    Me.lblDesc.Caption = JsonLanguage.item("FRM_RETOS_DESC").item("TEXTO")
    Me.lblOponente.Caption = JsonLanguage.item("FRM_RETOS_LBLOP").item("TEXTO")
    Me.lblOponente2.Caption = JsonLanguage.item("FRM_RETOS_LBLOP2").item("TEXTO")
    Me.lblOponente3.Caption = JsonLanguage.item("FRM_RETOS_LBLOP3").item("TEXTO")
    Me.lblCompa.Caption = JsonLanguage.item("FRM_RETOS_COMPA").item("TEXTO")
    Me.lblCompa2.Caption = JsonLanguage.item("FRM_RETOS_COMPA2").item("TEXTO")
    Me.Comenzar.Caption = JsonLanguage.item("FRM_RETOS_START").item("TEXTO")
    Me.Salir.Caption = JsonLanguage.item("FRM_RETOS_EXIT").item("TEXTO")
    Me.lblOro.Caption = JsonLanguage.item("FRM_RETOS_LBL_ORO").item("TEXTO")
    
End Sub

Private Sub Comenzar_Click()

    Dim ErrorMsg As String
    Dim ListUser As String
    
        If Not CheckDataReto(RetoModo, ListUser, ErrorMsg) Then
                MsgBox ErrorMsg
                Exit Sub
        End If
            
        Call Protocol.WriteFightSend(ListUser, Val(txtGld.Text))
        Unload Me
        
End Sub

Private Function CheckDataReto(ByVal Selected As Byte, _
                                ByRef ListUser As String, _
                                ByRef ErrorMsg As String) As Boolean
    CheckDataReto = False
    
    Dim a As Long
    
    If Val(txtGld.Text) < 0 Then
        ErrorMsg = "La apuesta minima es por 0 monedas de oro"
        Exit Function
    End If
    
    If Len(txtOponente(0).Text) <= 0 Then
        ErrorMsg = "Debes seleccionar al oponente nro 1"
        Exit Function
    End If
    
    ListUser = txtOponente(0).Text
    
    Select Case Selected
        Case 2
            If Len(txtOponente(1).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar al oponente nro 2"
                Exit Function
            End If
            
            If Len(txtCompa(0).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar a tu aliado"
                Exit Function
            End If
            
            ListUser = txtOponente(0).Text & "-" & txtOponente(1).Text & "-" & txtCompa(0).Text
        Case 3
            If Len(txtOponente(1).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar al oponente nro 2"
                Exit Function
            End If
            
            If Len(txtOponente(2).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar al oponente nro 3"
                Exit Function
            End If
            
            If Len(txtCompa(0).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar a tu aliado nro 2"
                Exit Function
            End If
            
            If Len(txtCompa(1).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar a tu aliado nro 3"
                Exit Function
            End If
            
            ListUser = txtOponente(0).Text & "-" & txtOponente(1) & "-" & txtOponente(2) & "-" & txtCompa(0).Text & "-" & txtCompa(1).Text
    End Select
    
    
    CheckDataReto = True
End Function

Private Sub Salir_Click()
    Unload Me
End Sub

Private Sub tresvstres_Click()

    txtOponente(1).Visible = True
    txtOponente(2).Visible = True
    lblOponente.Visible = True
    lblOponente2.Visible = True
    lblOponente3.Visible = True
    txtCompa(0).Visible = True
    txtCompa(1).Visible = True
    lblCompa.Visible = True
    lblCompa2.Visible = True
    RetoModo = 3
    
End Sub

Private Sub unovsuno_Click()

    txtOponente(1).Visible = False
    txtOponente(2).Visible = False
    lblOponente.Visible = True
    lblOponente2.Visible = False
    lblOponente3.Visible = False
    txtCompa(0).Visible = False
    txtCompa(1).Visible = False
    lblCompa.Visible = False
    lblCompa2.Visible = False
    RetoModo = 1
    
End Sub
