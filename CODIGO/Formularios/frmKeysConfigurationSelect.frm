VERSION 5.00
Begin VB.Form frmKeysConfigurationSelect 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Configuracion Controles / Config Keys"
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin AOLibre.uAOButton btnNormalKeys 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   7560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      TX              =   "Normal"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmKeysConfigurationSelect.frx":0000
      PICF            =   "frmKeysConfigurationSelect.frx":001C
      PICH            =   "frmKeysConfigurationSelect.frx":0038
      PICV            =   "frmKeysConfigurationSelect.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton btnAlternativeKeys 
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   7560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      TX              =   "Alternative"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmKeysConfigurationSelect.frx":0070
      PICF            =   "frmKeysConfigurationSelect.frx":008C
      PICH            =   "frmKeysConfigurationSelect.frx":00A8
      PICV            =   "frmKeysConfigurationSelect.frx":00C4
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblAlternativeTitle 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo Alternativo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   1215
      Left            =   5640
      TabIndex        =   6
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label lblNormalTitle 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo Normal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   1575
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      X1              =   4920
      X2              =   4920
      Y1              =   1800
      Y2              =   8040
   End
   Begin VB.Label lblAlternativeText 
      BackStyle       =   0  'Transparent
      Caption         =   "Legacy (Directional Arrows) "
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   5520
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label lblNormalText 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Image imgAlternativeKeyboard 
      Height          =   1335
      Left            =   5400
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Image imgNormalKeyboard 
      Height          =   1335
      Left            =   240
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmKeysConfigurationSelect.frx":00E0
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmKeysConfigurationSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Call LoadTextsForm
   Call LoadAOCustomControlsPictures(Me)
   imgAlternativeKeyboard.Picture = LoadPicture(Game.path(Interfaces) & "frmKeysConfigurationSelectAlternativeKeyboard.jpg")
   imgNormalKeyboard.Picture = LoadPicture(Game.path(Interfaces) & "frmKeysConfigurationSelectNormalKeyboard.jpg")
End Sub

Private Sub LoadTextsForm()
   lblAlternativeText.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_LBL_ALTERNATIVE_TEXT").item("TEXTO")
   lblAlternativeTitle.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_LBL_ALTERNATIVE_TITLE").item("TEXTO")
   lblNormalText.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_LBL_NORMAL_TEXT").item("TEXTO")
   lblNormalTitle.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_LBL_NORMAL_TITLE").item("TEXTO")
   lblTitle.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_LBL_TITLE").item("TEXTO")
   btnNormalKeys.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_BTN_NORMAL_KEYS").item("TEXTO")
   btnAlternativeKeys.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_BTN_ALTERNATIVE_KEYS").item("TEXTO")
End Sub

Private Sub btnAlternativeKeys_Click()
   CustomKeys.SetKeyConfigFileInUse("Alternative")
   SetFalseMostrarBindKeysSelection
   Unload Me
End Sub

Private Sub btnNormalKeys_Click()
   CustomKeys.SetKeyConfigFileInUse("Normal")
   SetFalseMostrarBindKeysSelection
   Unload Me
End Sub

Private Sub SetFalseMostrarBindKeysSelection()
   ClientSetup.MostrarBindKeysSelection = False
   Call WriteVar(Game.path(INIT) & "Config.ini", "OTHER", "MOSTRAR_BIND_KEYS_SELECTION", "False")
End Sub

