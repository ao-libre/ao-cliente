VERSION 5.00
Begin VB.Form frmQuestInfo 
   BorderStyle     =   0  'None
   Caption         =   "Informacion de la mision"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   6525
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQuestInfo.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   4695
   End
   Begin AOLibre.uAOButton Aceptar 
      Height          =   615
      Left            =   840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Aceptar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmQuestInfo.frx":1E9ED
      PICF            =   "frmQuestInfo.frx":1F417
      PICH            =   "frmQuestInfo.frx":200D9
      PICV            =   "frmQuestInfo.frx":2106B
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
   Begin AOLibre.uAOButton Rechazar 
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      TX              =   "Rechazar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmQuestInfo.frx":21F6D
      PICF            =   "frmQuestInfo.frx":22997
      PICH            =   "frmQuestInfo.frx":23659
      PICV            =   "frmQuestInfo.frx":245EB
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
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Informacion de la mision:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   375
      Width           =   3615
   End
End
Attribute VB_Name = "frmQuestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaQuest.jpg")
    
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)

End Sub

Private Sub LoadTextsForm()
    Me.lblDesc.Caption = JsonLanguage.item("FRM_QUEST_DESC").item("TEXTO")
    Me.Aceptar.Caption = JsonLanguage.item("FRM_QUEST_ACCEPT").item("TEXTO")
    Me.Rechazar.Caption = JsonLanguage.item("FRM_QUEST_EXIT").item("TEXTO")
End Sub

Private Sub Aceptar_Click()
    Call WriteQuestAccept
    Unload Me
End Sub

Private Sub Rechazar_Click()
    Unload Me
End Sub
