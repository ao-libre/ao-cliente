VERSION 5.00
Begin VB.Form frmCerrar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cerrar"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmCerrar.frx":0000
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin AOLibre.uAOButton cRegresar 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      TX              =   "Regresar Pantalla de Inicio"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCerrar.frx":9B51
      PICF            =   "frmCerrar.frx":A57B
      PICH            =   "frmCerrar.frx":B23D
      PICV            =   "frmCerrar.frx":C1CF
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton cSalir 
      CausesValidation=   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      TX              =   "Salir del Juego"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCerrar.frx":D0D1
      PICF            =   "frmCerrar.frx":DAFB
      PICH            =   "frmCerrar.frx":E7BD
      PICV            =   "frmCerrar.frx":F74F
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton cCancelQuit 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      TX              =   "Cancelar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCerrar.frx":10651
      PICF            =   "frmCerrar.frx":1107B
      PICH            =   "frmCerrar.frx":11D3D
      PICV            =   "frmCerrar.frx":12CCF
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCerrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub cCancelQuit_Click()
    Call Audio.PlayWave(SND_CLICK)
    Set clsFormulario = Nothing
    Unload Me
End Sub

Private Sub cRegresar_Click()
    Call Audio.PlayWave(SND_CLICK)
    
    Set clsFormulario = Nothing
    
    If UserParalizado Then 'Inmo

        With FontTypes(FontTypeNames.FONTTYPE_WARNING)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NO_SALIR").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
        End With
        
        Exit Sub
        
    End If
    
    ' Desactivamos los macros.
    If frmMain.trainingMacro.Enabled Then Call frmMain.DesactivarMacroHechizos
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    
    ' Nos desconectamos y lo mando al Panel de la Cuenta
    Call WriteQuit
    
    Call Unload(Me)
    
End Sub

Private Sub cSalir_Click()
    Call Audio.PlayWave(SND_CLICK)
    Set clsFormulario = Nothing
    Call CloseClient
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Set clsFormulario = Nothing
        Call Unload(Me)
    End If
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    Call clsFormulario.Initialize(Me)
    
    Me.Picture = LoadPicture(Game.path(Interfaces) & "frmCerrar.jpg")
    'Call LoadAOCustomControlsPictures(Me)
    'Todo: Poner la carga de botones como en el frmCambiaMotd.frm para mantener coherencia con el resto de la aplicacion
    'y poder borrar los frx de este archivo

    Call LoadFormTexts
End Sub

Private Sub LoadFormTexts()
    cRegresar.Caption = JsonLanguage.item("CERRAR").item("TEXTOS").item(1)
    cSalir.Caption = JsonLanguage.item("CERRAR").item("TEXTOS").item(2)
    cCancelQuit.Caption = JsonLanguage.item("CERRAR").item("TEXTOS").item(3)
End Sub

