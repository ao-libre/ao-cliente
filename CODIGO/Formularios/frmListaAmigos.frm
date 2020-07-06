VERSION 5.00
Begin VB.Form frmAmigos 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Amigos"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346.011
   ScaleMode       =   0  'User
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListAmigos 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
   End
   Begin AOLibre.uAOButton AgregarAmigo 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Agregar Amigo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmListaAmigos.frx":0000
      PICF            =   "frmListaAmigos.frx":001C
      PICH            =   "frmListaAmigos.frx":0038
      PICV            =   "frmListaAmigos.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton BorrarAmigo 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Borrar Amigo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmListaAmigos.frx":0070
      PICF            =   "frmListaAmigos.frx":008C
      PICH            =   "frmListaAmigos.frx":00A8
      PICV            =   "frmListaAmigos.frx":00C4
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton btnSalir 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   4560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmListaAmigos.frx":00E0
      PICF            =   "frmListaAmigos.frx":00FC
      PICH            =   "frmListaAmigos.frx":0118
      PICV            =   "frmListaAmigos.frx":0134
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Aqu� puedes ver tu lista de amigos, as� como agregar o eliminar usuarios de la lista."
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmAmigos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ActualizarLista()
' @@ Cuicui
    Dim i As Long
    
    ListAmigos.Clear
    
    For i = 1 To MAX_AMIGOS
        If LenB(amigos(i)) > 0 Then
            Call ListAmigos.AddItem(amigos(i))
        Else
            Exit For
        End If
    Next i
    
End Sub

Private Sub Form_Load()
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    Call clsFormulario.Initialize(Me)
    
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaListaAmigos.jpg")
    
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)

    Set picNegrita = LoadPicture(Game.path(Interfaces) & "OpcionPrendidaN.jpg")
    Set picCursiva = LoadPicture(Game.path(Interfaces) & "OpcionPrendidaC.jpg")

End Sub

Private Sub LoadTextsForm()
    Me.lblTitle.Caption = JsonLanguage.item("FRM_LISTAAMIGOS_TITLE").item("TEXTO")
    Me.AgregarAmigo.Caption = JsonLanguage.item("FRM_LISTAAMIGOS_AGREGAR").item("TEXTO")
    Me.BorrarAmigo.Caption = JsonLanguage.item("FRM_LISTAAMIGOS_BORRAR").item("TEXTO")
    Me.btnSalir.Caption = JsonLanguage.item("FRM_LISTAAMIGOS_SALIR").item("TEXTO")
End Sub

Private Sub AgregarAmigo_Click()

    Dim SendName As String
        SendName = InputBox("Escriba el nombre del usuario a agregar.", "Agregar Amigo")

    If LenB(Trim$(SendName)) Then
        
        If MsgBox("Seguro desea agregar a " & SendName & "?", vbYesNo, "Agregar Amigo") = vbYes Then
           Call WriteAddAmigo(SendName, 1)
        End If
        
    Else

        With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
            Call ShowConsoleMsg("Nombre Invalido", .Red, .Green, .Blue, .bold, .italic)
        End With

    End If

End Sub

Private Sub BorrarAmigo_Click()

    If LenB(ListAmigos.List(ListAmigos.ListIndex)) = 0 Then Exit Sub
    
    If MsgBox("Seguro desea borrar a " & ListAmigos.List(ListAmigos.ListIndex) & "?", vbYesNo, "Borrar Amigo") = vbYes Then
        Call WriteDelAmigo(ListAmigos.ListIndex + 1)
    End If
    
End Sub


Private Sub btnSalir_Click()
    Unload Me
End Sub
