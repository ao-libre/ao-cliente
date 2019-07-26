VERSION 5.00
Begin VB.Form frmCambiaMotd 
   BorderStyle     =   0  'None
   Caption         =   """ZMOTD"""
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5175
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCambiaMotd.frx":0000
   ScaleHeight     =   361
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMotd 
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
      Height          =   2250
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   840
      Width           =   4290
   End
   Begin AOLibre.uAOButton imgAceptar 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   4800
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      TX              =   "Aceptar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":21188
      PICF            =   "frmCambiaMotd.frx":21BB2
      PICH            =   "frmCambiaMotd.frx":22874
      PICV            =   "frmCambiaMotd.frx":23806
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton imgAzul 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Azul"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":24708
      PICF            =   "frmCambiaMotd.frx":25132
      PICH            =   "frmCambiaMotd.frx":25DF4
      PICV            =   "frmCambiaMotd.frx":26D86
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
   Begin AOLibre.uAOButton imgRojo 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Rojo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":27C88
      PICF            =   "frmCambiaMotd.frx":286B2
      PICH            =   "frmCambiaMotd.frx":29374
      PICV            =   "frmCambiaMotd.frx":2A306
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
   Begin AOLibre.uAOButton imgBlanco 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Blanco"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":2B208
      PICF            =   "frmCambiaMotd.frx":2BC32
      PICH            =   "frmCambiaMotd.frx":2C8F4
      PICV            =   "frmCambiaMotd.frx":2D886
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
   Begin AOLibre.uAOButton imgGris 
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Gris"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":2E788
      PICF            =   "frmCambiaMotd.frx":2F1B2
      PICH            =   "frmCambiaMotd.frx":2FE74
      PICV            =   "frmCambiaMotd.frx":30E06
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
   Begin AOLibre.uAOButton imgMarron 
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Marron"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":31D08
      PICF            =   "frmCambiaMotd.frx":32732
      PICH            =   "frmCambiaMotd.frx":333F4
      PICV            =   "frmCambiaMotd.frx":34386
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
   Begin AOLibre.uAOButton imgVerde 
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Verde"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":35288
      PICF            =   "frmCambiaMotd.frx":35CB2
      PICH            =   "frmCambiaMotd.frx":36974
      PICV            =   "frmCambiaMotd.frx":37906
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
   Begin AOLibre.uAOButton imgMorado 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Morado"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":38808
      PICF            =   "frmCambiaMotd.frx":39232
      PICH            =   "frmCambiaMotd.frx":39EF4
      PICV            =   "frmCambiaMotd.frx":3AE86
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
   Begin AOLibre.uAOButton imgAmarillo 
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      TX              =   "Amarillo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCambiaMotd.frx":3BD88
      PICF            =   "frmCambiaMotd.frx":3C7B2
      PICH            =   "frmCambiaMotd.frx":3D474
      PICV            =   "frmCambiaMotd.frx":3E406
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
   Begin VB.Label lblNegrita 
      Caption         =   "Negrita"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblCursiva 
      Caption         =   "Cursiva"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   4320
      Width           =   855
   End
   Begin VB.Image imgOptCursiva 
      Height          =   255
      Index           =   1
      Left            =   3360
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "No olvides agregar los colores al final de cada linea (Ver tabla de abajo)"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image imgOptNegrita 
      Height          =   255
      Index           =   1
      Left            =   1440
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image imgOptCursiva 
      Height          =   195
      Index           =   0
      Left            =   3060
      Top             =   4380
      Width           =   180
   End
   Begin VB.Image imgOptNegrita 
      Height          =   195
      Index           =   0
      Left            =   1170
      Top             =   4380
      Width           =   180
   End
End
Attribute VB_Name = "frmCambiaMotd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmCambiarMotd.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonAzul As clsGraphicalButton
Private cBotonRojo As clsGraphicalButton
Private cBotonBlanco As clsGraphicalButton
Private cBotonGris As clsGraphicalButton
Private cBotonAmarillo As clsGraphicalButton
Private cBotonMorado As clsGraphicalButton
Private cBotonVerde As clsGraphicalButton
Private cBotonMarron As clsGraphicalButton
Private cBotonAceptar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private picNegrita As Picture
Private picCursiva As Picture

Private yNegrita As Byte
Private yCursiva As Byte

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGraficos & "VentanaCambioMOTD.jpg")
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonAzul = New clsGraphicalButton
    Set cBotonRojo = New clsGraphicalButton
    Set cBotonBlanco = New clsGraphicalButton
    Set cBotonGris = New clsGraphicalButton
    Set cBotonAmarillo = New clsGraphicalButton
    Set cBotonMorado = New clsGraphicalButton
    Set cBotonVerde = New clsGraphicalButton
    Set cBotonMarron = New clsGraphicalButton
    Set cBotonAceptar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
    Call cBotonAzul.Initialize(imgAzul, GrhPath & "BotonAzul.jpg", _
                                    GrhPath & "BotonAzulRollover.jpg", _
                                    GrhPath & "BotonAzulClick.jpg", Me)

    Call cBotonRojo.Initialize(imgRojo, GrhPath & "BotonRojo.jpg", _
                                    GrhPath & "BotonRojoRollover.jpg", _
                                    GrhPath & "BotonRojoClick.jpg", Me)

    Call cBotonBlanco.Initialize(imgBlanco, GrhPath & "BotonBlanco.jpg", _
                                    GrhPath & "BotonBlancoRollover.jpg", _
                                    GrhPath & "BotonBlancoClick.jpg", Me)

    Call cBotonGris.Initialize(imgGris, GrhPath & "BotonGris.jpg", _
                                    GrhPath & "BotonGrisRollover.jpg", _
                                    GrhPath & "BotonGrisClick.jpg", Me)
                                    
    Call cBotonAmarillo.Initialize(imgAmarillo, GrhPath & "BotonAmarillo.jpg", _
                                    GrhPath & "BotonAmarilloRollover.jpg", _
                                    GrhPath & "BotonAmarilloClick.jpg", Me)

    Call cBotonMorado.Initialize(imgMorado, GrhPath & "BotonMorado.jpg", _
                                    GrhPath & "BotonMoradoRollover.jpg", _
                                    GrhPath & "BotonMoradoClick.jpg", Me)

    Call cBotonVerde.Initialize(imgVerde, GrhPath & "BotonVerde.jpg", _
                                    GrhPath & "BotonVerdeRollover.jpg", _
                                    GrhPath & "BotonVerdeClick.jpg", Me)

    Call cBotonMarron.Initialize(imgMarron, GrhPath & "BotonMarron.jpg", _
                                    GrhPath & "BotonMarronRollover.jpg", _
                                    GrhPath & "BotonMarronClick.jpg", Me)

    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotonAceptarMotd.jpg", _
                                    GrhPath & "BotonAceptarRolloverMotd.jpg", _
                                    GrhPath & "BotonAceptarClickMotd.jpg", Me)
                                    
    Set picNegrita = LoadPicture(DirGraficos & "OpcionPrendidaN.jpg")
    Set picCursiva = LoadPicture(DirGraficos & "OpcionPrendidaC.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgAceptar_Click()
    Dim T() As String
    Dim i As Long, N As Long, Pos As Long
    Dim Upper_t As Long, Lower_t As Long
    
    If Len(txtMotd.Text) >= 2 Then
        If Right$(txtMotd.Text, 2) = vbNewLine Then txtMotd.Text = Left$(txtMotd.Text, Len(txtMotd.Text) - 2)
    End If
    
    T = Split(txtMotd.Text, vbNewLine)
    Lower_t = LBound(T)
    Upper_t = UBound(T)
    
    For i = Lower_t To Upper_t
        N = 0
        Pos = InStr(1, T(i), "~")
        Do While Pos > 0 And Pos < Len(T(i))
            N = N + 1
            Pos = InStr(Pos + 1, T(i), "~")
        Loop
        If N <> 5 Then
            MsgBox "Error en el formato de la linea " & i + 1 & "."
            Exit Sub
        End If
    Next i
    
    Call WriteSetMOTD(txtMotd.Text)
    Unload Me

End Sub

Private Sub imgAmarillo_Click()
    txtMotd.Text = txtMotd & "~244~244~0~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub imgAzul_Click()
    txtMotd.Text = txtMotd & "~50~70~250~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub imgBlanco_Click()
    txtMotd.Text = txtMotd & "~255~255~255~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub imgGris_Click()
    txtMotd.Text = txtMotd & "~157~157~157~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub imgMarron_Click()
    txtMotd.Text = txtMotd & "~97~58~31~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub imgMorado_Click()
    txtMotd.Text = txtMotd & "~128~0~128~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub imgOptCursiva_Click(index As Integer)
    
    If yCursiva = 0 Then
        imgOptCursiva(0).Picture = picCursiva
        yCursiva = 1
    Else
        Set imgOptCursiva(0).Picture = Nothing
        yCursiva = 0
    End If

End Sub

Private Sub imgOptCursiva_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgOptNegrita_Click(index As Integer)
    
    If yNegrita = 0 Then
        imgOptNegrita(0).Picture = picNegrita
        yNegrita = 1
    Else
        Set imgOptNegrita(0).Picture = Nothing
        yNegrita = 0
    End If
    
End Sub

Private Sub imgOptNegrita_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgRojo_Click()
    txtMotd.Text = txtMotd & "~255~0~0~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub imgTirarTodo_Click()

End Sub

Private Sub imgVerde_Click()
    txtMotd.Text = txtMotd & "~23~104~26~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub txtMotd_Change()

End Sub

Private Sub uAOButton1_Click()

End Sub
