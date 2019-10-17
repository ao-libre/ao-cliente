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
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Caption         =   "Negrita"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   855
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCursiva 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Caption         =   "Cursiva"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   4320
      Width           =   855
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgOptCursiva 
      Height          =   255
      Index           =   1
      Left            =   3360
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Caption         =   "No olvides agregar los colores al final de cada linea (Ver tabla de abajo)"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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

Private picNegrita As Picture
Private picCursiva As Picture

Private yNegrita As Byte
Private yCursiva As Byte

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaCambioMOTD.jpg")
    
    Call LoadTextsForm
    Call LoadButtonsPictures

    Set picNegrita = LoadPicture(Game.path(Interfaces) & "OpcionPrendidaN.jpg")
    Set picCursiva = LoadPicture(Game.path(Interfaces) & "OpcionPrendidaC.jpg")
End Sub

Private Sub LoadTextsForm()
    Me.lblTitle.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_TITLE").Item("TEXTO")
    Me.imgAzul.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_AZUL").Item("TEXTO")
    Me.imgRojo.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_ROJO").Item("TEXTO")
    Me.imgBlanco.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_BLANCO").Item("TEXTO")
    Me.imgGris.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_GRIS").Item("TEXTO")
    Me.imgAmarillo.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_AMARILLO").Item("TEXTO")
    Me.imgMorado.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_MORADO").Item("TEXTO")
    Me.imgVerde.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_VERDE").Item("TEXTO")
    Me.imgMarron.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_MARRON").Item("TEXTO")
    Me.imgAceptar.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_ACEPTAR").Item("TEXTO")
    Me.lblCursiva.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_CURSIVA").Item("TEXTO")
    Me.lblNegrita.Caption = JsonLanguage.Item("FRM_CAMBIAMOTD_NEGRITA").Item("TEXTO")
End Sub

Private Sub LoadButtonsPictures()
    Dim DirButtons As String
    DirButtons = Game.path(Graficos) & "\Botones\"

    Dim cControl As Control
    For Each cControl In Me.Controls
        If TypeOf cControl Is uAOButton Then
            cControl.PictureEsquina = LoadPicture(ImgRequest(DirButtons & uAOButton_bEsquina))
            cControl.PictureFondo = LoadPicture(ImgRequest(DirButtons & uAOButton_bFondo))
            cControl.PictureHorizontal = LoadPicture(ImgRequest(DirButtons & uAOButton_bHorizontal))
            cControl.PictureVertical = LoadPicture(ImgRequest(DirButtons & uAOButton_bVertical))
        ElseIf TypeOf cControl Is uAOCheckbox Then
            cControl.Picture = LoadPicture(ImgRequest(DirButtons & uAOButton_cCheckboxSmall))
        End If
    Next
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

Private Sub imgOptCursiva_Click(Index As Integer)
    
    If yCursiva = 0 Then
        imgOptCursiva(0).Picture = picCursiva
        yCursiva = 1
    Else
        Set imgOptCursiva(0).Picture = Nothing
        yCursiva = 0
    End If

End Sub

Private Sub imgOptNegrita_Click(Index As Integer)
    
    If yNegrita = 0 Then
        imgOptNegrita(0).Picture = picNegrita
        yNegrita = 1
    Else
        Set imgOptNegrita(0).Picture = Nothing
        yNegrita = 0
    End If
    
End Sub

Private Sub imgRojo_Click()
    txtMotd.Text = txtMotd & "~255~0~0~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub

Private Sub imgVerde_Click()
    txtMotd.Text = txtMotd & "~23~104~26~" & CStr(yNegrita) & "~" & CStr(yCursiva)
End Sub
