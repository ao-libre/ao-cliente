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
      Left            =   435
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   795
      Width           =   4290
   End
   Begin VB.Image imgOptCursiva 
      Height          =   255
      Index           =   1
      Left            =   3360
      Top             =   4320
      Width           =   615
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
   Begin VB.Image imgAceptar 
      Height          =   375
      Left            =   480
      Top             =   4800
      Width           =   4350
   End
   Begin VB.Image imgMarron 
      Height          =   375
      Left            =   3720
      Top             =   3720
      Width           =   855
   End
   Begin VB.Image imgVerde 
      Height          =   375
      Left            =   2640
      Top             =   3720
      Width           =   855
   End
   Begin VB.Image imgMorado 
      Height          =   375
      Left            =   1560
      Top             =   3720
      Width           =   855
   End
   Begin VB.Image imgAmarillo 
      Height          =   375
      Left            =   480
      Top             =   3720
      Width           =   855
   End
   Begin VB.Image imgGris 
      Height          =   375
      Left            =   3720
      Top             =   3240
      Width           =   855
   End
   Begin VB.Image imgBlanco 
      Height          =   375
      Left            =   2640
      Top             =   3240
      Width           =   855
   End
   Begin VB.Image imgRojo 
      Height          =   375
      Left            =   1560
      Top             =   3240
      Width           =   855
   End
   Begin VB.Image imgAzul 
      Height          =   375
      Left            =   480
      Top             =   3240
      Width           =   855
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

Private clsFormulario    As clsFormMovementManager

Private cBotonAzul       As clsGraphicalButton
Private cBotonRojo       As clsGraphicalButton
Private cBotonBlanco     As clsGraphicalButton
Private cBotonGris       As clsGraphicalButton
Private cBotonAmarillo   As clsGraphicalButton
Private cBotonMorado     As clsGraphicalButton
Private cBotonVerde      As clsGraphicalButton
Private cBotonMarron     As clsGraphicalButton
Private cBotonAceptar    As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private picNegrita       As Picture
Private picCursiva       As Picture

Private yNegrita         As Byte
Private yCursiva         As Byte

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    
    On Error GoTo Form_Load_Err
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGraficos & "VentanaCambioMOTD.jpg")
    
    Call LoadButtons

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
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
    
    Call cBotonAzul.Initialize(imgAzul, GrhPath & "BotonAzul.jpg", GrhPath & "BotonAzulRollover.jpg", GrhPath & "BotonAzulClick.jpg", Me)

    Call cBotonRojo.Initialize(imgRojo, GrhPath & "BotonRojo.jpg", GrhPath & "BotonRojoRollover.jpg", GrhPath & "BotonRojoClick.jpg", Me)

    Call cBotonBlanco.Initialize(imgBlanco, GrhPath & "BotonBlanco.jpg", GrhPath & "BotonBlancoRollover.jpg", GrhPath & "BotonBlancoClick.jpg", Me)

    Call cBotonGris.Initialize(imgGris, GrhPath & "BotonGris.jpg", GrhPath & "BotonGrisRollover.jpg", GrhPath & "BotonGrisClick.jpg", Me)
                                    
    Call cBotonAmarillo.Initialize(imgAmarillo, GrhPath & "BotonAmarillo.jpg", GrhPath & "BotonAmarilloRollover.jpg", GrhPath & "BotonAmarilloClick.jpg", Me)

    Call cBotonMorado.Initialize(imgMorado, GrhPath & "BotonMorado.jpg", GrhPath & "BotonMoradoRollover.jpg", GrhPath & "BotonMoradoClick.jpg", Me)

    Call cBotonVerde.Initialize(imgVerde, GrhPath & "BotonVerde.jpg", GrhPath & "BotonVerdeRollover.jpg", GrhPath & "BotonVerdeClick.jpg", Me)

    Call cBotonMarron.Initialize(imgMarron, GrhPath & "BotonMarron.jpg", GrhPath & "BotonMarronRollover.jpg", GrhPath & "BotonMarronClick.jpg", Me)

    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotonAceptarMotd.jpg", GrhPath & "BotonAceptarRolloverMotd.jpg", GrhPath & "BotonAceptarClickMotd.jpg", Me)
                                    
    Set picNegrita = LoadPicture(DirGraficos & "OpcionPrendidaN.jpg")
    Set picCursiva = LoadPicture(DirGraficos & "OpcionPrendidaC.jpg")

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgAceptar_Click()
    
    On Error GoTo imgAceptar_Click_Err
    
    Dim T() As String, Upper_t As Long, Lower_t As Long, Len_t As Long
    Dim i   As Long, N As Long, Pos As Long
    
    If Len(txtMotd.Text) >= 2 Then
        If Right$(txtMotd.Text, 2) = vbNewLine Then txtMotd.Text = Left$(txtMotd.Text, Len(txtMotd.Text) - 2)

    End If
    
    T = Split(txtMotd.Text, vbNewLine)
    Lower_t = LBound(T)
    Upper_t = UBound(T)
    Len_t = Len(T(i))
    
    For i = Lower_t To Upper_t
        N = 0
        Pos = InStr(1, T(i), "~")

        Do While Pos > 0 And Pos < Len_t
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

    
    Exit Sub

imgAceptar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgAceptar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgAmarillo_Click()
    
    On Error GoTo imgAmarillo_Click_Err
    
    txtMotd.Text = txtMotd & "~244~244~0~" & CStr(yNegrita) & "~" & CStr(yCursiva)

    
    Exit Sub

imgAmarillo_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgAmarillo_Click"
    End If
Resume Next
    
End Sub

Private Sub imgAzul_Click()
    
    On Error GoTo imgAzul_Click_Err
    
    txtMotd.Text = txtMotd & "~50~70~250~" & CStr(yNegrita) & "~" & CStr(yCursiva)

    
    Exit Sub

imgAzul_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgAzul_Click"
    End If
Resume Next
    
End Sub

Private Sub imgBlanco_Click()
    
    On Error GoTo imgBlanco_Click_Err
    
    txtMotd.Text = txtMotd & "~255~255~255~" & CStr(yNegrita) & "~" & CStr(yCursiva)

    
    Exit Sub

imgBlanco_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgBlanco_Click"
    End If
Resume Next
    
End Sub

Private Sub imgGris_Click()
    
    On Error GoTo imgGris_Click_Err
    
    txtMotd.Text = txtMotd & "~157~157~157~" & CStr(yNegrita) & "~" & CStr(yCursiva)

    
    Exit Sub

imgGris_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgGris_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMarron_Click()
    
    On Error GoTo imgMarron_Click_Err
    
    txtMotd.Text = txtMotd & "~97~58~31~" & CStr(yNegrita) & "~" & CStr(yCursiva)

    
    Exit Sub

imgMarron_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgMarron_Click"
    End If
Resume Next
    
End Sub

Private Sub imgMorado_Click()
    
    On Error GoTo imgMorado_Click_Err
    
    txtMotd.Text = txtMotd & "~128~0~128~" & CStr(yNegrita) & "~" & CStr(yCursiva)

    
    Exit Sub

imgMorado_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgMorado_Click"
    End If
Resume Next
    
End Sub

Private Sub imgOptCursiva_Click(Index As Integer)
    
    On Error GoTo imgOptCursiva_Click_Err
    
    
    If yCursiva = 0 Then
        imgOptCursiva(0).Picture = picCursiva
        yCursiva = 1
    Else
        Set imgOptCursiva(0).Picture = Nothing
        yCursiva = 0

    End If

    
    Exit Sub

imgOptCursiva_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgOptCursiva_Click"
    End If
Resume Next
    
End Sub

Private Sub imgOptCursiva_MouseMove(Index As Integer, _
                                    Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)
    
    On Error GoTo imgOptCursiva_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

imgOptCursiva_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgOptCursiva_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgOptNegrita_Click(Index As Integer)
    
    On Error GoTo imgOptNegrita_Click_Err
    
    
    If yNegrita = 0 Then
        imgOptNegrita(0).Picture = picNegrita
        yNegrita = 1
    Else
        Set imgOptNegrita(0).Picture = Nothing
        yNegrita = 0

    End If
    
    
    Exit Sub

imgOptNegrita_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgOptNegrita_Click"
    End If
Resume Next
    
End Sub

Private Sub imgOptNegrita_MouseMove(Index As Integer, _
                                    Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)
    
    On Error GoTo imgOptNegrita_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

imgOptNegrita_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgOptNegrita_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgRojo_Click()
    
    On Error GoTo imgRojo_Click_Err
    
    txtMotd.Text = txtMotd & "~255~0~0~" & CStr(yNegrita) & "~" & CStr(yCursiva)

    
    Exit Sub

imgRojo_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgRojo_Click"
    End If
Resume Next
    
End Sub

Private Sub imgVerde_Click()
    
    On Error GoTo imgVerde_Click_Err
    
    txtMotd.Text = txtMotd & "~23~104~26~" & CStr(yNegrita) & "~" & CStr(yCursiva)

    
    Exit Sub

imgVerde_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCambiaMotd" & "->" & "imgVerde_Click"
    End If
Resume Next
    
End Sub

