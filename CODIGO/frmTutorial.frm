VERSION 5.00
Begin VB.Form frmTutorial 
   BorderStyle     =   0  'None
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmTutorial.frx":0000
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   390
      Left            =   525
      TabIndex        =   4
      Top             =   435
      Width           =   7725
   End
   Begin VB.Image imgCheck 
      Height          =   450
      Left            =   3060
      Picture         =   "frmTutorial.frx":1B45E
      Top             =   6900
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgMostrar 
      Height          =   570
      Left            =   3000
      Picture         =   "frmTutorial.frx":1BEF0
      Top             =   6855
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image imgSiguiente 
      Height          =   360
      Left            =   6840
      Picture         =   "frmTutorial.frx":20A9A
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Image imgAnterior 
      Height          =   360
      Left            =   480
      Picture         =   "frmTutorial.frx":26F77
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblMensaje 
      BackStyle       =   0  'Transparent
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
      Height          =   5790
      Left            =   525
      TabIndex        =   3
      Top             =   840
      Width           =   7725
   End
   Begin VB.Label lblPagTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   7365
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblPagActual 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   6870
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8430
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   75
      Width           =   255
   End
End
Attribute VB_Name = "frmTutorial"
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

Option Explicit

Private clsFormulario    As clsFormMovementManager

Private cBotonSiguiente  As clsGraphicalButton
Private cBotonAnterior   As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Type tTutorial

    sTitle As String
    sPage As String

End Type

Private picCheck    As Picture
Private picMostrar  As Picture

Private Tutorial()  As tTutorial
Private NumPages    As Long
Private CurrentPage As Long

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(DirGraficos & "VentanaTutorial.jpg")
    
    Call LoadButtons
    
    Call LoadTutorial
    
    CurrentPage = 1
    Call SelectPage(CurrentPage)

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmTutorial" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonSiguiente = New clsGraphicalButton
    Set cBotonAnterior = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonSiguiente.Initialize(imgSiguiente, GrhPath & "BotonSiguienteTutorial.jpg", GrhPath & "BotonSiguienteRolloverTutorial.jpg", GrhPath & "BotonSiguienteClickTutorial.jpg", Me, GrhPath & "BotonSiguienteGris.jpg")

    Call cBotonAnterior.Initialize(imgAnterior, GrhPath & "BotonAnteriorTutorial.jpg", GrhPath & "BotonAnteriorRolloverTutorial.jpg", GrhPath & "BotonAnteriorClickTutorial.jpg", Me, GrhPath & "BotonAnteriorGris.jpg", True)
                                    
    Set picCheck = LoadPicture(GrhPath & "CheckTutorial.bmp")
    Set picMostrar = LoadPicture(GrhPath & "NoMostrarTutorial.bmp")
    
    imgMostrar.Picture = picMostrar
    
    If Not bShowTutorial Then
        imgCheck.Picture = picCheck
    Else
        Set imgCheck.Picture = Nothing

    End If
    
    lblCerrar.MouseIcon = picMouseIcon

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmTutorial" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmTutorial" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgAnterior_Click()
    
    On Error GoTo imgAnterior_Click_Err
    

    If Not cBotonAnterior.IsEnabled Then Exit Sub
    
    CurrentPage = CurrentPage - 1
    
    If CurrentPage = 1 Then Call cBotonAnterior.EnableButton(False)
    
    If Not cBotonSiguiente.IsEnabled Then Call cBotonSiguiente.EnableButton(True)
    
    Call SelectPage(CurrentPage)

    
    Exit Sub

imgAnterior_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmTutorial" & "->" & "imgAnterior_Click"
    End If
Resume Next
    
End Sub

Private Sub imgCheck_Click()
    
    On Error GoTo imgCheck_Click_Err
    
    
    bShowTutorial = Not bShowTutorial
    
    If Not bShowTutorial Then
        imgCheck.Picture = picCheck
    Else
        Set imgCheck.Picture = Nothing

    End If

    
    Exit Sub

imgCheck_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmTutorial" & "->" & "imgCheck_Click"
    End If
Resume Next
    
End Sub

Private Sub imgSiguiente_Click()
    
    On Error GoTo imgSiguiente_Click_Err
    
    
    If Not cBotonSiguiente.IsEnabled Then Exit Sub
    
    CurrentPage = CurrentPage + 1
    
    ' DEshabilita el boton siguiente si esta en la ultima pagina
    If CurrentPage = NumPages Then Call cBotonSiguiente.EnableButton(False)
    
    ' Habilita el boton anterior
    If Not cBotonAnterior.IsEnabled Then Call cBotonAnterior.EnableButton(True)
    
    Call SelectPage(CurrentPage)

    
    Exit Sub

imgSiguiente_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmTutorial" & "->" & "imgSiguiente_Click"
    End If
Resume Next
    
End Sub

Private Sub lblCerrar_Click()
    
    On Error GoTo lblCerrar_Click_Err
    
    bShowTutorial = False 'Mientras no se pueda tildar/destildar para verlo más tarde, esto queda así :P
    Unload Me

    
    Exit Sub

lblCerrar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmTutorial" & "->" & "lblCerrar_Click"
    End If
Resume Next
    
End Sub

Private Sub LoadTutorial()
    
    On Error GoTo LoadTutorial_Err
    
    
    Dim TutorialPath As String
    Dim lPage        As Long
    Dim NumLines     As Long
    Dim lLine        As Long
    Dim sLine        As String
    
    TutorialPath = DirExtras & "Tutorial.dat"
    NumPages = Val(GetVar(TutorialPath, "INIT", "NumPags"))
    
    If NumPages > 0 Then
        ReDim Tutorial(1 To NumPages)
        
        ' Cargo paginas
        For lPage = 1 To NumPages
            NumLines = Val(GetVar(TutorialPath, "PAG" & lPage, "NumLines"))
            
            With Tutorial(lPage)
                
                .sTitle = GetVar(TutorialPath, "PAG" & lPage, "Title")
                
                ' Cargo cada linea de la pagina
                For lLine = 1 To NumLines
                    sLine = GetVar(TutorialPath, "PAG" & lPage, "Line" & lLine)
                    .sPage = .sPage & sLine & vbNewLine
                Next lLine

            End With
            
        Next lPage

    End If
    
    lblPagTotal.Caption = NumPages

    
    Exit Sub

LoadTutorial_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmTutorial" & "->" & "LoadTutorial"
    End If
Resume Next
    
End Sub

Private Sub SelectPage(ByVal lPage As Long)
    
    On Error GoTo SelectPage_Err
    
    lblTitulo.Caption = Tutorial(lPage).sTitle
    lblMensaje.Caption = Tutorial(lPage).sPage
    lblPagActual.Caption = lPage

    
    Exit Sub

SelectPage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmTutorial" & "->" & "SelectPage"
    End If
Resume Next
    
End Sub

Private Sub lblMensaje_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo lblMensaje_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

lblMensaje_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmTutorial" & "->" & "lblMensaje_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyUp_Err
    
    If KeyCode = vbKeyEscape Then
        Set clsFormulario = Nothing
        Unload Me
    End If

    Exit Sub

Form_KeyUp_Err:

    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmTutorial" & "->" & "Form_KeyUp"

    End If

    Resume Next
    
End Sub


