VERSION 5.00
Begin VB.Form frmNewPassword 
   BorderStyle     =   0  'None
   Caption         =   "Cambiar Contrase�a"
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewPassword.frx":0000
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   375
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2265
      Width           =   4005
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   375
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1545
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   375
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   825
      Width           =   4005
   End
   Begin VB.Image imgAceptar 
      Height          =   495
      Left            =   990
      Picture         =   "frmNewPassword.frx":19156
      Tag             =   "1"
      Top             =   2730
      Width           =   2775
   End
End
Attribute VB_Name = "frmNewPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario    As clsFormMovementManager

Private cBotonAceptar    As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    
    On Error GoTo Form_Load_Err
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(App.path & "\graficos\VentanaCambiarcontrasenia.jpg")
    
    Call LoadButtons

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmNewPassword" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonAceptar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotonAceptarCambiarContrasenia.jpg", GrhPath & "BotonAceptarRolloverCambiarContrasenia.jpg", GrhPath & "BotonAceptarClickCambiarContrasenia.jpg", Me)

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmNewPassword" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmNewPassword" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgAceptar_Click()
    
    On Error GoTo imgAceptar_Click_Err
    

    If Text2.Text <> Text3.Text Then
        Call MsgBox("Las contrase�as no coinciden", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Cambiar Contrase�a")
        Exit Sub

    End If
    
    Call WriteChangePassword(Text1.Text, Text2.Text)
    Unload Me

    
    Exit Sub

imgAceptar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmNewPassword" & "->" & "imgAceptar_Click"
    End If
Resume Next
    
End Sub

Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Text3_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Text3_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmNewPassword" & "->" & "Text3_MouseMove"
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
        LogError Err.number, Err.Description, "frmNewPassword" & "->" & "Form_KeyUp"
    End If
Resume Next
    
End Sub
