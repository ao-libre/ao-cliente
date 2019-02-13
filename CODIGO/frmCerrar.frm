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
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cCancelQuit 
      Caption         =   "Salir (ESC)"
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   1140
   End
   Begin VB.CommandButton cSalir 
      Caption         =   "Salir del Juego"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   4170
   End
   Begin VB.CommandButton cRegresar 
      Caption         =   "Regresar a la pantalla de inicio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   4170
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
    
    On Error GoTo cCancelQuit_Click_Err
    
    Call Audio.PlayWave(SND_CLICK)
    Set clsFormulario = Nothing
    Unload Me

    
    Exit Sub

cCancelQuit_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCerrar" & "->" & "cCancelQuit_Click"
    End If
Resume Next
    
End Sub

Private Sub cRegresar_Click()
    
    On Error GoTo cRegresar_Click_Err
    
    Call Audio.PlayWave(SND_CLICK)
    
    Set clsFormulario = Nothing
    
    If UserParalizado Then 'Inmo

        With FontTypes(FontTypeNames.FONTTYPE_WARNING)
            Call ShowConsoleMsg("No puedes salir estando paralizado.", .Red, .Green, .Blue, .bold, .italic)

        End With

        Exit Sub

    End If
    
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    
    Call WriteQuit
    Unload Me

    
    Exit Sub

cRegresar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCerrar" & "->" & "cRegresar_Click"
    End If
Resume Next
    
End Sub

Private Sub cSalir_Click()
    
    On Error GoTo cSalir_Click_Err
    
    Call Audio.PlayWave(SND_CLICK)
    Set clsFormulario = Nothing
    Call CloseClient

    
    Exit Sub

cSalir_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCerrar" & "->" & "cSalir_Click"
    End If
Resume Next
    
End Sub

Private Sub Form_Deactivate()
    
    On Error GoTo Form_Deactivate_Err
    
    Me.SetFocus

    Exit Sub

Form_Deactivate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCerrar" & "->" & "Form_Deactivate"
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
        LogError Err.number, Err.Description, "frmCerrar" & "->" & "Form_KeyUp"
    End If
Resume Next
    
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    
    On Error GoTo Form_Load_Err
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = LoadPicture(App.path & "\Graficos\frmCerrar.jpg")

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmCerrar" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

