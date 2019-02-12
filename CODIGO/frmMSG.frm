VERSION 5.00
Begin VB.Form frmMSG 
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   45
   ClientWidth     =   2445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMSG.frx":0000
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   163
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   1785
      Left            =   300
      TabIndex        =   0
      Top             =   615
      Width           =   1845
   End
   Begin VB.Image imgCerrar 
      Height          =   420
      Left            =   375
      Tag             =   "1"
      Top             =   2640
      Width           =   1710
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "Traer usuario"
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
   End
End
Attribute VB_Name = "frmMSG"
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

Private cBotonCerrar     As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Const MAX_GM_MSG = 300

Private MisMSG(0 To MAX_GM_MSG) As String
Private Apunt(0 To MAX_GM_MSG)  As Integer

Public Sub CrearGMmSg(Nick As String, msg As String)
    
    On Error GoTo CrearGMmSg_Err
    

    If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem Nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1

    End If

    
    Exit Sub

CrearGMmSg_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "CrearGMmSg"
    End If
Resume Next
    
End Sub

Private Sub Form_Deactivate()
    
    On Error GoTo Form_Deactivate_Err
    
    Me.Visible = False
    List1.Clear

    
    Exit Sub

Form_Deactivate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "Form_Deactivate"
    End If
Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    List1.Clear
    
    Me.Picture = LoadPicture(App.path & "\graficos\VentanaShowSos.jpg")
    
    Call LoadButtons

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarShowSos.jpg", GrhPath & "BotonCerrarRolloverShowSos.jpg", GrhPath & "BotonCerrarClickShowSos.jpg", Me)

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub imgCerrar_Click()
    
    On Error GoTo imgCerrar_Click_Err
    
    Me.Visible = False
    List1.Clear

    
    Exit Sub

imgCerrar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "imgCerrar_Click"
    End If
Resume Next
    
End Sub

Private Sub list1_Click()
    
    On Error GoTo list1_Click_Err
    
    Dim ind As Integer
    ind = Val(ReadField(2, List1.List(List1.ListIndex), Asc("-")))

    
    Exit Sub

list1_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "list1_Click"
    End If
Resume Next
    
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo List1_MouseDown_Err
    

    If Button = vbRightButton Then
        PopupMenu menU_usuario

    End If

    
    Exit Sub

List1_MouseDown_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "List1_MouseDown"
    End If
Resume Next
    
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo List1_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

List1_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "List1_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub mnuBorrar_Click()
    
    On Error GoTo mnuBorrar_Click_Err
    

    If List1.ListIndex < 0 Then Exit Sub
    'Pablo (ToxicWaste)
    Dim aux As String
    aux = mid$(ReadField(1, List1.List(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.ListIndex), Asc("-"))))
    Call WriteSOSRemove(aux)
    '/Pablo (ToxicWaste)
    'Call WriteSOSRemove(List1.List(List1.listIndex))
    
    List1.RemoveItem List1.ListIndex

    
    Exit Sub

mnuBorrar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "mnuBorrar_Click"
    End If
Resume Next
    
End Sub

Private Sub mnuIR_Click()
    'Pablo (ToxicWaste)
    
    On Error GoTo mnuIR_Click_Err
    
    Dim aux As String
    aux = mid$(ReadField(1, List1.List(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.ListIndex), Asc("-"))))
    Call WriteGoToChar(aux)
    '/Pablo (ToxicWaste)
    'Call WriteGoToChar(ReadField(1, List1.List(List1.listIndex), Asc("-")))
    
    
    Exit Sub

mnuIR_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "mnuIR_Click"
    End If
Resume Next
    
End Sub

Private Sub mnutraer_Click()
    'Pablo (ToxicWaste)
    
    On Error GoTo mnutraer_Click_Err
    
    Dim aux As String
    aux = mid$(ReadField(1, List1.List(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.ListIndex), Asc("-"))))
    Call WriteSummonChar(aux)

    'Pablo (ToxicWaste)
    'Call WriteSummonChar(ReadField(1, List1.List(List1.listIndex), Asc("-")))
    
    Exit Sub

mnutraer_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmMSG" & "->" & "mnutraer_Click"
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
        LogError Err.number, Err.Description, "frmMSG" & "->" & "Form_KeyUp"
    End If
Resume Next
    
End Sub
