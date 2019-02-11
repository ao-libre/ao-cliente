VERSION 5.00
Begin VB.Form frmParty 
   BorderStyle     =   0  'None
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   428
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox SendTxt 
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
      Height          =   255
      Left            =   555
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   720
      Width           =   4530
   End
   Begin VB.TextBox txtToAdd 
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
      Height          =   240
      Left            =   1530
      MaxLength       =   20
      TabIndex        =   1
      Top             =   4365
      Width           =   2580
   End
   Begin VB.ListBox lstMembers 
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
      Height          =   1395
      Left            =   1530
      TabIndex        =   0
      Top             =   1590
      Width           =   2595
   End
   Begin VB.Label lblTotalExp 
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   3075
      TabIndex        =   3
      Top             =   3150
      Width           =   1335
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   3840
      Tag             =   "1"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Image imgLiderGrupo 
      Height          =   360
      Left            =   2880
      Tag             =   "1"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Image imgExpulsar 
      Height          =   360
      Left            =   1320
      Tag             =   "1"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Image imgAgregar 
      Height          =   360
      Left            =   2040
      Tag             =   "1"
      Top             =   4830
      Width           =   1455
   End
   Begin VB.Image imgSalirParty 
      Height          =   375
      Left            =   300
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Image imgDisolver 
      Height          =   360
      Left            =   300
      Tag             =   "1"
      Top             =   5400
      Width           =   1455
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmParty.frm
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

Private clsFormulario            As clsFormMovementManager

Private cBotonAgregar            As clsGraphicalButton
Private cBotonCerrar             As clsGraphicalButton
Private cBotonDisolver           As clsGraphicalButton
Private cBotonLiderGrupo         As clsGraphicalButton
Private cBotonExpulsar           As clsGraphicalButton
Private cBotonSalirParty         As clsGraphicalButton

Public LastButtonPressed         As clsGraphicalButton

Private sPartyChat               As String
Private Const LEADER_FORM_HEIGHT As Integer = 6015
Private Const NORMAL_FORM_HEIGHT As Integer = 4455
Private Const OFFSET_BUTTONS     As Integer = 43 ' pixels

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    lstMembers.Clear
        
    If EsPartyLeader Then
        Me.Picture = LoadPicture(App.path & "\graficos\VentanaPartyLider.jpg")
        Me.Height = LEADER_FORM_HEIGHT
    Else
        Me.Picture = LoadPicture(App.path & "\graficos\VentanaPartyMiembro.jpg")
        Me.Height = NORMAL_FORM_HEIGHT

    End If
    
    Call LoadButtons

    MirandoParty = True
    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub LoadButtons()
    
    On Error GoTo LoadButtons_Err
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonAgregar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonDisolver = New clsGraphicalButton
    Set cBotonLiderGrupo = New clsGraphicalButton
    Set cBotonExpulsar = New clsGraphicalButton
    Set cBotonSalirParty = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonAgregar.Initialize(imgAgregar, GrhPath & "BotonAgregarParty.jpg", GrhPath & "BotonAgregarRolloverParty.jpg", GrhPath & "BotonAgregarClickParty.jpg", Me)
                                    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarParty.jpg", GrhPath & "BotonCerrarRolloverParty.jpg", GrhPath & "BotonCerrarClickParty.jpg", Me)
                                    
    Call cBotonDisolver.Initialize(imgDisolver, GrhPath & "BotonDisolverParty.jpg", GrhPath & "BotonDisolverRolloverParty.jpg", GrhPath & "BotonDisolverClickParty.jpg", Me)
                                    
    Call cBotonLiderGrupo.Initialize(imgLiderGrupo, GrhPath & "BotonLiderGrupoParty.jpg", GrhPath & "BotonLiderGrupoRolloverParty.jpg", GrhPath & "BotonLiderGrupoClickParty.jpg", Me)
                                    
    Call cBotonExpulsar.Initialize(imgExpulsar, GrhPath & "BotonExpulsarParty.jpg", GrhPath & "BotonExpulsarRolloverParty.jpg", GrhPath & "BotonExpulsarClickParty.jpg", Me)
                                    
    Call cBotonSalirParty.Initialize(imgSalirParty, GrhPath & "BotonSalirGrupoParty.jpg", GrhPath & "BotonSalirGrupoRolloverParty.jpg", GrhPath & "BotonSalirGrupoClickParty.jpg", Me)
                                    
    ' Botones visibles solo para el lider
    imgExpulsar.Visible = EsPartyLeader
    imgLiderGrupo.Visible = EsPartyLeader
    txtToAdd.Visible = EsPartyLeader
    imgAgregar.Visible = EsPartyLeader
    
    imgDisolver.Visible = EsPartyLeader
    imgSalirParty.Visible = Not EsPartyLeader
    
    imgSalirParty.Top = Me.ScaleHeight - OFFSET_BUTTONS
    imgDisolver.Top = Me.ScaleHeight - OFFSET_BUTTONS
    imgCerrar.Top = Me.ScaleHeight - OFFSET_BUTTONS

    
    Exit Sub

LoadButtons_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "LoadButtons"
    End If
Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

Form_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "Form_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Form_Unload_Err
    
    MirandoParty = False

    
    Exit Sub

Form_Unload_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "Form_Unload"
    End If
Resume Next
    
End Sub

Private Sub imgAgregar_Click()
    
    On Error GoTo imgAgregar_Click_Err
    

    If Len(txtToAdd) > 0 Then
        If Not IsNumeric(txtToAdd) Then
            Call WritePartyAcceptMember(Trim$(txtToAdd.Text))
            Unload Me
            Call WriteRequestPartyForm

        End If

    End If

    
    Exit Sub

imgAgregar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "imgAgregar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgCerrar_Click()
    
    On Error GoTo imgCerrar_Click_Err
    
    Unload Me

    
    Exit Sub

imgCerrar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "imgCerrar_Click"
    End If
Resume Next
    
End Sub

Private Sub imgDisolver_Click()
    
    On Error GoTo imgDisolver_Click_Err
    
    Call WritePartyLeave
    Unload Me

    
    Exit Sub

imgDisolver_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "imgDisolver_Click"
    End If
Resume Next
    
End Sub

Private Sub imgExpulsar_Click()
    
    On Error GoTo imgExpulsar_Click_Err
    
   
    If lstMembers.ListIndex < 0 Then Exit Sub
    
    Dim fName As String
    fName = GetName
    
    If fName <> "" Then
        Call WritePartyKick(fName)
        Unload Me
        
        ' Para que no llame al form si disolvió la party
        If UCase$(fName) <> UCase$(UserName) Then Call WriteRequestPartyForm

    End If

    
    Exit Sub

imgExpulsar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "imgExpulsar_Click"
    End If
Resume Next
    
End Sub

Private Function GetName() As String
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 27/12/2009
    '**************************************************************
    
    On Error GoTo GetName_Err
    
    Dim sName As String
    
    sName = Trim$(mid$(lstMembers.List(lstMembers.ListIndex), 1, InStr(lstMembers.List(lstMembers.ListIndex), " (")))

    If Len(sName) > 0 Then GetName = sName
        
    
    Exit Function

GetName_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "GetName"
    End If
Resume Next
    
End Function

Private Sub imgLiderGrupo_Click()
    
    On Error GoTo imgLiderGrupo_Click_Err
    
    
    If lstMembers.ListIndex < 0 Then Exit Sub
    
    Dim sName As String
    sName = GetName
    
    If sName <> "" Then
        Call WritePartySetLeader(sName)
        Unload Me
        Call WriteRequestPartyForm

    End If

    
    Exit Sub

imgLiderGrupo_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "imgLiderGrupo_Click"
    End If
Resume Next
    
End Sub

Private Sub imgSalirParty_Click()
    
    On Error GoTo imgSalirParty_Click_Err
    
    Call WritePartyLeave
    Unload Me

    
    Exit Sub

imgSalirParty_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "imgSalirParty_Click"
    End If
Resume Next
    
End Sub

Private Sub lstMembers_MouseDown(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo lstMembers_MouseDown_Err
    

    If EsPartyLeader Then
        LastButtonPressed.ToggleToNormal

    End If
    
    
    Exit Sub

lstMembers_MouseDown_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "lstMembers_MouseDown"
    End If
Resume Next
    
End Sub

Private Sub SendTxt_Change()
    
    On Error GoTo SendTxt_Change_Err
    

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 03/10/2009
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        sPartyChat = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i         As Long
        Dim tempstr   As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)

            End If

        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr

        End If
        
        sPartyChat = SendTxt.Text

    End If

    
    Exit Sub

SendTxt_Change_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "SendTxt_Change"
    End If
Resume Next
    
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    
    On Error GoTo SendTxt_KeyPress_Err
    

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

    
    Exit Sub

SendTxt_KeyPress_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "SendTxt_KeyPress"
    End If
Resume Next
    
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo SendTxt_KeyUp_Err
    

    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(sPartyChat) <> 0 Then Call WritePartyMessage(sPartyChat)
        
        sPartyChat = vbNullString
        SendTxt.Text = vbNullString
        KeyCode = 0
        SendTxt.SetFocus

    End If

    
    Exit Sub

SendTxt_KeyUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "SendTxt_KeyUp"
    End If
Resume Next
    
End Sub

Private Sub txtToAdd_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    On Error GoTo txtToAdd_MouseMove_Err
    
    LastButtonPressed.ToggleToNormal

    
    Exit Sub

txtToAdd_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "txtToAdd_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub txtToAdd_KeyPress(KeyAscii As Integer)
    
    On Error GoTo txtToAdd_KeyPress_Err
    

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

    
    Exit Sub

txtToAdd_KeyPress_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "txtToAdd_KeyPress"
    End If
Resume Next
    
End Sub

Private Sub txtToAdd_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo txtToAdd_KeyUp_Err
    

    If KeyCode = vbKeyReturn Then imgAgregar_Click

    
    Exit Sub

txtToAdd_KeyUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmParty" & "->" & "txtToAdd_KeyUp"
    End If
Resume Next
    
End Sub

