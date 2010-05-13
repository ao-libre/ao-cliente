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

Private clsFormulario As clsFormMovementManager

Private cBotonAgregar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton
Private cBotonDisolver As clsGraphicalButton
Private cBotonLiderGrupo As clsGraphicalButton
Private cBotonExpulsar As clsGraphicalButton
Private cBotonSalirParty As clsGraphicalButton

Public LastPressed As clsGraphicalButton


Private sPartyChat As String
Private Const LEADER_FORM_HEIGHT As Integer = 6015
Private Const NORMAL_FORM_HEIGHT As Integer = 4455
Private Const OFFSET_BUTTONS As Integer = 43 ' pixels


Private Sub Form_Load()
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
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = DirGraficos

    Set cBotonAgregar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonDisolver = New clsGraphicalButton
    Set cBotonLiderGrupo = New clsGraphicalButton
    Set cBotonExpulsar = New clsGraphicalButton
    Set cBotonSalirParty = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonAgregar.Initialize(imgAgregar, GrhPath & "BotonAgregarParty.jpg", _
                                    GrhPath & "BotonAgregarRolloverParty.jpg", _
                                    GrhPath & "BotonAgregarClickParty.jpg", Me)
                                    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarParty.jpg", _
                                    GrhPath & "BotonCerrarRolloverParty.jpg", _
                                    GrhPath & "BotonCerrarClickParty.jpg", Me)
                                    
    Call cBotonDisolver.Initialize(imgDisolver, GrhPath & "BotonDisolverParty.jpg", _
                                    GrhPath & "BotonDisolverRolloverParty.jpg", _
                                    GrhPath & "BotonDisolverClickParty.jpg", Me)
                                    
    Call cBotonLiderGrupo.Initialize(imgLiderGrupo, GrhPath & "BotonLiderGrupoParty.jpg", _
                                    GrhPath & "BotonLiderGrupoRolloverParty.jpg", _
                                    GrhPath & "BotonLiderGrupoClickParty.jpg", Me)
                                    
    Call cBotonExpulsar.Initialize(imgExpulsar, GrhPath & "BotonExpulsarParty.jpg", _
                                    GrhPath & "BotonExpulsarRolloverParty.jpg", _
                                    GrhPath & "BotonExpulsarClickParty.jpg", Me)
                                    
    Call cBotonSalirParty.Initialize(imgSalirParty, GrhPath & "BotonSalirGrupoParty.jpg", _
                                    GrhPath & "BotonSalirGrupoRolloverParty.jpg", _
                                    GrhPath & "BotonSalirGrupoClickParty.jpg", Me)
                                    
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


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoParty = False
End Sub

Private Sub imgAgregar_Click()
    If Len(txtToAdd) > 0 Then
        If Not IsNumeric(txtToAdd) Then
            Call WritePartyAcceptMember(Trim(txtToAdd.Text))
            Unload Me
            Call WriteRequestPartyForm
        End If
    End If
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgDisolver_Click()
    Call WritePartyLeave
    Unload Me
End Sub

Private Sub imgExpulsar_Click()
   
    If lstMembers.ListIndex < 0 Then Exit Sub
    
    Dim fName As String
    fName = GetName
    
    If fName <> "" Then
        Call WritePartyKick(fName)
        Unload Me
        
        ' Para que no llame al form si disolvió la party
        If UCase$(fName) <> UCase$(UserName) Then Call WriteRequestPartyForm
    End If

End Sub

Private Function GetName() As String
'**************************************************************
'Author: ZaMa
'Last Modify Date: 27/12/2009
'**************************************************************
    Dim sName As String
    
    sName = Trim(mid(lstMembers.List(lstMembers.ListIndex), 1, InStr(lstMembers.List(lstMembers.ListIndex), " (")))
    If Len(sName) > 0 Then GetName = sName
        
End Function

Private Sub imgLiderGrupo_Click()
    
    If lstMembers.ListIndex < 0 Then Exit Sub
    
    Dim sName As String
    sName = GetName
    
    If sName <> "" Then
        Call WritePartySetLeader(sName)
        Unload Me
        Call WriteRequestPartyForm
    End If
End Sub

Private Sub imgSalirParty_Click()
    Call WritePartyLeave
    Unload Me
End Sub

Private Sub lstMembers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If EsPartyLeader Then
        LastPressed.ToggleToNormal
    End If
    
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 03/10/2009
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        sPartyChat = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
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
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(sPartyChat) <> 0 Then Call WritePartyMessage(sPartyChat)
        
        sPartyChat = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.SetFocus
    End If
End Sub

Private Sub txtToAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub txtToAdd_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub txtToAdd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then imgAgregar_Click
End Sub


