VERSION 5.00
Begin VB.Form frmCommet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oferta de paz o alianza"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmCommet.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   2400
      MouseIcon       =   "frmCommet.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmCommet"
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

Private Const MAX_PROPOSAL_LENGTH As Integer = 520

Public Nombre As String
Public T As TIPO
Public Enum TIPO
    ALIANZA = 1
    PAZ = 2
    RECHAZOPJ = 3
End Enum

Public Sub SetTipo(ByVal T As TIPO)
    Select Case T
        Case TIPO.ALIANZA
            Me.Caption = "Detalle de solicitud de alianza"
            Me.Text1.MaxLength = 200
        Case TIPO.PAZ
            Me.Caption = "Detalle de solicitud de Paz"
            Me.Text1.MaxLength = 200
        Case TIPO.RECHAZOPJ
            Me.Caption = "Detalle de rechazo de membresía"
            Me.Text1.MaxLength = 50
    End Select
End Sub


Private Sub Command1_Click()


If Text1 = "" Then
    If T = PAZ Or T = ALIANZA Then
        MsgBox "Debes redactar un mensaje solicitando la paz o alianza al líder de " & Nombre
    Else
        MsgBox "Debes indicar el motivo por el cual rechazas la membresía de " & Nombre
    End If
    Exit Sub
End If

If T = PAZ Then
    Call WriteGuildOfferPeace(Nombre, Replace(Text1, vbCrLf, "º"))
ElseIf T = ALIANZA Then
    Call WriteGuildOfferAlliance(Nombre, Replace(Text1, vbCrLf, "º"))
ElseIf T = RECHAZOPJ Then
    Call WriteGuildRejectNewMember(Nombre, Replace(Replace(Text1.Text, ",", " "), vbCrLf, " "))
    'Sacamos el char de la lista de aspirantes
    Dim i As Long
    For i = 0 To frmGuildLeader.solicitudes.ListCount - 1
        If frmGuildLeader.solicitudes.List(i) = Nombre Then
            frmGuildLeader.solicitudes.RemoveItem i
            Exit For
        End If
    Next i
    
    Me.Hide
    Unload frmCharInfo
    'Call SendData("GLINFO")
End If
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_Change()
    If Len(Text1.Text) > MAX_PROPOSAL_LENGTH Then _
        Text1.Text = Left$(Text1.Text, MAX_PROPOSAL_LENGTH)
End Sub
