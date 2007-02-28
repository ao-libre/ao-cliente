VERSION 5.00
Begin VB.Form frmEligeAlineacion 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5250
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblSalir 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   10
      Top             =   4860
      Width           =   915
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   1
      Left            =   1095
      TabIndex        =   9
      Top             =   4140
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":00D5
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   5
      Left            =   1095
      TabIndex        =   8
      Top             =   3240
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00400040&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":01B1
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   2
      Left            =   1095
      TabIndex        =   7
      Top             =   2295
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":025D
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   4
      Left            =   1095
      TabIndex        =   6
      Top             =   1350
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":0326
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Index           =   0
      Left            =   1095
      TabIndex        =   5
      Top             =   270
      Width           =   5505
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación del mal"
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
      Index           =   1
      Left            =   1005
      TabIndex        =   4
      Top             =   3915
      Width           =   1680
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación criminal"
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
      Index           =   5
      Left            =   1005
      TabIndex        =   3
      Top             =   3015
      Width           =   1680
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación neutral"
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
      Index           =   2
      Left            =   1005
      TabIndex        =   2
      Top             =   2070
      Width           =   1635
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación legal"
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
      Index           =   4
      Left            =   1005
      TabIndex        =   1
      Top             =   1125
      Width           =   1455
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación Real"
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
      Index           =   0
      Left            =   1000
      TabIndex        =   0
      Top             =   45
      Width           =   1455
   End
End
Attribute VB_Name = "frmEligeAlineacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmEligeAlineacion.frm
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

Dim LastColoured As Byte

'odio programar sin tiempo (c) el oso

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDescripcion(LastColoured).BorderStyle = 0
    lblDescripcion(LastColoured).BackStyle = 0
End Sub

Private Sub lblDescripcion_Click(index As Integer)
    Call WriteGuildFundate(index)
    Unload Me
End Sub

Private Sub lblDescripcion_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If LastColoured <> index Then
        lblDescripcion(LastColoured).BorderStyle = 0
        lblDescripcion(LastColoured).BackStyle = 0
    End If
    
    lblDescripcion(index).BorderStyle = 1
    lblDescripcion(index).BackStyle = 1
    
    Select Case index
        Case 0
            lblDescripcion(index).BackColor = &H400000
        Case 4
            lblDescripcion(index).BackColor = &H800000
        Case 2
            lblDescripcion(index).BackColor = 4194368
        Case 5
            lblDescripcion(index).BackColor = &H80&
        Case 1
            lblDescripcion(index).BackColor = &H40&
    End Select
    
    LastColoured = index
End Sub


Private Sub lblNombre_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDescripcion(LastColoured).BorderStyle = 0
    lblDescripcion(LastColoured).BackStyle = 0
End Sub

Private Sub lblSalir_Click()
    Unload Me
End Sub
