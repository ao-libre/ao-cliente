VERSION 5.00
Begin VB.Form frmCambiaMotd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   """ZMOTD"""
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkItalic 
      Caption         =   "Cursiva"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   4320
      Width           =   855
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Negrita"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdMarron 
      Caption         =   "Marron"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdVerde 
      Caption         =   "Verde"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdMorado 
      Caption         =   "Morado"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdAmarillo 
      Caption         =   "Amarillo"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdGris 
      Caption         =   "Gris"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdBlanco 
      Caption         =   "Blanco"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdRojo 
      Caption         =   "Rojo"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdAzul 
      BackColor       =   &H00FF0000&
      Caption         =   "Azul"
      Height          =   375
      Left            =   600
      MaskColor       =   &H00FF0000&
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   4800
      Width           =   3855
   End
   Begin VB.TextBox txtMotd 
      Height          =   2415
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   660
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No olvides agregar los colores al final de cada línea (ver tabla de abajo)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   60
      Width           =   4455
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

Private Sub cmdOk_Click()
    Dim T() As String
    Dim i As Long, N As Long, Pos As Long
    
    If Len(txtMotd.Text) >= 2 Then
        If Right$(txtMotd.Text, 2) = vbCrLf Then txtMotd.Text = Left$(txtMotd.Text, Len(txtMotd.Text) - 2)
    End If
    
    T = Split(txtMotd.Text, vbCrLf)
    
    'hola~1~1~1~1~1
    
    For i = LBound(T) To UBound(T)
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

'A partir de Command2_Click son todos buttons para agregar color al texto
Private Sub cmdAzul_Click()
    txtMotd.Text = txtMotd & "~50~70~250~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
End Sub

Private Sub cmdRojo_Click()
    txtMotd.Text = txtMotd & "~255~0~0~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
End Sub

Private Sub cmdBlanco_Click()
    txtMotd.Text = txtMotd & "~255~255~255~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
End Sub

Private Sub cmdGris_Click()
    txtMotd.Text = txtMotd & "~157~157~157~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
End Sub

Private Sub cmdAmarillo_Click()
    txtMotd.Text = txtMotd & "~244~244~0~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
End Sub

Private Sub cmdMorado_Click()
    txtMotd.Text = txtMotd & "~128~0~128~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
End Sub

Private Sub cmdVerde_Click()
  txtMotd.Text = txtMotd & "~23~104~26~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
End Sub

Private Sub cmdMarron_Click()
    txtMotd.Text = txtMotd & "~97~58~31~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
End Sub
