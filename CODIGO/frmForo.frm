VERSION 5.00
Begin VB.Form frmForo 
   BackColor       =   &H00404080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000C0&
      Caption         =   "Lista de mensajes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2205
      MouseIcon       =   "frmForo.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6330
      Width           =   1560
   End
   Begin VB.TextBox MiMensaje 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   5070
      Index           =   1
      Left            =   330
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1005
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.TextBox MiMensaje 
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
      ForeColor       =   &H80000005&
      Height          =   345
      Index           =   0
      Left            =   330
      TabIndex        =   4
      Top             =   285
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      MouseIcon       =   "frmForo.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6330
      Width           =   1560
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "Dejar Mensaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      MouseIcon       =   "frmForo.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6330
      Width           =   1560
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   5505
      Index           =   0
      Left            =   330
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmForo.frx":03F6
      Top             =   285
      Visible         =   0   'False
      Width           =   5430
   End
   Begin VB.ListBox List 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   5520
      Left            =   330
      TabIndex        =   0
      Top             =   285
      Width           =   5430
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   330
      TabIndex        =   7
      Top             =   765
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   330
      TabIndex        =   6
      Top             =   45
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmForo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
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

Public ForoIndex As Integer
Private Sub Command1_Click()
Dim i
For Each i In Text
    i.Visible = False
Next

If Not MiMensaje(0).Visible Then
    List.Visible = False
    MiMensaje(0).Visible = True
    MiMensaje(1).Visible = True
    MiMensaje(0).SetFocus
    Command1.Enabled = False
    Label1.Visible = True
    Label2.Visible = True
Else
    Call SendData("DEMSG" & MiMensaje(0).Text & Chr(176) & Left(MiMensaje(1).Text, 450))
    List.AddItem MiMensaje(0).Text
    Load Text(List.ListCount)
    Text(List.ListCount - 1).Text = MiMensaje(1).Text
    List.Visible = True
    
    MiMensaje(0).Visible = False
    MiMensaje(1).Visible = False
    Command1.Enabled = True
    Label1.Visible = False
    Label2.Visible = False
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

MiMensaje(0).Visible = False
MiMensaje(1).Visible = False
Command1.Enabled = True
Label1.Visible = False
Label2.Visible = False
Dim i
For Each i In Text
    i.Visible = False
Next
List.Visible = True
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub


Private Sub List_Click()
List.Visible = False
Text(List.ListIndex).Visible = True

End Sub

Private Sub MiMensaje_Change(Index As Integer)
If Len(MiMensaje(0).Text) <> 0 And Len(MiMensaje(1).Text) <> 0 Then
Command1.Enabled = True
End If

End Sub

