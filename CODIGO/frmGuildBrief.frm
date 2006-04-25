VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
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
   ScaleHeight     =   7260
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Ofrecer Paz"
      Height          =   375
      Left            =   1680
      MouseIcon       =   "frmGuildBrief.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton aliado 
      Caption         =   "Ofrecer Alianza"
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmGuildBrief.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Guerra 
      Caption         =   "Declarar Guerra"
      Height          =   375
      Left            =   4560
      MouseIcon       =   "frmGuildBrief.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Solicitar Ingreso"
      Height          =   375
      Left            =   6000
      MouseIcon       =   "frmGuildBrief.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildBrief.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   5400
      Width           =   7215
      Begin VB.TextBox Desc 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Codex"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   2970
      Width           =   7215
      Begin VB.Label Codex 
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   11
         Top             =   600
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   120
      TabIndex        =   0
      Top             =   -15
      Width           =   7215
      Begin VB.Label antifaccion 
         Caption         =   "Puntos Antifaccion:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   6975
      End
      Begin VB.Label Aliados 
         Caption         =   "Clanes Aliados:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   6975
      End
      Begin VB.Label Enemigos 
         Caption         =   "Clanes Enemigos:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   6975
      End
      Begin VB.Label lblAlineacion 
         Caption         =   "Alineacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   6975
      End
      Begin VB.Label eleccion 
         Caption         =   "Elecciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   6975
      End
      Begin VB.Label Miembros 
         Caption         =   "Miembros:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   6975
      End
      Begin VB.Label web 
         Caption         =   "Web site:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   6975
      End
      Begin VB.Label lider 
         Caption         =   "Lider:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   6975
      End
      Begin VB.Label creacion 
         Caption         =   "Fecha de creacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   6975
      End
      Begin VB.Label fundador 
         Caption         =   "Fundador:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   6975
      End
      Begin VB.Label nombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6975
      End
   End
End
Attribute VB_Name = "frmGuildBrief"
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

Public EsLeader As Boolean


Public Sub ParseGuildInfo(ByVal Buffer As String)

If Not EsLeader Then
    guerra.Visible = False
    aliado.Visible = False
    Command3.Visible = False
Else
    guerra.Visible = True
    aliado.Visible = True
    Command3.Visible = True
End If

Nombre.Caption = "Nombre:" & ReadField(1, Buffer, Asc("¬"))
fundador.Caption = "Fundador:" & ReadField(2, Buffer, Asc("¬"))
creacion.Caption = "Fecha de creacion:" & ReadField(3, Buffer, Asc("¬"))
lider.Caption = "Lider:" & ReadField(4, Buffer, Asc("¬"))
web.Caption = "Web site:" & ReadField(5, Buffer, Asc("¬"))
Miembros.Caption = "Miembros:" & ReadField(6, Buffer, Asc("¬"))
eleccion.Caption = "Dias para proxima eleccion de lider:" & ReadField(7, Buffer, Asc("¬"))
'Oro.Caption = "Oro:" & ReadField(8, Buffer, Asc("¬"))
lblAlineacion.Caption = "Alineación: " & ReadField(8, Buffer, Asc("¬"))
Enemigos.Caption = "Clanes enemigos:" & ReadField(9, Buffer, Asc("¬"))
aliados.Caption = "Clanes aliados:" & ReadField(10, Buffer, Asc("¬"))
antifaccion.Caption = "Puntos Antifaccion: " & ReadField(11, Buffer, Asc("¬"))

Dim T As Long

For T = 1 To 8
    Codex(T - 1).Caption = ReadField(11 + T, Buffer, Asc("¬"))
Next T

Dim des As String

des = ReadField(20, Buffer, Asc("¬"))
desc.Text = Replace(des, "º", vbCrLf)

Me.Show vbModal, frmMain

End Sub

Private Sub aliado_Click()
frmCommet.Nombre = Right(Nombre.Caption, Len(Nombre.Caption) - 7)
frmCommet.T = ALIANZA
frmCommet.Caption = "Ingrese propuesta de alianza"
Call frmCommet.Show(vbModal, frmGuildBrief)

'Call SendData("OFRECALI" & Right(Nombre, Len(Nombre) - 7))
'Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

Call frmGuildSol.RecieveSolicitud(Right$(Nombre, Len(Nombre) - 7))
Call frmGuildSol.Show(vbModal, frmGuildBrief)
'Unload Me

End Sub

Private Sub Command3_Click()
frmCommet.Nombre = Right(Nombre.Caption, Len(Nombre.Caption) - 7)
frmCommet.T = PAZ
frmCommet.Caption = "Ingrese propuesta de paz"
Call frmCommet.Show(vbModal, frmGuildBrief)
'Unload Me
End Sub


Private Sub Guerra_Click()
Call SendData("DECGUERR" & Right(Nombre.Caption, Len(Nombre.Caption) - 7))
Unload Me
End Sub
