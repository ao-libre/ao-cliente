VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información del personaje"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
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
   ScaleHeight     =   6195
   ScaleWidth      =   6390
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton desc 
      Caption         =   "Peticion"
      Height          =   495
      Left            =   2655
      MouseIcon       =   "frmCharInfo.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   5625
      Width           =   1000
   End
   Begin VB.CommandButton Echar 
      Caption         =   "Echar"
      Height          =   495
      Left            =   1395
      MouseIcon       =   "frmCharInfo.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5625
      Width           =   1000
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5085
      MouseIcon       =   "frmCharInfo.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5625
      Width           =   1000
   End
   Begin VB.CommandButton Rechazar 
      Caption         =   "Rechazar"
      Height          =   495
      Left            =   3870
      MouseIcon       =   "frmCharInfo.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   5625
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmCharInfo.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   5625
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   135
      TabIndex        =   9
      Top             =   2355
      Width           =   6075
      Begin VB.TextBox txtMiembro 
         Height          =   1110
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   1800
         Width           =   5790
      End
      Begin VB.TextBox txtPeticiones 
         Height          =   1110
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   450
         Width           =   5790
      End
      Begin VB.Label lblMiembro 
         Caption         =   "Ultimos clanes en los que participó:"
         Height          =   255
         Left            =   135
         TabIndex        =   23
         Top             =   1620
         Width           =   2985
      End
      Begin VB.Label lblSolicitado 
         Caption         =   "Ultimas membresías solicitadas:"
         Height          =   255
         Left            =   135
         TabIndex        =   21
         Top             =   270
         Width           =   2985
      End
   End
   Begin VB.Frame charinfo 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6075
      Begin VB.Label reputacion 
         Caption         =   "Reputacion:"
         Height          =   255
         Left            =   3060
         TabIndex        =   20
         Top             =   1560
         Width           =   2445
      End
      Begin VB.Label criminales 
         Caption         =   "Criminales asesinados:"
         Height          =   255
         Left            =   3060
         TabIndex        =   19
         Top             =   1325
         Width           =   2900
      End
      Begin VB.Label Ciudadanos 
         Caption         =   "Ciudadanos asesinados:"
         Height          =   255
         Left            =   3060
         TabIndex        =   18
         Top             =   1080
         Width           =   2850
      End
      Begin VB.Label ejercito 
         Caption         =   "Faccion:"
         Height          =   255
         Left            =   3060
         TabIndex        =   17
         Top             =   844
         Width           =   2880
      End
      Begin VB.Label guildactual 
         Caption         =   "Clan Actual:"
         Height          =   255
         Left            =   3030
         TabIndex        =   16
         Top             =   600
         Width           =   2880
      End
      Begin VB.Label status 
         Caption         =   "Status:"
         Height          =   255
         Left            =   3060
         TabIndex        =   8
         Top             =   1800
         Width           =   2760
      End
      Begin VB.Label Banco 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   2985
      End
      Begin VB.Label Oro 
         Caption         =   "Oro:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   2805
      End
      Begin VB.Label Genero 
         Caption         =   "Genero:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Raza 
         Caption         =   "Raza:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2880
      End
      Begin VB.Label Clase 
         Caption         =   "Clase:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3270
      End
      Begin VB.Label Nivel 
         Caption         =   "Nivel:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   3105
      End
      Begin VB.Label Nombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   135
         TabIndex        =   1
         Top             =   360
         Width           =   5640
      End
   End
End
Attribute VB_Name = "frmCharInfo"
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

Public frmmiembros As Boolean
Public frmsolicitudes As Boolean

Private Sub Aceptar_Click()
frmmiembros = False
frmsolicitudes = False
Call SendData("ACEPTARI" & Trim$(Right(Nombre, Len(Nombre) - 8)))
Unload frmGuildLeader
Call SendData("GLINFO")
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub


Public Sub parseCharInfo(ByVal Rdata As String)

If frmmiembros Then
    Rechazar.Visible = False
    Aceptar.Visible = False
    Echar.Visible = True
    desc.Visible = False
Else
    Rechazar.Visible = True
    Aceptar.Visible = True
    Echar.Visible = False
    desc.Visible = True
End If

'    tstr = Personaje & "¬"1
'    tstr = tstr & GetVar(UserFile, "INIT", "Raza") & "¬"2
'    tstr = tstr & GetVar(UserFile, "INIT", "Clase") & "¬"3
'    tstr = tstr & GetVar(UserFile, "INIT", "Genero") & "¬"4
'    tstr = tstr & GetVar(UserFile, "STATS", "ELV") & "¬"5
'    tstr = tstr & GetVar(UserFile, "STATS", "GLD") & "¬"6
'    tstr = tstr & GetVar(UserFile, "STATS", "Banco") & "¬"7
'    tstr = tstr & GetVar(UserFile, "REP", "Promedio") & "¬"8


Nombre.Caption = "Nombre: " & ReadField(1, Rdata, Asc("¬"))
Raza.Caption = "Raza: " & ReadField(2, Rdata, Asc("¬"))
Clase.Caption = "Clase: " & ReadField(3, Rdata, Asc("¬"))
Genero.Caption = "Genero: " & ReadField(4, Rdata, Asc("¬"))
Nivel.Caption = "Nivel: " & ReadField(5, Rdata, Asc("¬"))
Oro.Caption = "Oro: " & ReadField(6, Rdata, Asc("¬"))
Banco.Caption = "Banco: " & ReadField(7, Rdata, Asc("¬"))
Me.reputacion.Caption = "Reputación: " & ReadField(8, Rdata, Asc("¬"))


'    Peticiones = GetVar(UserFile, "GUILDS", "Pedidos")9
'    tstr = tstr & IIf(Len(Peticiones > 400), ".." & Right$(Peticiones, 400), Peticiones) & "¬"
    
'    Miembro = GetVar(UserFile, "GUILDS", "Miembro")10
'    tstr = tstr & IIf(Len(Miembro) > 400, ".." & Right$(Miembro, 400), Miembro) & "¬"

Me.txtPeticiones.Text = ReadField(9, Rdata, Asc("¬"))
Me.txtMiembro.Text = ReadField(10, Rdata, Asc("¬"))


'GuildActual = val(GetVar(UserFile, "GUILD", "GuildIndex"))11
Me.guildactual.Caption = "Clan: " & ReadField(11, Rdata, Asc("¬"))


'    tstr = tstr & GetVar(UserFile, "FACCIONES", "EjercitoReal") & "¬"12
'    tstr = tstr & GetVar(UserFile, "FACCIONES", "EjercitoCaos") & "¬"13
'    tstr = tstr & GetVar(UserFile, "FACCIONES", "CiudMatados") & "¬"14
'    tstr = tstr & GetVar(UserFile, "FACCIONES", "CrimMatados") & "¬"15

Me.ejercito.Caption = "Ejército: " & IIf(Val(ReadField(12, Rdata, Asc("¬"))) <> 0, "Armada Real", IIf(Val(ReadField(13, Rdata, Asc("¬"))) <> 0, "Legión Oscura", "-"))

Ciudadanos.Caption = "Ciudadanos asesinados: " & ReadField(14, Rdata, Asc("¬"))
criminales.Caption = "Criminales asesinados: " & ReadField(15, Rdata, Asc("¬"))


status.Caption = IIf(Val(ReadField(8, Rdata, Asc("¬"))) > 0, " (Ciudadano)", " (Criminal)")
status.ForeColor = IIf(Val(ReadField(8, Rdata, Asc("¬"))) > 0, vbBlue, vbRed)
Me.Show vbModeless, frmMain


End Sub

Private Sub desc_Click()
Call SendData("ENVCOMEN" & Right(Nombre, Len(Nombre) - 7))
End Sub

Private Sub Echar_Click()
Call SendData("ECHARCLA" & Right(Nombre, Len(Nombre) - 7))
frmmiembros = False
frmsolicitudes = False
Unload frmGuildLeader
Call SendData("GLINFO")
Unload Me
End Sub

Private Sub Rechazar_Click()
Load frmCommet
frmCommet.T = RECHAZOPJ
frmCommet.Nombre = Right$(Nombre, Len(Nombre) - 7)
frmCommet.Caption = "Ingrese motivo para rechazo"
frmCommet.Show vbModeless, frmCharInfo

End Sub
