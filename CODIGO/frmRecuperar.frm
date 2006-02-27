VERSION 5.00
Begin VB.Form frmRecuperar 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txtcorreo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   345
      TabIndex        =   3
      Top             =   1830
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   315
      MouseIcon       =   "frmRecuperar.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recuperar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3180
      MouseIcon       =   "frmRecuperar.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtNombre 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   405
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Label lblWhat 
      Caption         =   $"frmRecuperar.frx":02A4
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1995
      Left            =   165
      TabIndex        =   7
      Top             =   645
      Width           =   4290
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Dirección de correo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   375
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Nombre del personaje:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   420
      TabIndex        =   1
      Top             =   810
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmRecuperar.frx":033B
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   4500
   End
End
Attribute VB_Name = "frmRecuperar"
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

Private Sub Command1_Click()
'Ojo
EstadoLogin = RecuperarPass
Me.MousePointer = 11

#If UsarWrench = 1 Then
frmMain.Socket1.HostName = CurServerIp
'frmMain.Socket1.HostName = "127.0.0.1"
frmMain.Socket1.RemotePort = CurServerPasRecPort
frmMain.Socket1.Connect
#Else
If frmMain.Winsock1.State <> sckClosed Then _
    frmMain.Winsock1.Close
frmMain.Winsock1.Connect CurServerIp, CurServerPasRecPort
#End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

