VERSION 5.00
Begin VB.Form frmHerrero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herrero"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      MouseIcon       =   "frmHerrero.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3015
      Width           =   1710
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Construir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2520
      MouseIcon       =   "frmHerrero.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3015
      Width           =   1710
   End
   Begin VB.ListBox lstArmas 
      Height          =   2010
      Left            =   150
      TabIndex        =   2
      Top             =   810
      Width           =   4080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Armaduras"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2370
      MouseIcon       =   "frmHerrero.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   210
      Width           =   1710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Armas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      MouseIcon       =   "frmHerrero.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   210
      Width           =   1710
   End
   Begin VB.ListBox lstArmaduras 
      Height          =   2010
      Left            =   135
      TabIndex        =   5
      Top             =   825
      Width           =   4080
   End
End
Attribute VB_Name = "frmHerrero"
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

Private Sub Command1_Click()
lstArmaduras.Visible = False
lstArmas.Visible = True
End Sub

Private Sub Command2_Click()
lstArmaduras.Visible = True
lstArmas.Visible = False
End Sub

Private Sub Command3_Click()
On Error Resume Next

    If lstArmas.Visible Then
        Call WriteCraftBlacksmith(ArmasHerrero(lstArmas.listIndex + 1))
        
        If frmMain.macrotrabajo.Enabled Then _
            MacroBltIndex = ArmasHerrero(lstArmas.listIndex + 1)
    Else
        Call WriteCraftBlacksmith(ArmadurasHerrero(lstArmaduras.listIndex + 1))
        
        If frmMain.macrotrabajo.Enabled Then _
            MacroBltIndex = ArmadurasHerrero(lstArmaduras.listIndex + 1)
    End If

    Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

