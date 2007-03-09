VERSION 5.00
Begin VB.Form frmGuildAdm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
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
   ScaleHeight     =   3390
   ScaleWidth      =   4065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildAdm.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Solicitar Ingreso"
      Height          =   375
      Left            =   1080
      MouseIcon       =   "frmGuildAdm.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Detalles"
      Height          =   375
      Left            =   2640
      MouseIcon       =   "frmGuildAdm.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ListBox GuildsList 
         Height          =   2010
         ItemData        =   "frmGuildAdm.frx":03F6
         Left            =   240
         List            =   "frmGuildAdm.frx":03F8
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmGuildAdm"
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
    frmGuildBrief.EsLeader = False
    Call WriteGuildRequestDetails(guildslist.List(guildslist.listIndex))
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub
