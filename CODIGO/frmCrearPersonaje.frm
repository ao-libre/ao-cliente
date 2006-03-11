VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox lstProfesion 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2700
      Width           =   2820
   End
   Begin VB.ComboBox lstGenero 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00D1
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":00DB
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2250
      Width           =   2820
   End
   Begin VB.ComboBox lstRaza 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00EE
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":0101
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   1800
      Width           =   2820
   End
   Begin VB.ComboBox lstHogar 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":012E
      Left            =   9015
      List            =   "frmCrearPersonaje.frx":013E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3510
      Width           =   2565
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000012&
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
      Height          =   225
      Left            =   3570
      TabIndex        =   0
      Top             =   1275
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   4020
      TabIndex        =   33
      Top             =   4260
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   3270
      Stretch         =   -1  'True
      Top             =   4710
      Width           =   2475
   End
   Begin VB.Label Puntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   32
      Top             =   8535
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":0167
      MousePointer    =   99  'Custom
      Top             =   3465
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   7785
      MouseIcon       =   "frmCrearPersonaje.frx":02B9
      MousePointer    =   99  'Custom
      Top             =   3675
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   7
      Left            =   7785
      MouseIcon       =   "frmCrearPersonaje.frx":040B
      MousePointer    =   99  'Custom
      Top             =   3885
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":055D
      MousePointer    =   99  'Custom
      Top             =   4110
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":06AF
      MousePointer    =   99  'Custom
      Top             =   4335
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":0801
      MousePointer    =   99  'Custom
      Top             =   4545
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":0953
      MousePointer    =   99  'Custom
      Top             =   4785
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":0AA5
      MousePointer    =   99  'Custom
      Top             =   4965
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":0BF7
      MousePointer    =   99  'Custom
      Top             =   5175
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":0D49
      MousePointer    =   99  'Custom
      Top             =   5385
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":0E9B
      MousePointer    =   99  'Custom
      Top             =   5610
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":0FED
      MousePointer    =   99  'Custom
      Top             =   5820
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   27
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":113F
      MousePointer    =   99  'Custom
      Top             =   6015
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":1291
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":13E3
      MousePointer    =   99  'Custom
      Top             =   3270
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   2
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":1535
      MousePointer    =   99  'Custom
      Top             =   3495
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":1687
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   6
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":17D9
      MousePointer    =   99  'Custom
      Top             =   3945
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   8
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":192B
      MousePointer    =   99  'Custom
      Top             =   4155
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":1A7D
      MousePointer    =   99  'Custom
      Top             =   4380
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":1BCF
      MousePointer    =   99  'Custom
      Top             =   4605
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   14
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":1D21
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   16
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":1E73
      MousePointer    =   99  'Custom
      Top             =   4995
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   18
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":1FC5
      MousePointer    =   99  'Custom
      Top             =   5220
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":2117
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":2269
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":23BB
      MousePointer    =   99  'Custom
      Top             =   5850
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   26
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":250D
      MousePointer    =   99  'Custom
      Top             =   6075
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   28
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":265F
      MousePointer    =   99  'Custom
      Top             =   6285
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   29
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":27B1
      MousePointer    =   99  'Custom
      Top             =   6270
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":2903
      MousePointer    =   99  'Custom
      Top             =   6495
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   31
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":2A55
      MousePointer    =   99  'Custom
      Top             =   6465
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":2BA7
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":2CF9
      MousePointer    =   99  'Custom
      Top             =   6690
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":2E4B
      MousePointer    =   99  'Custom
      Top             =   6945
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   35
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":2F9D
      MousePointer    =   99  'Custom
      Top             =   6915
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   36
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":30EF
      MousePointer    =   99  'Custom
      Top             =   7170
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":3241
      MousePointer    =   99  'Custom
      Top             =   7125
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   38
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":3393
      MousePointer    =   99  'Custom
      Top             =   7395
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":34E5
      MousePointer    =   99  'Custom
      Top             =   7335
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":3637
      MousePointer    =   99  'Custom
      Top             =   7590
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   7815
      MouseIcon       =   "frmCrearPersonaje.frx":3789
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   135
   End
   Begin VB.Image boton 
      Height          =   645
      Index           =   2
      Left            =   315
      MouseIcon       =   "frmCrearPersonaje.frx":38DB
      MousePointer    =   99  'Custom
      Top             =   3015
      Width           =   660
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   1
      Left            =   1905
      MouseIcon       =   "frmCrearPersonaje.frx":3A2D
      MousePointer    =   99  'Custom
      Top             =   8565
      Width           =   765
   End
   Begin VB.Image boton 
      Height          =   210
      Index           =   0
      Left            =   9075
      MouseIcon       =   "frmCrearPersonaje.frx":3B7F
      MousePointer    =   99  'Custom
      Top             =   8565
      Width           =   1200
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   7995
      TabIndex        =   28
      Top             =   7515
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   7995
      TabIndex        =   27
      Top             =   7306
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   7995
      TabIndex        =   26
      Top             =   7092
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   7995
      TabIndex        =   25
      Top             =   6878
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   7995
      TabIndex        =   24
      Top             =   6664
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7995
      TabIndex        =   23
      Top             =   6450
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   7995
      TabIndex        =   22
      Top             =   6236
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   7995
      TabIndex        =   21
      Top             =   6022
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   7995
      TabIndex        =   20
      Top             =   5808
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   7995
      TabIndex        =   19
      Top             =   5594
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   7995
      TabIndex        =   18
      Top             =   5380
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   7995
      TabIndex        =   17
      Top             =   5166
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   7995
      TabIndex        =   16
      Top             =   4952
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7995
      TabIndex        =   15
      Top             =   4738
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   7995
      TabIndex        =   14
      Top             =   4524
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7995
      TabIndex        =   13
      Top             =   4310
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   7995
      TabIndex        =   12
      Top             =   4096
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7995
      TabIndex        =   11
      Top             =   3882
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   7995
      TabIndex        =   10
      Top             =   3668
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   7995
      TabIndex        =   9
      Top             =   3240
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7995
      TabIndex        =   8
      Top             =   3450
      Width           =   270
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   8775
      Picture         =   "frmCrearPersonaje.frx":3CD1
      Top             =   3900
      Width           =   2985
   End
   Begin VB.Label lbCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   6
      Top             =   3570
      Width           =   225
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   4260
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   4
      Top             =   3240
      Width           =   210
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   3
      Top             =   3900
      Width           =   225
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   2
      Top             =   2910
      Width           =   225
   End
   Begin VB.Label lbFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   1
      Top             =   2580
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
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

Public SkillPoints As Byte

Function CheckData() As Boolean
If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = "" Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim I As Integer
For I = 1 To NUMATRIBUTOS
    If UserAtributos(I) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next I

CheckData = True


End Function

Private Sub boton_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        
        Dim I As Integer
        Dim k As Object
        I = 1
        For Each k In Skill
            UserSkills(I) = k.Caption
            I = I + 1
        Next
        
        UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.List(lstRaza.listIndex)
        UserSexo = lstGenero.List(lstGenero.listIndex)
        UserClase = lstProfesion.List(lstProfesion.listIndex)
        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbInteligencia.Caption)
        UserAtributos(3) = Val(lbAgilidad.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        UserHogar = lstHogar.List(lstHogar.listIndex)
        
        'Barrin 3/10/03
        If CheckData() Then
            frmPasswdSinPadrinos.Show vbModal, Me
        End If
        
    Case 1
        If Musica Then
            Call Audio.PlayMIDI("2.mid")
        End If
        
        frmConnect.FONDO.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        Me.Visible = False
        
        
    Case 2
        Call Audio.PlayWave(SND_DICE)
        Call TirarDados
      
End Select


End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub TirarDados()
'lbFuerza.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbInteligencia.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbAgilidad.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbCarisma.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbConstitucion.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))

#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State = sckConnected Then
#End If
        Call SendData("TIRDAD")
    End If

End Sub

Private Sub Command1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If Index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
SkillPoints = 10
Puntos.Caption = SkillPoints
Me.Picture = LoadPicture(App.Path & "\graficos\CP-Interface.jpg")
imgHogar.Picture = LoadPicture(App.Path & "\graficos\CP-Ullathorpe.jpg")

Dim I As Integer
lstProfesion.Clear
For I = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(I)
Next I

lstProfesion.listIndex = 1

Image1.Picture = LoadPicture(App.Path & "\graficos\" & lstProfesion.Text & ".jpg")
Call TirarDados
End Sub


Private Sub lstProfesion_Click()
On Error Resume Next
Image1.Picture = LoadPicture(App.Path & "\graficos\" & lstProfesion.Text & ".jpg")
End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
