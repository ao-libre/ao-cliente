VERSION 5.00
Begin VB.Form frmCustomKeys 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar y Salir"
      Height          =   375
      Left            =   3480
      TabIndex        =   54
      Top             =   5640
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar Teclas por defecto"
      Height          =   375
      Left            =   3480
      TabIndex        =   53
      Top             =   5160
      Width           =   4695
   End
   Begin VB.Frame Frame5 
      Caption         =   "Otros"
      Height          =   2415
      Left            =   3480
      TabIndex        =   4
      Top             =   2640
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   1920
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   1920
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   1920
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   1920
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   1920
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   1920
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Capturar Pantalla"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Macro Trabajo"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Macro Hechizos"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Meditar"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar Opciones"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar/Ocultar FPS"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Hablar"
      Height          =   975
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   1560
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   1560
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Hablar al Clan"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Hablar a Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Acciones"
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   1440
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   1440
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   1440
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   1440
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   1440
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1440
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1440
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1440
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Atacar"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Usar"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Tirar"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Ocultar"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Modo Seguro"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Robar"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Domar"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Equipar"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Modo Combate"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Agarrar"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones Personales"
      Height          =   1335
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2640
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   2640
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2640
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar/Ocultar Nombres"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Corregir Posicion"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Activar/Desactivar Musica"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movimiento"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Derecha"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Izquierda"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Abajo"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Arriba"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCustomKeys"
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

''
'frmCustomKeys - Allows the user to customize keys.
'Implements class clsCustomKeys
'
'@author Rapsodius
'@date 20070805
'@version 1.0.0
'@see clsCustomKeys

Option Explicit

Private Sub Command1_Click()
Call CustomKeys.LoadDefaults
Dim i As Long

For i = 1 To CustomKeys.Count
    Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
Next i
End Sub

Private Sub Command2_Click()

Dim i As Long

For i = 1 To CustomKeys.Count
    If LenB(Text1(i).Text) = 0 Then
        Call MsgBox("Hay una o mas teclas no validas, por favor verifique.", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Argentum Online")
        Exit Sub
    End If
Next i

Unload Me
End Sub

Private Sub Form_Load()
Dim i As Long

For i = 1 To CustomKeys.Count
    Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
Next i
End Sub

Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
'If key is not valid, we exit

Text1(index).Text = CustomKeys.ReadableName(KeyCode)
Text1(index).SelStart = Len(Text1(index).Text)

If CustomKeys.KeyAssigned(KeyCode) Then
    Text1(index).Text = "" 'If the key is already assigned, simply reject it
    Call Beep 'Alert the user
    KeyCode = 0
    Exit Sub
End If

'NOTE: Altough class clsCustomKeys automatically checks if key is not valid
'and if it's already assigned, I want to manually implement this checks

CustomKeys.BindedKey(index) = KeyCode

End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
Call Text1_KeyDown(index, KeyCode, Shift)
End Sub
