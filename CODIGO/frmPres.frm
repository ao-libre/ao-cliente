VERSION 5.00
Begin VB.Form frmPres 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3400
      Left            =   1125
      Top             =   1230
   End
End
Attribute VB_Name = "frmPres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
Option Explicit

Dim puedo As Boolean

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then If puedo Then Unload Me

End Sub

Private Sub Form_Load()
    Me.Width = 800 * Screen.TwipsPerPixelX
    Me.Height = 600 * Screen.TwipsPerPixelY
    puedo = False

End Sub

Private Sub Timer1_Timer()
    Static ticks As Long

    Dim PresPath As String

    ticks = ticks + 1

    If ticks = 1 Then
        PresPath = DirGraficos & "HostingAlkon.jpg"
        Me.Picture = LoadPicture(PresPath)
        puedo = True

    ElseIf ticks = 2 Then
        PresPath = DirGraficos & "Presentacion" & RandomNumber(5, 8) & ".jpg"
        Me.Picture = LoadPicture(PresPath)
    
    Else
        Unload Me

    End If

End Sub
