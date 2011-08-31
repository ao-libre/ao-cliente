Attribute VB_Name = "Carteles"
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

Private Const XPosCartel = 360
Private Const YPosCartel = 335
Private Const MAXLONG = 40

'Carteles
Public Cartel As Boolean
Private Leyenda As String
Private LeyendaFormateada() As String
Private textura As Integer

Sub InitCartel(ByRef Ley As String, ByVal Grh As Integer)
Dim i As Integer, k As Integer, anti As Integer

If Cartel Then Exit Sub

Leyenda = Ley
textura = Grh
Cartel = True

ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2))) As String
anti = 1
k = 0
i = 0
Call DarFormato(Leyenda, i, k, anti)
i = 0

Do While (Len(LeyendaFormateada(i)) <> 0) And (i < UBound(LeyendaFormateada))
   i = i + 1
Loop

ReDim Preserve LeyendaFormateada(0 To i) As String
End Sub

Private Function DarFormato(ByRef s As String, ByRef i As Integer, ByRef k As Integer, ByRef anti As Integer)
If anti + i <= Len(s) + 1 Then
    If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
        LeyendaFormateada(k) = mid$(s, anti, i + 1)
        k = k + 1
        anti = anti + i + 1
        i = 0
    Else
        i = i + 1
    End If
    Call DarFormato(s, i, k, anti)
End If
End Function

Sub DibujarCartel()
Dim X As Integer, Y As Integer
Dim j As Integer, desp As Integer

If Not Cartel Then Exit Sub

X = XPosCartel + 20
Y = YPosCartel + 60
Call DDrawTransGrhIndextoSurface(textura, XPosCartel, YPosCartel, 0)

For j = 0 To UBound(LeyendaFormateada)
    RenderText X, Y + desp, LeyendaFormateada(j), vbWhite, frmMain.font
    desp = desp + (frmMain.font.size) + 5
Next
End Sub
