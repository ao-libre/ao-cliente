Attribute VB_Name = "Mod_MODOS_DE_VIDEO"
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

'Testea si la maquina soporta un modo de video ;-)
Function SoportaDisplay(DD As DirectDraw7, DDSDaTestear As DDSURFACEDESC2) As Boolean
Dim ddsd As DDSURFACEDESC2
Dim DDEM As DirectDrawEnumModes

Set DDEM = DD.GetDisplayModesEnum(DDEDM_DEFAULT, ddsd)

Dim loopc As Integer
Dim flag As Boolean
loopc = 1
   
Do While loopc <> DDEM.GetCount And Not flag

    DDEM.GetItem loopc, ddsd
    flag = ddsd.lHeight = DDSDaTestear.lHeight _
    And ddsd.lWidth = DDSDaTestear.lWidth _
    And ddsd.ddpfPixelFormat.lRGBBitCount = _
    DDSDaTestear.ddpfPixelFormat.lRGBBitCount
    loopc = loopc + 1
Loop
SoportaDisplay = flag
End Function

Function ModosDeVideoIguales(dd1 As DDSURFACEDESC2, dd2 As DDSURFACEDESC2) As Boolean
ModosDeVideoIguales = _
    dd1.lHeight = dd2.lHeight _
    And dd1.lWidth = dd2.lWidth _
    And dd1.ddpfPixelFormat.lRGBBitCount = _
    dd2.ddpfPixelFormat.lRGBBitCount
End Function


