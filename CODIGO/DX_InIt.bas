Attribute VB_Name = "Mod_DX"
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

Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7

Public PrimarySurface As DirectDrawSurface7
Public PrimaryClipper As DirectDrawClipper
Public SecundaryClipper As DirectDrawClipper
Public BackBufferSurface As DirectDrawSurface7

'Public SurfaceDB() As DirectDrawSurface7

'### 08/04/03 ###
#If (UsarDinamico = 1) Then
    Public SurfaceDB As New CBmpMan
#Else
    Public SurfaceDB As New CBmpManNoDyn
#End If

Public oldResHeight As Long, oldResWidth As Long
Public bNoResChange As Boolean

Private Sub IniciarDXobject(DX As DirectX7)

Err.Clear

On Error Resume Next

Set DX = New DirectX7

If Err Then
    MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
    LogError "Error producido por Set DX = New DirectX7"
    End
End If

End Sub

Private Sub IniciarDDobject(DD As DirectDraw7)
Err.Clear
On Error Resume Next
Set DD = DirectX.DirectDrawCreate("")
If Err Then
    MsgBox "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
    LogError "Error producido en Private Sub IniciarDDobject(DD As DirectDraw7)"
    End
End If
End Sub

Public Sub IniciarObjetosDirectX()

On Error Resume Next

Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectX....", 0, 0, 0, 0, 0, True)
Call IniciarDXobject(DirectX)
Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1, , False)

Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectDraw....", 0, 0, 0, 0, 0, True)
Call IniciarDDobject(DirectDraw)
Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1, , False)

Call AddtoRichTextBox(frmCargando.status, "Analizando y preparando la placa de video....", 0, 0, 0, 0, 0, True)

  
    
Dim lRes As Long
Dim MidevM As typDevMODE
lRes = EnumDisplaySettings(0, 0, MidevM)
    
Dim intWidth As Integer
Dim intHeight As Integer

oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
oldResHeight = Screen.Height \ Screen.TwipsPerPixelY

Dim CambiarResolucion As Boolean

If NoRes Then
    CambiarResolucion = (oldResWidth < 800 Or oldResHeight < 600)
Else
    CambiarResolucion = (oldResWidth <> 800 Or oldResHeight <> 600)
End If

If CambiarResolucion Then
      With MidevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
            .dmBitsPerPel = 16
      End With
      lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
Else
      bNoResChange = True
End If

Call AddtoRichTextBox(frmCargando.status, "¡DirectX OK!", 0, 251, 0, 1, 0)

Exit Sub

End Sub

Public Sub LiberarObjetosDX()
Err.Clear
On Error GoTo fin:
Dim loopc As Integer

Set PrimarySurface = Nothing
Set PrimaryClipper = Nothing
Set BackBufferSurface = Nothing

Call SurfaceDB.BorrarTodo

Set DirectDraw = Nothing

Set DirectX = Nothing
Exit Sub
fin: LogError "Error producido en Public Sub LiberarObjetosDX()"
End Sub

