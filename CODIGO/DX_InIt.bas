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

Public Const NumSoundBuffers = 20

Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7
Public DirectSound As DirectSound

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

Public Perf As DirectMusicPerformance
Public Seg As DirectMusicSegment
Public SegState As DirectMusicSegmentState
Public Loader As DirectMusicLoader

Public oldResHeight As Long, oldResWidth As Long
Public bNoResChange As Boolean

Public LastSoundBufferUsed As Integer
Public DSBuffers(1 To NumSoundBuffers) As DirectSoundBuffer

Public ddsd2 As DDSURFACEDESC2
Public ddsd4 As DDSURFACEDESC2
Public ddsd5 As DDSURFACEDESC2
Public ddsAlphaPicture As DirectDrawSurface7
Public ddsSpotLight As DirectDrawSurface7


Private Sub IniciarDirectSound()
Err.Clear
On Error GoTo fin
    Set DirectSound = DirectX.DirectSoundCreate("")
    If Err Then
        MsgBox "Error iniciando DirectSound"
        End
    End If
    
    LastSoundBufferUsed = 1
    '<----------------Direct Music--------------->
    Set Perf = DirectX.DirectMusicPerformanceCreate()
    Call Perf.Init(Nothing, 0)
    Perf.SetPort -1, 80
    Call Perf.SetMasterAutoDownload(True)
    '<------------------------------------------->
    Exit Sub
fin:

LogError "Error al iniciar IniciarDirectSound, asegurese de tener bien configurada la placa de sonido."

Musica = 1
Fx = 1

End Sub

Private Sub LiberarDirectSound()
Dim cloop As Integer
For cloop = 1 To NumSoundBuffers
    Set DSBuffers(cloop) = Nothing
Next cloop
Set DirectSound = Nothing
End Sub

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

If Musica = 0 Or Fx = 0 Then
    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound....", 0, 0, 0, 0, 0, True)
    Call IniciarDirectSound
    Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1, , False)
End If

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

LiberarDirectSound

Call SurfaceDB.BorrarTodo

Set DirectDraw = Nothing

For loopc = 1 To NumSoundBuffers
    Set DSBuffers(loopc) = Nothing
Next loopc


Set Loader = Nothing
Set Perf = Nothing
Set Seg = Nothing
Set DirectSound = Nothing

Set DirectX = Nothing
Exit Sub
fin: LogError "Error producido en Public Sub LiberarObjetosDX()"
End Sub

