Attribute VB_Name = "Resolution"
'**************************************************************
' Resolution.bas - Performs resolution changes.
'
' Designed and implemented by Juan Martin Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martin Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Resolution.bas
' @author   Juan Martin Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.1.0
' @date     20080329

'**************************************************************************
' - HISTORY
'       v1.0.0  -   Initial release ( 2007/08/14 - Juan Martin Sotuyo Dodero )
'       v1.1.0  -   Made it reset original depth and frequency at exit ( 2008/03/29 - Juan Martin Sotuyo Dodero )
'**************************************************************************

Option Explicit

Private Const CCDEVICENAME As Long = 32
Private Const CCFORMNAME As Long = 32
Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const DM_DISPLAYFREQUENCY As Long = &H400000
Private Const CDS_TEST As Long = &H4
Private Const ENUM_CURRENT_SETTINGS As Long = -1

Private Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Private oldResHeight As Long
Private oldResWidth As Long
Private oldDepth As Integer
Private oldFrequency As Long

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

Private MiDevM As typDevMODE

'TODO : Change this to not depend on any external public variable using args instead!

Public Sub SetResolution()
    '**************************************************************************************************************
    'Autor: Unknown
    'Last Modification: 03/29/08
    'Changes the display resolution if needed.
    'Last Modified By: Juan Martin Sotuyo Dodero (Maraxus)
    ' 03/29/2008: Maraxus - Retrieves current settings storing display depth and frequency for proper restoration.
    ' 22/04/2019: Jopi - Arreglado "cambio" de resolucion al elegir "NO".
    '***************************************************************************************************************

    Dim lRes              As Long

    Dim CambiarResolucion As Boolean
    
    lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MiDevM)
   
    oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.Height \ Screen.TwipsPerPixelY
        
    If MsgBox(JsonLanguage.Item("PANTALLA_COMPLETA").Item("TEXTO"), vbYesNo, "Argentum Online") = vbYes Then
        frmMain.WindowState = vbMaximized

        With MiDevM
            .dmBitsPerPel = 32
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
            oldDepth = .dmBitsPerPel
            oldFrequency = .dmDisplayFrequency

        End With
 
        lRes = ChangeDisplaySettings(MiDevM, CDS_TEST)
    
        ' En pantalla chica que pueda mover
        NoRes = True
                        
    Else
        frmMain.WindowState = vbNormal
    
        ' En pantalla grande no porque japish runtime
        NoRes = False
                        
    End If
  
    CambiarResolucion = (oldResWidth < 800 Or oldResHeight < 600)

End Sub

Public Function GetResolutionState() As Boolean
    
    '**************************************************************************************************************
    'Autor: Jopi
    'Se fija si hubo algun cambio de resolucion al iniciar el juego.
    'Se usa en el Sub CloseClient para resetear la  resolucion si devuelve TRUE.
    '***************************************************************************************************************
    
    ' Obtenemos los parametros actuales de la resolucion
    Dim lRes As Long
        lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MiDevM)
    
    ' Si la resolucion actual es diferente a original significa que la cambiamos, devolvemos TRUE
    ' Si la resolucion actual es la misma que la original, devolvemos FALSE
    If MiDevM.dmPelsWidth <> oldResWidth And MiDevM.dmPelsHeight <> oldResHeight Then
        GetResolutionState = True
    Else
        GetResolutionState = False

    End If
    
End Function

Public Sub ResetResolution()

    '*************************************************************************************************************
    'Autor: Unknown
    'Last Modification: 22/04/2019
    'Changes the display resolution if needed.
    'Last Modified By: Jopi
    ' 03/29/2008: Maraxus - Properly restores display depth and frequency.
    ' 22/04/2019: Jopi - No se cambiaba la resolucion al llamar al Sub CloseClient estando en pantalla completa.
    '*************************************************************************************************************
    
    Dim lRes    As Long
        lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MiDevM)
        
    With MiDevM
        .dmFields = DM_PELSWIDTH And DM_PELSHEIGHT And DM_BITSPERPEL And DM_DISPLAYFREQUENCY
        .dmPelsWidth = oldResWidth
        .dmPelsHeight = oldResHeight
        .dmBitsPerPel = oldDepth
        .dmDisplayFrequency = oldFrequency

    End With
        
    lRes = ChangeDisplaySettings(MiDevM, CDS_TEST)
    
End Sub

