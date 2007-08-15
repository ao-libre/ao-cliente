Attribute VB_Name = "Resolution"
'**************************************************************
' Resolution.bas - Performs resolution changes.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
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
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Resolution.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20070814

Option Explicit

Private Const CCDEVICENAME As Long = 32
Private Const CCFORMNAME As Long = 32
Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const CDS_UPDATEREGISTRY As Long = &H1
Private Const CDS_TEST As Long = &H4
Private Const DISP_CHANGE_SUCCESSFUL As Long = 0
Private Const DISP_CHANGE_RESTART As Long = 1

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
Private bNoResChange As Boolean


Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long


'TODO : Change this to not depend on any external public variable using args instead!

Public Sub SetResolution()
    Dim lRes As Long
    Dim MidevM As typDevMODE
    Dim intWidth As Integer
    Dim intHeight As Integer
    Dim CambiarResolucion As Boolean
    
    lRes = EnumDisplaySettings(0, 0, MidevM)
    
    oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.Height \ Screen.TwipsPerPixelY
    
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
End Sub

Public Sub ResetResolution()
    Dim typDevM As typDevMODE
    Dim lRes As Long
    
    If Not bNoResChange Then
    
        lRes = EnumDisplaySettings(0, 0, typDevM)
        
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
        End With
        
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If
End Sub
