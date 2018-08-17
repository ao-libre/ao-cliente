Attribute VB_Name = "modConsole"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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

Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type ENLINK
    hdr As NMHDR
    msg As Long
    wParam As Long
    lParam As Long
    chrg As CHARRANGE
End Type

Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const WM_NOTIFY = &H4E
Private Const EM_SETEVENTMASK = &H445
Private Const EM_GETEVENTMASK = &H43B
Private Const EM_GETTEXTRANGE = &H44B
Private Const EM_AUTOURLDETECT = &H45B
Private Const EN_LINK = &H70B
Private Const WM_LBUTTONDOWN = &H201

Private Const ENM_LINK = &H4000000
Private Const GWL_WNDPROC = (-4)
Private Const SW_SHOW = 5

Private lOldProc As Long
Private hWndRTB As Long
Private hWndParent As Long

Public Sub EnableURLDetect(ByVal hWndRichTextbox As Long, ByVal hWndOwner As Long)
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2010
'Enables url detection in richtexbox.
'***************************************************
    SendMessage hWndRichTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWndRichTextbox, EM_GETEVENTMASK, 0, 0)
    SendMessage hWndRichTextbox, EM_AUTOURLDETECT, 1, ByVal 0
    
    hWndParent = hWndOwner
    hWndRTB = hWndRichTextbox
End Sub

Public Sub DisableURLDetect()
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2010
'Disables url detection in richtexbox.
'***************************************************
    SendMessage hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0
    StopCheckingLinks
End Sub

Public Sub StartCheckingLinks()
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2010
'Starts checking links (in console range)
'***************************************************
    If lOldProc = 0 Then
        lOldProc = SetWindowLong(hWndParent, GWL_WNDPROC, AddressOf WndProc)
    End If
End Sub

Public Sub StopCheckingLinks()
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2010
'Stops checking links (out of console range)
'***************************************************
    If lOldProc Then
        SetWindowLong hWndParent, GWL_WNDPROC, lOldProc
        lOldProc = 0
    End If
End Sub

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2010
'Get "Click" event on link and open browser.
'***************************************************
    Dim uHead As NMHDR
    Dim eLink As ENLINK
    Dim eText As TEXTRANGE
    Dim sText As String
    Dim lLen As Long
    
    If uMsg = WM_NOTIFY Then
        CopyMemory uHead, ByVal lParam, Len(uHead)
        If (uHead.hWndFrom = hWndRTB) And (uHead.code = EN_LINK) Then
            CopyMemory eLink, ByVal lParam, Len(eLink)
            
            Select Case eLink.msg
                Case WM_LBUTTONDOWN
                    eText.chrg.cpMin = eLink.chrg.cpMin
                    eText.chrg.cpMax = eLink.chrg.cpMax
                    eText.lpstrText = Space$(1024)
                    
                    lLen = SendMessage(hWndRTB, EM_GETTEXTRANGE, 0, eText)
                    
                    sText = Left$(eText.lpstrText, lLen)
                    ShellExecute hWndParent, vbNullString, sText, vbNullString, vbNullString, SW_SHOW
            End Select
        End If
    End If
    
    WndProc = CallWindowProc(lOldProc, hwnd, uMsg, wParam, lParam)
End Function

