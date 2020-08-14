Attribute VB_Name = "modRemoveTitleBar"
Option Explicit

'Remove Title Bar
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
 
' Sacado de https://www.vbforums.com/showthread.php?379880-RESOLVED-Remove-Title-Bar-Off-Of-Form-Using-API-s
' Borro algunas partes innecesarias (WyroX)
Public Sub Form_RemoveTitleBar(f As Form)
    Dim Style As Long
    ' Get window's current style bits.
    Style = GetWindowLong(f.hwnd, GWL_STYLE)
    ' Set the style bit for the title off.
    Style = Style And Not WS_CAPTION

    ' Send the new style to the window.
    SetWindowLong f.hwnd, GWL_STYLE, Style
    ' Repaint the window.
    'SetWindowPos f.hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
End Sub
