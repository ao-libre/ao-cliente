Attribute VB_Name = "modWatching"
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public IsOnFocus As Boolean

Private MouseX       As Integer
Private MouseY       As Integer
Private LastMousePosX As Integer
Private LastMousePosY As Integer

Public Sub SendPositionMouse()

    If WatchingMe Then
    
        Dim mouse As POINTAPI
        Dim MainLeft As Long
        Dim MainTop As Long
        Dim MainWidth As Long
        Dim MainHeight As Long
        
        MainWidth = frmMain.Width / 15
        MainHeight = frmMain.Height / 15
        MainLeft = frmMain.Left / 15
        MainTop = frmMain.Top / 15
        
        GetCursorPos mouse
        
        If mouse.X > MainLeft And mouse.Y > MainTop And mouse.X < MainWidth + MainLeft And mouse.Y < MainHeight + MainTop Then
            MouseX = mouse.X - MainLeft
            MouseY = mouse.Y - MainTop
            
            If LastMousePosX = MouseX And _
                LastMousePosY = MouseY Then
                Exit Sub
            End If
            
            Call WriteWatchMouse(MouseX, MouseY, 0)
            
            LastMousePosX = MouseX
            LastMousePosY = MouseY
        End If
    End If

End Sub

