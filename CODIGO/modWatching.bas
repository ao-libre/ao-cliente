Attribute VB_Name = "modWatching"
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public WatchingIndex As Integer
Public IsOnFocus As Boolean

Private MouseX       As Integer
Private MouseY       As Integer
Private LastMousePosX As Integer
Private LastMousePosY As Integer

Public Sub SendPositionMouse()

    If Not WatchingMe Then Exit Sub
    
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

End Sub

Public Sub ChangeCameraPos(ByVal X As Integer, ByVal Y As Integer)

    If Not ImWatching Then Exit Sub

    Dim Direccion As E_Heading
    
    If charlist(WatchingIndex).Pos.X < X Then
        Direccion = E_Heading.EAST
    ElseIf charlist(WatchingIndex).Pos.X > X Then
        Direccion = E_Heading.WEST
    ElseIf charlist(WatchingIndex).Pos.Y < Y Then
        Direccion = E_Heading.SOUTH
    ElseIf charlist(WatchingIndex).Pos.Y > Y Then
        Direccion = E_Heading.NORTH
    Else
        Direccion = E_Heading.nada
    End If
    
    ' Quitamos anterior lugar
    If MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex Then
        MapData(UserPos.X, UserPos.Y).CharIndex = 0
    End If
    
    Call Char_MovebyHead(WatchingIndex, Direccion)
    Call Char_MoveScreen(Direccion)

    ' Colocamos en el nuevo lugar
    UserPos.X = X
    UserPos.Y = Y
    MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
    charlist(WatchingIndex).Pos = UserPos

    bTecho = Char_Techo '// Pos : Techo
    frmMain.Coord.Caption = "Map:" & UserMap & " X:" & X & " Y:" & Y
    Call frmMain.ActualizarMiniMapa

End Sub
