Attribute VB_Name = "mDx8_Projectile"
Option Explicit

Public LastProjectile As Integer    'Last projectile index used

'Projectile information
Public Type Projectile
    X As Single
    Y As Single
    TX As Single
    TY As Single
    RotateSpeed As Byte
    Rotate As Single
    Grh As Grh
    OffsetX As Integer
    OffsetY As Integer
    
End Type

Public ProjectileList() As Projectile   'Holds info about all the active projectiles (arrows, ninja stars, bullets, etc)

Public Sub Engine_Projectile_Create(ByVal AttackerIndex As Integer, ByVal TargetIndex As Integer, ByVal GrhIndex As Long, ByVal Rotation As Byte, Optional ByVal Fallo As Boolean = False)
'*****************************************************************
'Creates a projectile for a ranged weapon
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Projectile_Create
'*****************************************************************
Dim ProjectileIndex As Integer

    If AttackerIndex = 0 Then Exit Sub
    If TargetIndex = 0 Then Exit Sub
    If AttackerIndex > UBound(charlist) Then Exit Sub
    If TargetIndex > UBound(charlist) Then Exit Sub

    'Get the next open projectile slot
    Do
        ProjectileIndex = ProjectileIndex + 1
        
        'Update LastProjectile if we go over the size of the current array
        If ProjectileIndex > LastProjectile Then
            LastProjectile = ProjectileIndex
            ReDim Preserve ProjectileList(1 To LastProjectile)
            Exit Do
        End If
        
    Loop While ProjectileList(ProjectileIndex).Grh.GrhIndex > 0
    
    'Figure out the initial rotation value
    ProjectileList(ProjectileIndex).Rotate = Engine_GetAngle(charlist(AttackerIndex).Pos.X, charlist(AttackerIndex).Pos.Y, charlist(TargetIndex).Pos.X, charlist(TargetIndex).Pos.Y)
    If ProjectileList(ProjectileIndex).Rotate > 224 Then ProjectileList(ProjectileIndex).Rotate = ProjectileList(ProjectileIndex).Rotate - 350
    ProjectileList(ProjectileIndex).OffsetX = 0
    ProjectileList(ProjectileIndex).OffsetY = 0
        
    'Fill in the values
    ProjectileList(ProjectileIndex).TX = (charlist(TargetIndex).Pos.X + IIf(Fallo = True, RandomNumber(-2, 2), 0)) * 32
    ProjectileList(ProjectileIndex).TY = (charlist(TargetIndex).Pos.Y + IIf(Fallo = True, RandomNumber(-2, 0), 0)) * 32
    ProjectileList(ProjectileIndex).RotateSpeed = Rotation

    If charlist(AttackerIndex).Pos.X <= 17 Then
        Select Case charlist(AttackerIndex).Pos.X
            Case 9
                ProjectileList(ProjectileIndex).OffsetX = 288
            Case 10
                ProjectileList(ProjectileIndex).OffsetX = 268
            Case 11
                ProjectileList(ProjectileIndex).OffsetX = 228
            Case 12
                ProjectileList(ProjectileIndex).OffsetX = 198
            Case 13
                ProjectileList(ProjectileIndex).OffsetX = 148
            Case 14
                ProjectileList(ProjectileIndex).OffsetX = 128
            Case 15
                ProjectileList(ProjectileIndex).OffsetX = 98
            Case 16
                ProjectileList(ProjectileIndex).OffsetX = 68
            Case 17
                ProjectileList(ProjectileIndex).OffsetX = 38
        End Select
    End If
    If charlist(AttackerIndex).Pos.Y <= 15 Then
        Select Case charlist(AttackerIndex).Pos.Y
            Case 8
                ProjectileList(ProjectileIndex).OffsetY = 258
            Case 9
                ProjectileList(ProjectileIndex).OffsetY = 228
            Case 10
                ProjectileList(ProjectileIndex).OffsetY = 198
            Case 11
                ProjectileList(ProjectileIndex).OffsetY = 148
            Case 12
                ProjectileList(ProjectileIndex).OffsetY = 128
            Case 13
                ProjectileList(ProjectileIndex).OffsetY = 98
            Case 14
                ProjectileList(ProjectileIndex).OffsetY = 68
            Case 15
                ProjectileList(ProjectileIndex).OffsetY = 38
        End Select
    End If
    
    ProjectileList(ProjectileIndex).X = charlist(AttackerIndex).Pos.X * 32
    ProjectileList(ProjectileIndex).Y = charlist(AttackerIndex).Pos.Y * 32
    
    Call InitGrh(ProjectileList(ProjectileIndex).Grh, GrhIndex)
    
End Sub

Public Sub Engine_Projectile_Erase(ByVal ProjectileIndex As Integer)
'*****************************************************************
'Erase a projectile by the projectile index
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Projectile_Erase
'*****************************************************************
    'Clear the selected index
    ProjectileList(ProjectileIndex).Grh.FrameCounter = 0
    ProjectileList(ProjectileIndex).Grh.GrhIndex = 0
    ProjectileList(ProjectileIndex).X = 0
    ProjectileList(ProjectileIndex).Y = 0
    ProjectileList(ProjectileIndex).TX = 0
    ProjectileList(ProjectileIndex).TY = 0
    ProjectileList(ProjectileIndex).Rotate = 0
    ProjectileList(ProjectileIndex).RotateSpeed = 0
 
    'Update LastProjectile
    If ProjectileIndex = LastProjectile Then
        Do Until ProjectileList(ProjectileIndex).Grh.GrhIndex > 1
            'Move down one projectile
            LastProjectile = LastProjectile - 1
            If LastProjectile = 0 Then Exit Do
        Loop
        If ProjectileIndex <> LastProjectile Then
            'We still have projectiles, resize the array to end at the last used slot
            If LastProjectile > 0 Then
                ReDim Preserve ProjectileList(1 To LastProjectile)
            Else
                Erase ProjectileList
            End If
        End If
    End If
 
End Sub


