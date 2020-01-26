Attribute VB_Name = "mDx8_Auras"
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 26/05/10
'Blisse-AO | Sistema de Auras
'***************************************************

Option Explicit

Public Type Aura
    Grh As Long '   GrhIndex
    
    Rotation As Byte '   Rotate or Not
    angle As Single '   Angle
    Speed As Single '   Speed
    TickCount As Long '   TickCount from Speed Controls
    
    Color(0 To 3) As Long '   Color
    
    OffsetX As Integer '   PixelOffset X
    OffsetY As Integer '   PixelOffset Y
End Type

Public Auras() As Aura '   List of Aura's

Public Sub SetCharacterAura(ByVal CharIndex As Integer, ByVal AuraIndex As Byte, ByVal slot As Byte)
'***************************************************
'Author: Standelf
'Last Modify Date: 27/05/2010
'***************************************************
    If slot <= 0 Or slot >= 5 Then Exit Sub
    Set_Aura CharIndex, slot, AuraIndex
End Sub

Public Sub Load_Auras()
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 26/05/10
'Load Auras
'***************************************************
    Dim i As Integer, AurasTotales As Integer, Leer As New clsIniManager
    Leer.Initialize Game.path(INIT) & "auras.ini"

    AurasTotales = Val(Leer.GetValue("Auras", "NumAuras"))
    
    ReDim Preserve Auras(1 To AurasTotales)
    
            For i = 1 To AurasTotales
                Auras(i).Grh = Val(Leer.GetValue(i, "GrhIndex"))
                
                Auras(i).Rotation = Val(Leer.GetValue(i, "Rotate"))
                Auras(i).angle = 0
                Auras(i).Speed = Leer.GetValue(i, "Speed")
                
                Auras(i).OffsetX = Val(Leer.GetValue(i, "OffsetX"))
                Auras(i).OffsetY = Val(Leer.GetValue(i, "OffsetY"))

            Dim ColorSet As Byte, TempSet As String
            
            For ColorSet = 0 To 3
                TempSet = Leer.GetValue(Val(i), "Color" & ColorSet)
                Auras(i).Color(ColorSet) = D3DColorXRGB(ReadField(1, TempSet, Asc(",")), ReadField(2, TempSet, Asc(",")), ReadField(3, TempSet, Asc(",")))
            Next ColorSet
                
                Auras(i).TickCount = 0
            Next i
                                                         
    Set Leer = Nothing
End Sub

Public Sub DeInit_Auras()
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 26/05/10
'DeInit Auras
'***************************************************
    '   Erase Data
    Erase Auras()
    
    '   Finish
    Exit Sub
End Sub

Public Sub Set_Aura(ByVal CharIndex As Integer, slot As Byte, Aura As Byte)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 26/05/10
'Set Aura to Char
'***************************************************
    If slot <= 0 Or slot >= 5 Then Exit Sub
    
    With charlist(CharIndex).Aura(slot)
        .Grh = Auras(Aura).Grh
            
        .angle = Auras(Aura).angle
        .Rotation = Auras(Aura).Rotation
        .Speed = Auras(Aura).Speed
        
        .OffsetX = Auras(Aura).OffsetX
        .OffsetY = Auras(Aura).OffsetY
        
        .Color(0) = Auras(Aura).Color(0)
        .Color(1) = Auras(Aura).Color(1)
        .Color(2) = Auras(Aura).Color(2)
        .Color(3) = Auras(Aura).Color(3)
        
        .TickCount = GetTickCount
    End With
End Sub

Public Sub Delete_All_Auras(ByVal CharIndex As Integer)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 26/05/10
'Kill all of aura´s from Char
'***************************************************
    Delete_Aura CharIndex, 1
    Delete_Aura CharIndex, 2
    Delete_Aura CharIndex, 3
    Delete_Aura CharIndex, 4
End Sub
    
Public Sub Delete_Aura(ByVal CharIndex As Integer, slot As Byte)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 26/05/10
'Kill Aura from Char
'***************************************************
    If slot <= 0 Or slot >= 5 Then Exit Sub
    
    charlist(CharIndex).Aura(slot) = Auras(1) '1 = Fake Aura
End Sub

Public Sub Update_Aura(ByVal CharIndex As Integer, slot As Byte)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 26/05/10
'Update Angle of Aura
'***************************************************
    If slot <= 0 Or slot >= 5 Then Exit Sub
    
    With charlist(CharIndex).Aura(slot)
        If GetTickCount - .TickCount > FPS Then
            .angle = .angle + .Speed
            If .angle >= 360 Then .angle = 0
            .TickCount = GetTickCount
        End If
    End With
End Sub

Public Sub Render_Auras(ByVal CharIndex As Integer, X As Integer, Y As Integer)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 26/05/10
'Render the Auras from a Char
'***************************************************
On Error GoTo handle
    Dim i As Byte
        For i = 1 To 4
            With charlist(CharIndex).Aura(i)
                If .Grh <> 0 Then
                    If .Rotation = 1 Then Update_Aura CharIndex, i
                    Call Draw_GrhIndex(.Grh, X + .OffsetX, Y + .OffsetY, 1, .Color(), .angle, True)
                End If
            End With
        Next i
handle:
    Exit Sub
End Sub



