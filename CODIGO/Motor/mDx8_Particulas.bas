Attribute VB_Name = "mDx8_Particulas"
'*************************************************************
'ImperiumAO 1.4.6
'*************************************************************
'Este modulo contiene TODOS los procedimientos que conforma
'el Sistema de Particulas ORE.
'*************************************************************

Option Explicit

Public Type RGB
    r As Long
    g As Long
    b As Long
End Type

Public Type Stream
    name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    alphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    speed As Single
    life_counter As Long
End Type

Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
End Type

Private Type Particle_Group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    Particle_Count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alphaBlend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    live As Long
    liv1 As Integer
    liveend As Long
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    
    'Added by Juan Martin Sotuyo Dodero
    speed As Single
    life_counter As Long
End Type

Dim particle_group_list() As Particle_Group
Dim particle_group_count As Long
Dim particle_group_last As Long

Public TotalStreams As Integer
Public StreamData() As Stream

Public Const PI As Single = 3.14159265358979

Private RainParticle As Long

Public Enum eWeather
    Rain
    Snow
End Enum

Public Sub CargarParticulas()
    Dim LoopC As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim Leer As New clsIniManager

    Call Leer.Initialize(path(INIT) & "Particulas.ini")

    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'fill StreamData array with info from Particles.ini
    For LoopC = 1 To TotalStreams
        With StreamData(LoopC)
            .name = Leer.GetValue(Val(LoopC), "Name")
            .NumOfParticles = Leer.GetValue(Val(LoopC), "NumOfParticles")
            .x1 = Leer.GetValue(Val(LoopC), "X1")
            .y1 = Leer.GetValue(Val(LoopC), "Y1")
            .x2 = Leer.GetValue(Val(LoopC), "X2")
            .y2 = Leer.GetValue(Val(LoopC), "Y2")
            .angle = Leer.GetValue(Val(LoopC), "Angle")
            .vecx1 = Leer.GetValue(Val(LoopC), "VecX1")
            .vecx2 = Leer.GetValue(Val(LoopC), "VecX2")
            .vecy1 = Leer.GetValue(Val(LoopC), "VecY1")
            .vecy2 = Leer.GetValue(Val(LoopC), "VecY2")
            .life1 = Leer.GetValue(Val(LoopC), "Life1")
            .life2 = Leer.GetValue(Val(LoopC), "Life2")
            .friction = Leer.GetValue(Val(LoopC), "Friction")
            .spin = Leer.GetValue(Val(LoopC), "Spin")
            .spin_speedL = Leer.GetValue(Val(LoopC), "Spin_SpeedL")
            .spin_speedH = Leer.GetValue(Val(LoopC), "Spin_SpeedH")
            .alphaBlend = Leer.GetValue(Val(LoopC), "AlphaBlend")
            .gravity = Leer.GetValue(Val(LoopC), "Gravity")
            .grav_strength = Leer.GetValue(Val(LoopC), "Grav_Strength")
            .bounce_strength = Leer.GetValue(Val(LoopC), "Bounce_Strength")
            .XMove = Leer.GetValue(Val(LoopC), "XMove")
            .YMove = Leer.GetValue(Val(LoopC), "YMove")
            .move_x1 = Leer.GetValue(Val(LoopC), "move_x1")
            .move_x2 = Leer.GetValue(Val(LoopC), "move_x2")
            .move_y1 = Leer.GetValue(Val(LoopC), "move_y1")
            .move_y2 = Leer.GetValue(Val(LoopC), "move_y2")
            .life_counter = Leer.GetValue(Val(LoopC), "life_counter")
            .speed = Val(Leer.GetValue(Val(LoopC), "Speed"))
            
            .NumGrhs = Leer.GetValue(Val(LoopC), "NumGrhs")
            
            ReDim .grh_list(1 To .NumGrhs)
            GrhListing = Leer.GetValue(Val(LoopC), "Grh_List")
            
            For i = 1 To .NumGrhs
                .grh_list(i) = ReadField(i, GrhListing, Asc(","))
            Next i
            
            .grh_list(i - 1) = .grh_list(i - 1)
            
            For ColorSet = 1 To 4
                TempSet = Leer.GetValue(Val(LoopC), "ColorSet" & ColorSet)
                .colortint(ColorSet - 1).r = ReadField(1, TempSet, Asc(","))
                .colortint(ColorSet - 1).g = ReadField(2, TempSet, Asc(","))
                .colortint(ColorSet - 1).b = ReadField(3, TempSet, Asc(","))
            Next ColorSet

        End With
    Next LoopC
    
    Set Leer = Nothing

End Sub

Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, _
                                             ByVal char_index As Integer, _
                                             Optional ByVal particle_life As Long = 0) As Long

    Dim rgb_list(0 To 3) As Long

    With StreamData(ParticulaInd)
        rgb_list(0) = RGB(.colortint(0).r, .colortint(0).g, .colortint(0).b)
        rgb_list(1) = RGB(.colortint(1).r, .colortint(1).g, .colortint(1).b)
        rgb_list(2) = RGB(.colortint(2).r, .colortint(2).g, .colortint(2).b)
        rgb_list(3) = RGB(.colortint(3).r, .colortint(3).g, .colortint(3).b)

        General_Char_Particle_Create = Char_Particle_Group_Create(char_index, .grh_list, rgb_list(), .NumOfParticles, ParticulaInd, .alphaBlend, IIf(particle_life = 0, .life_counter, particle_life), .speed, , .x1, .y1, .angle, .vecx1, .vecx2, .vecy1, .vecy2, .life1, .life2, .friction, .spin_speedL, .gravity, .grav_strength, .bounce_strength, .x2, .y2, .XMove, .move_x1, .move_x2, .move_y1, .move_y2, .YMove, .spin_speedH, .spin)

    End With

End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, _
                                        ByVal X As Integer, _
                                        ByVal Y As Integer, _
                                        Optional ByVal particle_life As Long = 0) As Long

    Dim rgb_list(0 To 3) As Long

    With StreamData(ParticulaInd)
        rgb_list(0) = RGB(.colortint(0).r, .colortint(0).g, .colortint(0).b)
        rgb_list(1) = RGB(.colortint(1).r, .colortint(1).g, .colortint(1).b)
        rgb_list(2) = RGB(.colortint(2).r, .colortint(2).g, .colortint(2).b)
        rgb_list(3) = RGB(.colortint(3).r, .colortint(3).g, .colortint(3).b)
    
        General_Particle_Create = Particle_Group_Create(X, Y, .grh_list, rgb_list(), .NumOfParticles, ParticulaInd, .alphaBlend, IIf(particle_life = 0, .life_counter, particle_life), .speed, , .x1, .y1, .angle, .vecx1, .vecx2, .vecy1, .vecy2, .life1, .life2, .friction, .spin_speedL, .gravity, .grav_strength, .bounce_strength, .x2, .y2, .XMove, .move_x1, .move_x2, .move_y1, .move_y2, .YMove, .spin_speedH, .spin)

    End With

End Function

Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, _
                                           ByVal stream_type As Long)

    '**************************************************************
    'Author: Augusto Jos� Rando
    '**************************************************************
    Dim char_part_index As Integer

    If Char_Check(char_index) Then
    
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)

        If char_part_index = -1 Then Exit Function
        
        Call Particle_Group_Remove(char_part_index)

    End If

End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto Jose Rando
'**************************************************************
    Dim i As Integer
    
    If Char_Check(char_index) And Not charlist(char_index).Particle_Count = 0 Then
        For i = 1 To UBound(charlist(char_index).Particle_Group)
            If charlist(char_index).Particle_Group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).Particle_Group(i))
        Next i
        Erase charlist(char_index).Particle_Group
        charlist(char_index).Particle_Count = 0
    End If
    
End Function

Public Function Particle_Group_Remove(ByVal Particle_Group_Index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(Particle_Group_Index) Then
        Particle_Group_Destroy Particle_Group_Index
        Particle_Group_Remove = True
    End If
End Function

Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
End Function

Public Sub Particle_Group_Render(ByVal Particle_Group_Index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martin Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim LoopC As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    If Particle_Group_Index > UBound(particle_group_list) Then Exit Sub
    
    If GetTickCount - particle_group_list(Particle_Group_Index).live > (particle_group_list(Particle_Group_Index).liv1 * 25) And Not particle_group_list(Particle_Group_Index).liv1 = -1 Then
        Call Particle_Group_Destroy(Particle_Group_Index)
        Exit Sub
    End If
        
    With particle_group_list(Particle_Group_Index)
    
        'Set colors
        temp_rgb(0) = .rgb_list(0)
        temp_rgb(1) = .rgb_list(1)
        temp_rgb(2) = .rgb_list(2)
        temp_rgb(3) = .rgb_list(3)

        'See if it is time to move a particle
        .frame_counter = .frame_counter + timerTicksPerFrame
        If .frame_counter > .frame_speed Then
            .frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
            
        'If it's still alive render all the particles inside
        For LoopC = 1 To .Particle_Count
                
        'Render particle
            Particle_Render .particle_stream(LoopC), _
                        screen_x, screen_y, _
                        .grh_index_list(Round(RandomNumber(1, .grh_index_count), 0)), _
                        temp_rgb(), _
                        .alphaBlend, no_move, _
                        .x1, .y1, .angle, _
                        .vecx1, .vecx2, _
                        .vecy1, .vecy2, _
                        .life1, .life2, _
                        .fric, .spin_speedL, _
                        .gravity, .grav_strength, _
                        .bounce_strength, .x2, _
                        .y2, .XMove, _
                        .move_x1, .move_x2, _
                        .move_y1, .move_y2, _
                        .YMove, .spin_speedH, _
                        .spin
        Next LoopC
                
        If no_move = False Then
            'Update the group alive counter
            If .never_die = False Then
                .alive_counter = .alive_counter - 1
            End If
        End If
        
    End With
    
End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alphaBlend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martin Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************

    With temp_particle
    
        If no_move = False Then
        
            If .alive_counter = 0 Then
            
                'Start new particle
                Call InitGrh(.Grh, grh_index)
                .X = RandomNumber(x1, x2) - 16
                .Y = RandomNumber(y1, y2) - 16
                .vector_x = RandomNumber(vecx1, vecx2)
                .vector_y = RandomNumber(vecy1, vecy2)
                .alive_counter = RandomNumber(life1, life2)
                .friction = fric
                
            Else
                
                'Continue old particle
                'Do gravity
                If gravity = True Then
                    
                    .vector_y = .vector_y + grav_strength
                    
                    If .Y > 0 Then
                        'bounce
                        .vector_y = bounce_strength
                    End If
                    
                End If
                
                'Do rotation
                If spin Then .angle = .angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
                If .angle >= 360 Then
                    .angle = 0
                End If
                
                If XMove = True Then .vector_x = RandomNumber(move_x1, move_x2)
                If YMove = True Then .vector_y = RandomNumber(move_y1, move_y2)
            End If
            
            'Add in vector
            .X = .X + (.vector_x \ .friction)
            .Y = .Y + (.vector_y \ .friction)
        
            'decrement counter
             .alive_counter = .alive_counter - 1
        End If
        
        'Draw it
        If .Grh.GrhIndex Then
            Call Draw_Grh(.Grh, .X + screen_x, .Y + screen_y, 1, rgb_list(), 1, True, .angle)
        End If
        
    End With
    
End Sub

Private Function Particle_Group_Next_Open() As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim LoopC As Long
    
    If particle_group_last = 0 Then
        Particle_Group_Next_Open = 1
        Exit Function
    End If
    
    LoopC = 1

    Do Until particle_group_list(LoopC).active = False

        If LoopC = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If

        LoopC = LoopC + 1
    Loop
    
    Particle_Group_Next_Open = LoopC
    
    Exit Function
    
ErrorHandler:
    Particle_Group_Next_Open = 1

End Function

Private Function Particle_Group_Check(ByVal Particle_Group_Index As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    
    'check index
    If Particle_Group_Index > 0 And Particle_Group_Index <= particle_group_last Then
        If particle_group_list(Particle_Group_Index).active Then
            Particle_Group_Check = True
        End If
    End If

End Function

Private Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal Particle_Count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alphaBlend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'Returns the particle_group_index if successful, else 0
'Modified by Juan Martin Sotuyo Dodero
'Modified by Augusto Jose Rando
'**************************************************************
    
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Call Particle_Group_Make(Particle_Group_Create, map_x, map_y, Particle_Count, stream_type, grh_index_list(), rgb_list(), alphaBlend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin)
        End If
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Call Particle_Group_Make(Particle_Group_Create, map_x, map_y, Particle_Count, stream_type, grh_index_list(), rgb_list(), alphaBlend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin)
    End If

End Function

Private Function Particle_Group_Find(ByVal id As Long) As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    'Find the index related to the handle
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim LoopC As Long
        LoopC = 1

    Do Until particle_group_list(LoopC).id = id

        If LoopC = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If

        LoopC = LoopC + 1
        
    Loop
    
    Particle_Group_Find = LoopC
    
    Exit Function
    
ErrorHandler:
    Particle_Group_Find = 0

End Function

Private Function Particle_Get_Type(ByVal Particle_Group_Index As Long) As Byte

    On Error GoTo ErrorHandler:
    
    Particle_Get_Type = particle_group_list(Particle_Group_Index).stream_type
    
    Exit Function
    
ErrorHandler:
    Particle_Get_Type = 0

End Function

Private Sub Particle_Group_Destroy(ByVal Particle_Group_Index As Long)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    On Error Resume Next

    Dim temp As Particle_Group
    Dim i    As Integer

    With particle_group_list(Particle_Group_Index)

        If .map_x > 0 And .map_y > 0 Then
            MapData(.map_x, .map_y).Particle_Group_Index = 0
        ElseIf .char_index Then

            If Char_Check(.char_index) Then

                For i = 1 To charlist(.char_index).Particle_Count

                    If charlist(.char_index).Particle_Group(i) = Particle_Group_Index Then
                        charlist(.char_index).Particle_Group(i) = 0
                        Exit For

                    End If

                Next i

            End If

        End If

    End With

    particle_group_list(Particle_Group_Index) = temp
    
    'Update array size
    If Particle_Group_Index = particle_group_last Then

        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1

            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub

            End If

        Loop
        
        ReDim Preserve particle_group_list(1 To particle_group_last) As Particle_Group

    End If

    particle_group_count = particle_group_count - 1

End Sub

Private Sub Particle_Group_Make(ByVal Particle_Group_Index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal Particle_Count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alphaBlend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martin Sotuyo Dodero
'*****************************************************************
    'Update array size
    If Particle_Group_Index > particle_group_last Then
        particle_group_last = Particle_Group_Index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    
    With particle_group_list(Particle_Group_Index)
        'Make active
        .active = True
        
        'Map pos
        If (map_x <> -1) And (map_y <> -1) Then
            .map_x = map_x
            .map_y = map_y
        End If
        
        'Grh list
        ReDim .grh_index_list(1 To UBound(grh_index_list))
        .grh_index_list() = grh_index_list()
        .grh_index_count = UBound(grh_index_list)
        
        'Sets alive vars
        If alive_counter = -1 Then
            .alive_counter = -1
            .liv1 = -1
            .never_die = True
        Else
            .alive_counter = alive_counter
            .liv1 = alive_counter
            .never_die = False
        End If
        
        'alpha blending
        .alphaBlend = alphaBlend
        
        'stream type
        .stream_type = stream_type
        
        'speed
        .frame_speed = frame_speed
        
        .x1 = x1
        .y1 = y1
        .x2 = x2
        .y2 = y2
        .angle = angle
        .vecx1 = vecx1
        .vecx2 = vecx2
        .vecy1 = vecy1
        .vecy2 = vecy2
        .life1 = life1
        .life2 = life2
        .fric = fric
        .spin = spin
        .spin_speedL = spin_speedL
        .spin_speedH = spin_speedH
        .gravity = gravity
        .grav_strength = grav_strength
        .bounce_strength = bounce_strength
        .XMove = XMove
        .YMove = YMove
        .move_x1 = move_x1
        .move_x2 = move_x2
        .move_y1 = move_y1
        .move_y2 = move_y2
        
        .rgb_list(0) = rgb_list(0)
        .rgb_list(1) = rgb_list(1)
        .rgb_list(2) = rgb_list(2)
        .rgb_list(3) = rgb_list(3)
        
        'handle
        .id = id
        
        .live = GetTickCount()
        
        'create particle stream
        .Particle_Count = Particle_Count
        ReDim .particle_stream(1 To Particle_Count)
    
    End With
    
    'plot particle group on map
    If (map_x <> -1 And map_x <> 0) And (map_y <> -1 And map_x <> 0) Then
        MapData(map_x, map_y).Particle_Group_Index = Particle_Group_Index
    End If
    
End Sub
Private Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).Particle_Group_Index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function

Private Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal Particle_Count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alphaBlend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
    Dim char_part_free_index As Integer
    
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, Particle_Count, stream_type, grh_index_list(), rgb_list(), alphaBlend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin
    End If

End Function

Private Function Char_Particle_Group_Find(ByVal char_index As Integer, _
                                          ByVal stream_type As Long) As Integer

    '*****************************************************************
    'Author: Augusto Jos� Rando
    'Modified: returns slot or -1
    '*****************************************************************
    On Error Resume Next

    Dim i As Integer

    For i = 1 To charlist(char_index).Particle_Count

        If particle_group_list(charlist(char_index).Particle_Group(i)).stream_type = stream_type Then
            Char_Particle_Group_Find = charlist(char_index).Particle_Group(i)
            Exit Function

        End If

    Next i

    Char_Particle_Group_Find = -1

End Function

Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer

    '*****************************************************************
    'Author: Augusto Jose Rando
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim LoopC As Long
    
    If charlist(char_index).Particle_Count = 0 Then
        Char_Particle_Group_Next_Open = charlist(char_index).Particle_Count + 1
        charlist(char_index).Particle_Count = Char_Particle_Group_Next_Open
        ReDim Preserve charlist(char_index).Particle_Group(1 To Char_Particle_Group_Next_Open) As Long
        Exit Function

    End If
    
    LoopC = 1

    Do Until charlist(char_index).Particle_Group(LoopC) = 0

        If LoopC = charlist(char_index).Particle_Count Then
            Char_Particle_Group_Next_Open = charlist(char_index).Particle_Count + 1
            charlist(char_index).Particle_Count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).Particle_Group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function

        End If

        LoopC = LoopC + 1
    Loop
    
    Char_Particle_Group_Next_Open = LoopC

    Exit Function

ErrorHandler:
    charlist(char_index).Particle_Count = 1
    ReDim charlist(char_index).Particle_Group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1

End Function

Private Function Char_Check(ByVal char_index As Integer) As Boolean

    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)
    End If
    
End Function

Private Sub Char_Particle_Group_Make(ByVal Particle_Group_Index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal Particle_Count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alphaBlend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martin Sotuyo Dodero
'*****************************************************************
    'Update array size
    If Particle_Group_Index > particle_group_last Then
        particle_group_last = Particle_Group_Index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    With particle_group_list(Particle_Group_Index)
        
        'Make active
        .active = True
        
        'Char index
        .char_index = char_index
        
        'Grh list
        ReDim .grh_index_list(1 To UBound(grh_index_list))
        .grh_index_list() = grh_index_list()
        .grh_index_count = UBound(grh_index_list)
        
        'Sets alive vars
        If alive_counter = -1 Then
            .alive_counter = -1
            .liv1 = -1
            .never_die = True
        Else
            .alive_counter = alive_counter
            .liv1 = alive_counter
            .never_die = False
        End If
        
        'alpha blending
        .alphaBlend = alphaBlend
        
        'stream type
        .stream_type = stream_type
        
        'speed
        .frame_speed = frame_speed
        
        .x1 = x1
        .y1 = y1
        .x2 = x2
        .y2 = y2
        .angle = angle
        .vecx1 = vecx1
        .vecx2 = vecx2
        .vecy1 = vecy1
        .vecy2 = vecy2
        .life1 = life1
        .life2 = life2
        .fric = fric
        .spin = spin
        .spin_speedL = spin_speedL
        .spin_speedH = spin_speedH
        .gravity = gravity
        .grav_strength = grav_strength
        .bounce_strength = bounce_strength
        .XMove = XMove
        .YMove = YMove
        .move_x1 = move_x1
        .move_x2 = move_x2
        .move_y1 = move_y1
        .move_y2 = move_y2
        
        .rgb_list(0) = rgb_list(0)
        .rgb_list(1) = rgb_list(1)
        .rgb_list(2) = rgb_list(2)
        .rgb_list(3) = rgb_list(3)
        
        'handle
        .id = id
        .live = GetTickCount()
        
        'create particle stream
        .Particle_Count = Particle_Count
        ReDim .particle_stream(1 To Particle_Count)
    
    End With
    
    'plot particle group on char
    charlist(char_index).Particle_Group(particle_char_index) = Particle_Group_Index
End Sub

Public Sub Engine_Weather_Update()
'*****************************************************************
'Author: Lucas Recoaro (Recox)
'Last Modify Date: 19/12/2019
'Controla los climas, aqui se renderizan la lluvia, nieve, etc.
'*****************************************************************
    'TODO: Hay un bug no muy importante que hace que no se renderice la lluvia
    'en caso que empiece a llover, tiro el comando /salir y vuelvo a entrar al juego
    'Sin embargo al cambiar de mapa o al entrar y salir de un techo la particula se vuelve a cargar
    'Este error NO pasa cuando esta lloviendo y recien abro el juego y entro, en ese caso la lluvia se ve bien (Recox)

    If bRain And MapDat.zone <> "DUNGEON" And Not bTecho Then
        'Primero verificamos que las particulas de lluvia esten creadas en la coleccion de particulas
        'Si estan creadas las renderizamos, sino las creamos
        If RainParticle <= 0 Then
            'Creamos las particulas de lluvia
            Call mDx8_Particulas.LoadWeatherParticles(eWeather.Rain)
        ElseIf RainParticle > 0 Then
            Call mDx8_Particulas.Particle_Group_Render(RainParticle, 250, -1)
        End If
    Else
        'Borramos las particulas de lluvia en caso de que pare la lluvia o nos escondamos en un techo
        Call mDx8_Particulas.RemoveWeatherParticles(eWeather.Rain)
    End If

End Sub

Public Sub LoadWeatherParticles(ByVal Weather As Byte)
'*****************************************************************
'Author: Lucas Recoaro (Recox)
'Last Modify Date: 19/12/2019
'Crea las particulas de clima.
'*****************************************************************
    Select Case Weather

        Case eWeather.Rain
            RainParticle = mDx8_Particulas.General_Particle_Create(8, -1, -1)

    End Select
End Sub

Public Sub RemoveWeatherParticles(ByVal Weather As Byte)
'*****************************************************************
'Author: Lucas Recoaro (Recox)
'Last Modify Date: 19/12/2019
'Remueve las particulas de clima.
'*****************************************************************
    Select Case Weather

        Case eWeather.Rain
            Particle_Group_Remove (RainParticle)
            RainParticle = 0

    End Select
End Sub
