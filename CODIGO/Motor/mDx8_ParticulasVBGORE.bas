Attribute VB_Name = "mDx8_ParticulasVBGORE"
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

'Texture for particle effects - this is handled differently then the rest of the graphics
Public ParticleTexture(1 To 20) As tParticleTexture

Private Type Effect

        X As Single                     'Location of effect
        Y As Single
        GoToX As Single                 'Location to move to
        GoToY As Single
        KillWhenAtTarget As Boolean     'If the effect is at its target (GoToX/Y), then Progression is set to 0
        KillWhenTargetLost As Boolean   'Kill the effect if the target is lost (sets progression = 0)
        Gfx As Byte                     'Particle texture used
        Used As Boolean                 'If the effect is in use
        EffectNum As Byte               'What number of effect that is used
        Modifier As Integer             'Misc variable (depends on the effect)
        FloatSize As Long               'The size of the particles
        Direction As Integer            'Misc variable (depends on the effect)
        Particles() As Particle         'Information on each particle
        Progression As Single           'Progression state, best to design where 0 = effect ends
        PartVertex() As ParticleVA      'Used to point render particles
        PreviousFrame As Long           'Tick time of the last frame
        ParticleCount As Integer        'Number of particles total
        ParticlesLeft As Integer        'Number of particles left - only for non-repetitive effects
        BindToChar As Integer           'Setting this value will bind the effect to move towards the character
        BindSpeed As Single             'How fast the effect moves towards the character
        BoundToMap As Byte              'If the effect is bound to the map or not (used only by the map editor)
        r As Single
        G As Single
        B As Single
        EcuationCount As Byte
        Sng As Single                   'Misc variable

End Type

Private Type tParticleTexture
    Texture As Direct3DTexture8 'Holds the texture of the text
    TextureWidth As Integer
    TextureHeight As Integer
End Type

Private Type ParticleVA
    X As Integer
    Y As Integer
    W As Integer
    H As Integer
    
    Tx1 As Single
    Tx2 As Single
    Ty1 As Single
    Ty2 As Single
End Type

Public NumEffects                       As Byte   'Maximum number of effects at once
Public Effect()                         As Effect   'List of all the active effects

Public Enum eParticulas
        Snow = 1                        'Snow that covers the screen - weather effect
        Summon = 2                      'Summon effect
        Rain = 3                        'Exact same as snow, but moves much faster and more alpha value - weather effect
End Enum

Public Sub Engine_Init_ParticleEngine()

    '*****************************************************************
    'Loads all particles into memory - unlike normal textures, these stay in memory. This isn't
    'done for any reason in particular, they just use so little memory since they are so small
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_ParticleEngine
    '*****************************************************************

    Dim i As Long
    Dim TexInfo As D3DXIMAGE_INFO_A
    
    'Set the particles texture
    NumEffects = 25
    ReDim Effect(1 To NumEffects)

    For i = 1 To UBound(ParticleTexture())

        If Not ParticleTexture(i) Is Nothing Then
            Set ParticleTexture(i) = Nothing
        End If
        
        Set ParticleTexture(i).Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, _
                                                                    Game.path(Graficos) & "p" & i & ".png", _
                                                                    D3DX_DEFAULT, _
                                                                    D3DX_DEFAULT, _
                                                                    D3DX_DEFAULT, _
                                                                    0, _
                                                                    D3DFMT_UNKNOWN, _
                                                                    D3DPOOL_MANAGED, _
                                                                    D3DX_FILTER_POINT, _
                                                                    D3DX_FILTER_POINT, _
                                                                    &HFF000000, _
                                                                    ByVal 0, _
                                                                    ByVal 0)
                                                                    
        'Store the size of the texture
        ParticleTexture(i).TextureWidth = TexInfo.Width
        ParticleTexture(i).TextureHeight = TexInfo.Height

    Next i

End Sub

Private Function Effect_FToDW(ByVal f As Single) As Long

        '*****************************************************************
        'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
        'More info: http://www.vbgore.com/CommonCode.Particles.Effect_FToDW
        '*****************************************************************

        Dim buf As D3DXBuffer

        With DirectD3D8
            
            'Converts a single into a long (Float to DWORD)
            Set buf = .CreateBuffer(4)
            
            Call .BufferSetData(buf, 0, 4, 1, f)
            Call .BufferGetData(buf, 0, 4, 1, Effect_FToDW)
            
        End With
        
End Function

Public Sub Effect_Kill(Optional ByVal EffectIndex As Integer = 1, _
                Optional ByVal KillAll As Boolean = False)

    '*****************************************************************
    'Kills (stops) a single effect or all effects
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Kill
    '*****************************************************************

    Dim LoopC As Long

    'Check If To Kill All Effects
    If KillAll = True Then

        'Loop Through Every Effect
        For LoopC = 1 To NumEffects
        
            'Stop The Effect
            Effect(LoopC).Used = False

        Next

    Else
    
        'Stop The Selected Effect
        Effect(EffectIndex).Used = False

    End If

End Sub

Private Function Effect_NextOpenSlot() As Integer

    '*****************************************************************
    'Finds the next open effects index
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_NextOpenSlot
    '*****************************************************************

    Dim EffectIndex As Integer
    
    With Effect(EffectIndex)
    
        'Find The Next Open Effect Slot
        Do
        
            EffectIndex = EffectIndex + 1   'Check The Next Slot

            If EffectIndex > NumEffects Then    'Dont Go Over Maximum Amount
                Effect_NextOpenSlot = -1

                Exit Function

            End If

        Loop While .Used = True  'Check Next If Effect Is In Use

        'Return the next open slot
        Effect_NextOpenSlot = EffectIndex
    
        'Clear the old information from the effect
        Erase .Particles()
        Erase .PartVertex()
          
        Call ZeroMemory(Effect(EffectIndex), LenB(Effect(EffectIndex)))
    
        .GoToX = -30000
        .GoToY = -30000
    
    End With
          
End Function

Private Sub Effect_UpdateOffset(ByVal EffectIndex As Integer)
    '***************************************************
    'Update an effect's position if the screen has moved
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateOffset
    '***************************************************
    
    With Effect(EffectIndex)
    
        If UserCharIndex <> 0 Then
        
            If EffectIndex <> charlist(UserCharIndex).ParticleIndex Then
            
                .X = .X + (LastOffsetX - ParticleOffsetX)
                .Y = .Y + (LastOffsetY - ParticleOffsetY)
                
            End If
            
        End If
    
    End With
    
End Sub

Private Sub Effect_UpdateBinding(ByVal EffectIndex As Integer)
 
    '***************************************************
    'Updates the binding of a particle effect to a target, if
    'the effect is bound to a character
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateBinding
    '***************************************************
    Dim TargetI As Integer
    Dim TargetA As Single
    
    With Effect(EffectIndex)
    
        'Update position through character binding
        If .BindToChar > 0 Then
 
            'Store the character index
            TargetI = .BindToChar
 
            'Check for a valid binding index
            If TargetI > LastChar Then
                .BindToChar = 0

                If .KillWhenTargetLost Then
                    .Progression = 0
                    Exit Sub

                End If

            ElseIf charlist(TargetI).active = 0 Then
                .BindToChar = 0

                If .KillWhenTargetLost Then
                    .Progression = 0
                    Exit Sub

                End If

            Else
 
                'Calculate the X and Y positions
                .GoToX = Engine_TPtoSPX(charlist(.BindToChar).Pos.X)
                .GoToY = Engine_TPtoSPY(charlist(.BindToChar).Pos.Y)
 
            End If
 
        End If
 
        'Move to the new position if needed
        If .GoToX > -30000 Or .GoToY > -30000 Then
            If .GoToX <> .X Or .GoToY <> .Y Then
 
                'Calculate the angle
                TargetA = Engine_GetAngle(.X, .Y, .GoToX, .GoToY) + 180
 
                'Update the position of the effect
                .X = .X - Sin(TargetA * DegreeToRadian) * .BindSpeed
                .Y = .Y + Cos(TargetA * DegreeToRadian) * .BindSpeed
 
                'Check if the effect is close enough to the target to just stick it at the target
                If .GoToX > -30000 Then
                    If Abs(.X - .GoToX) < 6 Then .X = .GoToX
                End If

                If .GoToY > -30000 Then
                    If Abs(.Y - .GoToY) < 6 Then .Y = .GoToY
                End If
 
                'Check if the position of the effect is equal to that of the target
                If .X = .GoToX Then
                    If .Y = .GoToY Then
 
                        'For some effects, if the position is reached, we want to end the effect
                        If .KillWhenAtTarget Then
                            .BindToChar = 0
                            .Progression = 0
                            .GoToX = .X
                            .GoToY = .Y
                        End If

                        Exit Sub    'The effect is at the right position, don't update
 
                    End If

                End If
 
            End If

        End If
    
    End With
    
End Sub

Public Sub Effect_Render(ByVal EffectIndex As Integer, Optional ByVal SetRenderStates As Boolean = True)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Render
'*****************************************************************

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Set the render state for the size of the particle
    Call DirectDevice.SetRenderState(D3DRS_POINTSIZE, Effect(EffectIndex).FloatSize)
    
    'Set the render state to point blitting
    If SetRenderStates Then
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    With Effect(EffectIndex)
    
    'Set the texture
    Call SpriteBatch.SetTexture(ParticleTexture(Effect(EffectIndex).Gfx))

    'Draw all the particles at once
    With Effect(EffectIndex).PartVertex(0)
    
        Call SpriteBatch.Draw(.X, .Y, , , .Color(), .tu, .tv, , , True, , D3DPT_POINTLIST)
    
    End With
    

    'Reset the render state back to normal
    If SetRenderStates Then DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Sub Effect_UpdateAll()

    '*****************************************************************
    'Updates all of the effects and renders them
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateAll
    '*****************************************************************

    Dim LoopC As Long
    
    'Make sure the particle engine is on.
    If Not ClientSetup.ParticleEngine Then Exit Sub
    
    'Make sure we have effects
    If NumEffects = 0 Then Exit Sub
    
    'Set the render state for the particle effects
    Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_ONE)

    'Update every effect in use
    For LoopC = 1 To NumEffects

        'Make sure the effect is in use
        If Effect(LoopC).Used Then
            
            'Update the effect position if the screen has moved
            Call Effect_UpdateOffset(LoopC)
            
            'Update the effect position if it is binded
            Call Effect_UpdateBinding(LoopC)
            
            'Find out which effect is selected, then update it
            Select Case Effect(LoopC).EffectNum
            
                Case eParticulas.Rain
                    Call Effect_Rain_Update(LoopC)
                    
                Case eParticulas.Snow
                    Call Effect_Snow_Update(LoopC)
                    
                Case eParticulas.Summon
                    Call Effect_Summon_Update(LoopC)
                    
            End Select
            
            'Actualizamos el Clima siempre.
            Call Engine_Weather_Update
            
            'Render the effect
            Call Effect_Render(LoopC, False)

        End If

    Next

    'Set the render state back for normal rendering
    Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)

End Sub

Public Sub Effect_Begin(ByVal EffectIndex As eParticulas, _
                        ByVal X As Single, _
                        ByVal Y As Single, _
                        ByVal GfxIndex As Byte, _
                        ByVal Particles As Byte, _
                        Optional ByVal Direction As Single = 180, _
                        Optional ByVal BindToMap As Boolean = False)

        '*****************************************************************
        'A very simplistic form of initialization for particle effects
        'Should only be used for starting MAP-BASED effects
        '*****************************************************************

        Dim RetNum As Byte

        Select Case EffectIndex

                'Case 4
                    'Call Effect_Waterfall_Begin(X, Y, GfxIndex, Particles)
                        
                'Case 7
                    'Call Effect_Portal_Begin(X, Y, GfxIndex, Particles, 100)
                    
        End Select

        'Bind the effect to the map if needed
        If BindToMap Then Effect(RetNum).BoundToMap = 1
        
End Sub

Public Sub Effect_Create(ByVal CharIndex As Integer, ByVal Effect As eParticulas, ByVal X As Single, ByVal Y As Single, ByVal Loops As Integer)
    
        '*****************************************************************
        'A very simplistic form of initialization for particle effects
        'Should only be used for starting CHAR-BASED effects
        '*****************************************************************
    
    Select Case Effect
    
        Case eParticulas.Summon
            charlist(CharIndex).ParticleIndex = Effect_Summon_Begin(X, Y, 1, 500, 0.1)

    End Select

End Sub
