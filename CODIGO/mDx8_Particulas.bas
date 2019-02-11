Attribute VB_Name = "mDx8_Particulas"
Option Explicit

Public ParticleTexture(1 To 15) As Direct3DTexture8

Public Const DegreeToRadian     As Single = 0.01745329251994 'Pi / 180

Public WeatherEffectIndex       As Integer

Public OnRampage                As Long
Public OnRampageImg             As Long
Public OnRampageImgGrh          As Integer
Public LastEffect               As Integer

Private Type Effect

    X As Single
    Y As Single
    GoToX As Single
    GoToY As Single
    KillWhenAtTarget As Boolean     'If the effect is at its target (GoToX/Y), then Progression is set to 0
    KillWhenTargetLost As Boolean   'Kill the effect if the target is lost (sets progression = 0)
    Gfx As Byte                 'Particle texture used
    Used As Boolean             'If the effect is in use
    EffectNum As Byte           'What number of effect that is used
    Modifier As Integer         'Misc variable (depends on the effect)
    FloatSize As Long           'The size of the particles
    Direction As Integer        'Misc variable (depends on the effect)
    Particles() As Particle     'Information on each particle
    Progression As Single       'Progression state, best to design where 0 = effect ends
    PartVertex() As TLVERTEX    'Used to point render particles
    PreviousFrame As Long       'Tick time of the last frame
    ParticleCount As Integer    'Number of particles total
    ParticlesLeft As Integer    'Number of particles left - only for non-repetitive effects
    BindToChar As Integer       'Setting this value will bind the effect to move towards the character
    BindSpeed As Single         'How fast the effect moves towards the character
    BoundToMap As Byte          'If the effect is bound to the map or not (used only by the map editor)

End Type

Public NumEffects                       As Byte   'Maximum number of effects at once
Public Effect()                         As Effect   'List of all the active effects

Public Const EffectNum_Fire             As Byte = 1             'Burn baby, burn! Flame from a central point that blows in a specified direction
Public Const EffectNum_Snow             As Byte = 2             'Snow that covers the screen - weather effect
Public Const EffectNum_Heal             As Byte = 3             'Healing effect that can bind to a character, ankhs float up and fade
Public Const EffectNum_Bless            As Byte = 4            'Following three effects are same: create a circle around the central point
Public Const EffectNum_Protection       As Byte = 5       '   (often the character) and makes the given particle on the perimeter
Public Const EffectNum_Strengthen       As Byte = 6       '   which float up and fade out
Public Const EffectNum_Rain             As Byte = 7             'Exact same as snow, but moves much faster and more alpha value - weather effect
Public Const EffectNum_EquationTemplate As Byte = 8 'Template for creating particle effects through equations - a page with some equations can be found here: http://www.vbgore.com/modules.php?name=Forums&file=viewtopic&t=221
Public Const EffectNum_Waterfall        As Byte = 9        'Waterfall effect
Public Const EffectNum_Summon           As Byte = 10          'Summon effect
Public Const EffectNum_Rayo             As Byte = 11
Public Const EffectNum_Smoke            As Byte = 12

'Effect Spell
Public SpellGrhIndex                    As Integer

Private Declare Sub ZeroMemory _
                Lib "kernel32.dll" _
                Alias "RtlZeroMemory" (ByRef destination As Any, _
                                       ByVal length As Long)

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Sub Engine_Init_ParticleEngine(Optional ByVal SkipToTextures As Boolean = False)
    '*****************************************************************
    'Loads all particles into memory - unlike normal textures, these stay in memory. This isn't
    'done for any reason in particular, they just use so little memory since they are so small
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_ParticleEngine
    '*****************************************************************
    Dim i As Byte

    If Not SkipToTextures Then
        'Set the particles texture
        NumEffects = 25
        ReDim Effect(1 To NumEffects)
    
    End If
    
    For i = 1 To UBound(ParticleTexture())

        If ParticleTexture(i) Is Nothing Then Set ParticleTexture(i) = Nothing
        Set ParticleTexture(i) = DirectD3D8.CreateTextureFromFileEx(DirectDevice, App.path & "\graficos\p" & i & ".png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
    Next i

End Sub

Private Function Effect_FToDW(f As Single) As Long
    '*****************************************************************
    'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_FToDW
    '*****************************************************************
    Dim buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set buf = DirectD3D8.CreateBuffer(4)
    DirectD3D8.BufferSetData buf, 0, 4, 1, f
    DirectD3D8.BufferGetData buf, 0, 4, 1, Effect_FToDW

End Function

Sub Effect_Kill(ByVal EffectIndex As Integer, Optional ByVal KillAll As Boolean = False)
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

    'Find The Next Open Effect Slot
    Do
        EffectIndex = EffectIndex + 1   'Check The Next Slot

        If EffectIndex > NumEffects Then    'Dont Go Over Maximum Amount
            Effect_NextOpenSlot = -1
            Exit Function

        End If

    Loop While Effect(EffectIndex).Used = True    'Check Next If Effect Is In Use

    'Return the next open slot
    Effect_NextOpenSlot = EffectIndex

    'Clear the old information from the effect
    Erase Effect(EffectIndex).Particles()
    Erase Effect(EffectIndex).PartVertex()
    ZeroMemory Effect(EffectIndex), LenB(Effect(EffectIndex))
    Effect(EffectIndex).GoToX = -30000
    Effect(EffectIndex).GoToY = -30000

End Function

Public Sub Effect_Render(ByVal EffectIndex As Integer, _
                         Optional ByVal SetRenderStates As Boolean = True)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Render
    '*****************************************************************

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Set the render state for the size of the particle
    DirectDevice.SetRenderState D3DRS_POINTSIZE, Effect(EffectIndex).FloatSize
    
    'Set the render state to point blitting
    If SetRenderStates Then DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    'Set the texture
    DirectDevice.SetTexture 0, ParticleTexture(Effect(EffectIndex).Gfx)

    'Draw all the particles at once
    DirectDevice.DrawPrimitiveUP D3DPT_POINTLIST, Effect(EffectIndex).ParticleCount, Effect(EffectIndex).PartVertex(0), Len(Effect(EffectIndex).PartVertex(0))

    'Reset the render state back to normal
    If SetRenderStates Then DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Function Effect_Snow_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Snow_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Snow      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(10)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Snow_Reset EffectIndex, LoopC, 1
    Next LoopC

    'Set the initial time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Snow_Reset(ByVal EffectIndex As Integer, _
                              ByVal Index As Long, _
                              Optional ByVal FirstReset As Byte = 0)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Reset
    '*****************************************************************

    If FirstReset = 1 Then

        'The very first reset
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (frmMain.ScaleWidth + 400)), Rnd * (frmMain.ScaleHeight + 50), Rnd * 5, 5 + Rnd * 3, 0, 0

    Else

        'Any reset after first
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (frmMain.ScaleWidth + 400)), -15 - Rnd * 185, Rnd * 5, 5 + Rnd * 3, 0, 0

        If Effect(EffectIndex).Particles(Index).sngX < -20 Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngX > frmMain.ScaleWidth Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngY > frmMain.ScaleHeight Then Effect(EffectIndex).Particles(Index).sngX = Rnd * (frmMain.ScaleWidth + 50)

    End If

    'Set the color
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.8, 0

End Sub

Private Sub Effect_Snow_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check if particle is in use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if to reset the particle
            If Effect(EffectIndex).Particles(LoopC).sngX < -200 Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngX > (frmMain.ScaleWidth + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngY > (frmMain.ScaleHeight + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0

            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Reset the particle
                Effect_Snow_Reset EffectIndex, LoopC

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Sub Effect_UpdateAll()

    If ClientSetup.ParticleEngine = False And bRain = False Then Exit Sub
    '*****************************************************************
    'Updates all of the effects and renders them
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateAll
    '*****************************************************************
    Dim LoopC As Long

    'Make sure we have effects
    If NumEffects = 0 Then Exit Sub

    'Set the render state for the particle effects
    DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    'Update every effect in use
    For LoopC = 1 To NumEffects

        'Make sure the effect is in use
        If Effect(LoopC).Used Then
        
            'Update the effect position if it is binded
            Effect_UpdateBinding LoopC
            
            'Find out which effect is selected, then update it
            
            If ClientSetup.ParticleEngine = True Then
                If Effect(LoopC).EffectNum = EffectNum_Fire Then Effect_Fire_Update LoopC
                If Effect(LoopC).EffectNum = EffectNum_Heal Then Effect_Heal_Update LoopC
                If Effect(LoopC).EffectNum = EffectNum_Bless Then Effect_Bless_Update LoopC
                If Effect(LoopC).EffectNum = EffectNum_Protection Then Effect_Protection_Update LoopC
                If Effect(LoopC).EffectNum = EffectNum_Strengthen Then Effect_Strengthen_Update LoopC
                If Effect(LoopC).EffectNum = EffectNum_EquationTemplate Then Effect_EquationTemplate_Update LoopC
                If Effect(LoopC).EffectNum = EffectNum_Waterfall Then Effect_Waterfall_Update LoopC
                If Effect(LoopC).EffectNum = EffectNum_Summon Then Effect_Summon_Update LoopC
                If Effect(LoopC).EffectNum = EffectNum_Rayo Then Effect_Rayo_Update LoopC
                If Effect(LoopC).EffectNum = EffectNum_Smoke Then Effect_Smoke_Update LoopC

            End If
            
            If Effect(LoopC).EffectNum = EffectNum_Rain Then Effect_Rain_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Snow Then Effect_Snow_Update LoopC
            
            'Render the effect
            Effect_Render LoopC, False

        End If

    Next
    
    'Set the render state back for normal rendering
    DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Function Effect_Rain_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Rain_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Rain      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(10)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Rain_Reset EffectIndex, LoopC, 1
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Rain_Reset(ByVal EffectIndex As Integer, _
                              ByVal Index As Long, _
                              Optional ByVal FirstReset As Byte = 0)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Reset
    '*****************************************************************

    If FirstReset = 1 Then

        'The very first reset
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (frmMain.ScaleWidth + 400)), Rnd * (frmMain.ScaleHeight + 50), Rnd * 5, 25 + Rnd * 12, 0, 0

    Else

        'Any reset after first
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * 1200), -15 - Rnd * 185, Rnd * 5, 25 + Rnd * 12, 0, 0

        If Effect(EffectIndex).Particles(Index).sngX < -20 Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngX > frmMain.ScaleWidth Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngY > frmMain.ScaleHeight Then Effect(EffectIndex).Particles(Index).sngX = Rnd * (frmMain.ScaleWidth + 50)

    End If

    'Set the color
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.4, 0

End Sub

Private Sub Effect_Rain_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check if the particle is in use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update the particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if to reset the particle
            If Effect(EffectIndex).Particles(LoopC).sngX < -200 Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngX > (frmMain.ScaleWidth + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngY > (frmMain.ScaleHeight + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0

            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Reset the particle
                Effect_Rain_Reset EffectIndex, LoopC

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_EquationTemplate_Begin(ByVal X As Single, _
                                       ByVal Y As Single, _
                                       ByVal Gfx As Integer, _
                                       ByVal Particles As Integer, _
                                       Optional ByVal Progression As Single = 1) As Integer
    '*****************************************************************
    'Particle effect template for effects as described on the
    'wiki page: http://www.vbgore.com/Particle_effect_equations
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_EquationTemplate_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_EquationTemplate  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_EquationTemplate_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_EquationTemplate_Reset(ByVal EffectIndex As Integer, _
                                          ByVal Index As Long)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
    '*****************************************************************
    Dim X As Single
    Dim Y As Single
    Dim r As Single
 
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    r = (Index / 20) * Exp(Index / Effect(EffectIndex).Progression Mod 3)
    X = r * Cos(Index)
    Y = r * Sin(Index)
 
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 1, 0.2 + (Rnd * 0.2)

End Sub

Private Sub Effect_EquationTemplate_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_EquationTemplate_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Fire_Begin(ByVal X As Single, _
                           ByVal Y As Single, _
                           ByVal Gfx As Integer, _
                           ByVal Particles As Integer, _
                           Optional ByVal Direction As Integer = 180, _
                           Optional ByVal Progression As Single = 1) As Integer
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Fire_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Fire      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Fire_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Fire_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
    '*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2, 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)

End Sub

Private Sub Effect_Fire_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then

                    'Reset the particle
                    Effect_Fire_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Protection_Begin(ByVal X As Single, _
                                 ByVal Y As Single, _
                                 ByVal Gfx As Integer, _
                                 ByVal Particles As Integer, _
                                 Optional ByVal Size As Byte = 30, _
                                 Optional ByVal Time As Single = 10) As Integer
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Protection_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Protection    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Protection_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Protection_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Reset
    '*****************************************************************
    Dim a As Single
    Dim X As Single
    Dim Y As Single

    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Private Sub Effect_Protection_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Protection_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Waterfall_Begin(ByVal X As Single, _
                                ByVal Y As Single, _
                                ByVal Gfx As Integer, _
                                ByVal Particles As Integer) As Integer
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Waterfall_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Waterfall     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(32)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Waterfall_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Waterfall_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Reset
    '*****************************************************************

    If Int(Rnd * 10) = 1 Then
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 332), Effect(EffectIndex).Y + (Rnd * 130), 0, 8 + (Rnd * 6), 0, 0
    Else
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 332), Effect(EffectIndex).Y + (Rnd * 10), 0, 8 + (Rnd * 6), 0, 0

    End If

    Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.2, 0.2 + (Rnd * 0.4), 0.1
    
End Sub

Private Sub Effect_Waterfall_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(LoopC)
    
            'Check if the particle is in use
            If .Used Then
    
                'Update The Particle
                .UpdateParticle ElapsedTime

                'Check if the particle is ready to die
                If (.sngY > Effect(EffectIndex).Y + 140) Or (.sngA = 0) Then
    
                    'Reset the particle
                    Effect_Waterfall_Reset EffectIndex, LoopC
    
                Else

                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                    Effect(EffectIndex).PartVertex(LoopC).X = .sngX
                    Effect(EffectIndex).PartVertex(LoopC).Y = .sngY
    
                End If
    
            End If
            
        End With

    Next LoopC

End Sub

Function Effect_Summon_Begin(ByVal X As Single, _
                             ByVal Y As Single, _
                             ByVal Gfx As Integer, _
                             ByVal Particles As Integer, _
                             Optional ByVal Progression As Single = 0) As Integer
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Summon_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Summon    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Summon_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Summon_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Reset
    '*****************************************************************
    Dim X As Single
    Dim Y As Single
    Dim r As Single
    
    If Effect(EffectIndex).Progression > 1000 Then
        Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 1.4
    Else
        Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.5

    End If

    r = (Index / 30) * Exp(Index / Effect(EffectIndex).Progression)
    X = r * Cos(Index)
    Y = r * Sin(Index)
    
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0, Rnd, 0, 0.9, 0.2 + (Rnd * 0.2)
 
End Sub

Private Sub Effect_Summon_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 1800 Then

                    'Reset the particle
                    Effect_Summon_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
            
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Bless_Begin(ByVal X As Single, _
                            ByVal Y As Single, _
                            ByVal Gfx As Integer, _
                            ByVal Particles As Integer, _
                            Optional ByVal Size As Byte = 30, _
                            Optional ByVal Time As Single = 10) As Integer
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Bless_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Bless     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Bless_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Bless_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Reset
    '*****************************************************************
    Dim a As Single
    Dim X As Single
    Dim Y As Single

    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Private Sub Effect_Bless_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Bless_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Heal_Begin(ByVal X As Single, _
                           ByVal Y As Single, _
                           ByVal Gfx As Integer, _
                           ByVal Particles As Integer, _
                           Optional ByVal Progression As Single = 1) As Integer
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Heal_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Heal      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Progression = Progression   'Loop the effect
    Effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
    Effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)
    
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Heal_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Heal_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Reset
    '*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0.8, 0.2, 0.2, 0.6 + (Rnd * 0.2), 0.01 + (Rnd * 0.5)
    
End Sub

Private Sub Effect_Heal_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long
    Dim i           As Integer

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then

                    'Reset the particle
                    Effect_Heal_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Strengthen_Begin(ByVal X As Single, _
                                 ByVal Y As Single, _
                                 ByVal Gfx As Integer, _
                                 ByVal Particles As Integer, _
                                 Optional ByVal Size As Byte = 30, _
                                 Optional ByVal Time As Single = 10) As Integer
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Strengthen_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Strengthen    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Strengthen_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Strengthen_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Reset
    '*****************************************************************
    Dim a As Single
    Dim X As Single
    Dim Y As Single

    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 0.2, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Private Sub Effect_Strengthen_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check if particle is in use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update the particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Strengthen_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Public Sub SetSpellCast()

    If frmMain.hlst.List(frmMain.hlst.ListIndex) = "Resucitar" Then
        SpellGrhIndex = 24665
        Exit Sub
    ElseIf frmMain.hlst.List(frmMain.hlst.ListIndex) = "Toxina" Then
        SpellGrhIndex = 24671
        Exit Sub
    ElseIf frmMain.hlst.List(frmMain.hlst.ListIndex) = "Paralizar" Or frmMain.hlst.List(frmMain.hlst.ListIndex) = "Inmovilizar" Then
        SpellGrhIndex = 24669
        Exit Sub
    ElseIf frmMain.hlst.List(frmMain.hlst.ListIndex) = "Antídoto Mágico" Or frmMain.hlst.List(frmMain.hlst.ListIndex) = "Curar Heridas Leves" Or frmMain.hlst.List(frmMain.hlst.ListIndex) = "Curar Heridas Graves" Then
        SpellGrhIndex = 24673
        Exit Sub
    Else
        SpellGrhIndex = 0

    End If

End Sub

Private Sub Effect_UpdateBinding(ByVal EffectIndex As Integer)
 
    '***************************************************
    'Updates the binding of a particle effect to a target, if
    'the effect is bound to a character
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateBinding
    '***************************************************
    Dim TargetI As Integer
    Dim TargetA As Single
 
    'Update position through character binding
    If Effect(EffectIndex).BindToChar > 0 Then
 
        'Store the character index
        TargetI = Effect(EffectIndex).BindToChar
 
        'Check for a valid binding index
        If TargetI > LastChar Then
            Effect(EffectIndex).BindToChar = 0

            If Effect(EffectIndex).KillWhenTargetLost Then
                Effect(EffectIndex).Progression = 0
                Exit Sub

            End If

        ElseIf charlist(TargetI).active = 0 Then
            Effect(EffectIndex).BindToChar = 0

            If Effect(EffectIndex).KillWhenTargetLost Then
                Effect(EffectIndex).Progression = 0
                Exit Sub

            End If

        Else
 
            'Calculate the X and Y positions
            Effect(EffectIndex).GoToX = Engine_TPtoSPX(charlist(Effect(EffectIndex).BindToChar).Pos.X)
            Effect(EffectIndex).GoToY = Engine_TPtoSPY(charlist(Effect(EffectIndex).BindToChar).Pos.Y)
 
        End If
 
    End If
 
    'Move to the new position if needed
    If Effect(EffectIndex).GoToX > -30000 Or Effect(EffectIndex).GoToY > -30000 Then
        If Effect(EffectIndex).GoToX <> Effect(EffectIndex).X Or Effect(EffectIndex).GoToY <> Effect(EffectIndex).Y Then
 
            'Calculate the angle
            TargetA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY) + 180
 
            'Update the position of the effect
            Effect(EffectIndex).X = Effect(EffectIndex).X - Sin(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
            Effect(EffectIndex).Y = Effect(EffectIndex).Y + Cos(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
 
            'Check if the effect is close enough to the target to just stick it at the target
            If Effect(EffectIndex).GoToX > -30000 Then
                If Abs(Effect(EffectIndex).X - Effect(EffectIndex).GoToX) < 6 Then Effect(EffectIndex).X = Effect(EffectIndex).GoToX

            End If

            If Effect(EffectIndex).GoToY > -30000 Then
                If Abs(Effect(EffectIndex).Y - Effect(EffectIndex).GoToY) < 6 Then Effect(EffectIndex).Y = Effect(EffectIndex).GoToY

            End If
 
            'Check if the position of the effect is equal to that of the target
            If Effect(EffectIndex).X = Effect(EffectIndex).GoToX Then
                If Effect(EffectIndex).Y = Effect(EffectIndex).GoToY Then
 
                    'For some effects, if the position is reached, we want to end the effect
                    If Effect(EffectIndex).KillWhenAtTarget Then
                        Effect(EffectIndex).BindToChar = 0
                        Effect(EffectIndex).Progression = 0
                        Effect(EffectIndex).GoToX = Effect(EffectIndex).X
                        Effect(EffectIndex).GoToY = Effect(EffectIndex).Y

                    End If

                    Exit Sub    'The effect is at the right position, don't update
 
                End If

            End If
 
        End If

    End If
 
End Sub

Public Function Engine_GetAngle(ByVal CenterX As Integer, _
                                ByVal CenterY As Integer, _
                                ByVal TargetX As Integer, _
                                ByVal TargetY As Integer) As Single
    '************************************************************
    'Gets the angle between two points in a 2d plane
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetAngle
    '************************************************************
    Dim SideA As Single
    Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then

        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90

            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270

        End If

        'Exit the function
        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then

        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180

        End If

        'Exit the function
        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function

    Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

    Exit Function

End Function

Function Effect_Rayo_Begin(ByVal X As Single, _
                           ByVal Y As Single, _
                           ByVal Gfx As Integer, _
                           ByVal Particles As Integer, _
                           Optional ByVal Progression As Single = 1) As Integer
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rayo_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Rayo_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Rayo      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Progression = Progression   'Loop the effect
    Effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
    Effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)
    
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Rayo_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Rayo_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rayo_Reset
    '*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0, 0.8, 0.8, 0.6 + (Rnd * 0.2), 0.001 + (Rnd * 0.5)
    '      Effect(EffectIndex).Particles(Index).ResetColor (Rnd * 0.8), (Rnd * 0.8), (Rnd * 0.8), 0.6 + (Rnd * 0.2), 0.001 + (Rnd * 0.5)

End Sub

Private Sub Effect_Rayo_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rayo_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long
    Dim i           As Integer

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then

                    'Reset the particle
                    Effect_Rayo_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Sub Engine_Weather_Update()

    If bRain And bLluvia(UserMap) = 1 And CurMapAmbient.Rain = True Then
        If WeatherEffectIndex <= 0 Then
            WeatherEffectIndex = Effect_Rain_Begin(9, 500)
        ElseIf Effect(WeatherEffectIndex).EffectNum <> EffectNum_Rain Then
            Effect_Kill WeatherEffectIndex
            WeatherEffectIndex = Effect_Rain_Begin(9, 500)
        ElseIf Not Effect(WeatherEffectIndex).Used Then
            WeatherEffectIndex = Effect_Rain_Begin(9, 500)

        End If

    End If

    If CurMapAmbient.Snow = True Then
        If WeatherEffectIndex <= 0 Then
            WeatherEffectIndex = Effect_Snow_Begin(14, 200)
        ElseIf Effect(WeatherEffectIndex).EffectNum <> EffectNum_Snow Then
            Effect_Kill WeatherEffectIndex
            WeatherEffectIndex = Effect_Snow_Begin(14, 200)
        ElseIf Not Effect(WeatherEffectIndex).Used Then
            WeatherEffectIndex = Effect_Snow_Begin(14, 200)

        End If

    End If
            
    If CurMapAmbient.Fog <> -1 Then
        Engine_Weather_UpdateFog

    End If
    
    If OnRampageImgGrh <> 0 Then
        DDrawTransGrhIndextoSurface OnRampageImgGrh, 0, 0, 0, Normal_RGBList(), 0, True

    End If
    
End Sub

Sub Engine_Weather_UpdateFog()
    '*****************************************************************
    'Update the fog effects
    '*****************************************************************

    Dim i           As Long
    Dim X           As Long
    Dim Y           As Long
    Dim CC(3)       As Long
    Dim ElapsedTime As Single
    ElapsedTime = Engine_ElapsedTime

    If WeatherFogCount = 0 Then WeatherFogCount = 13

    WeatherFogX1 = WeatherFogX1 + (ElapsedTime * (0.018 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY1 = WeatherFogY1 + (ElapsedTime * (0.013 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
    
    Do While WeatherFogX1 < -512
        WeatherFogX1 = WeatherFogX1 + 512
    Loop

    Do While WeatherFogY1 < -512
        WeatherFogY1 = WeatherFogY1 + 512
    Loop

    Do While WeatherFogX1 > 0
        WeatherFogX1 = WeatherFogX1 - 512
    Loop

    Do While WeatherFogY1 > 0
        WeatherFogY1 = WeatherFogY1 - 512
    Loop
    
    WeatherFogX2 = WeatherFogX2 - (ElapsedTime * (0.037 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY2 = WeatherFogY2 - (ElapsedTime * (0.021 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)

    Do While WeatherFogX2 < -512
        WeatherFogX2 = WeatherFogX2 + 512
    Loop

    Do While WeatherFogY2 < -512
        WeatherFogY2 = WeatherFogY2 + 512
    Loop

    Do While WeatherFogX2 > 0
        WeatherFogX2 = WeatherFogX2 - 512
    Loop

    Do While WeatherFogY2 > 0
        WeatherFogY2 = WeatherFogY2 - 512
    Loop
    
    'Render fog 2
    X = 2
    Y = -1

    CC(1) = D3DColorARGB(CurMapAmbient.Fog, 255, 255, 255)
    CC(2) = D3DColorARGB(CurMapAmbient.Fog, 255, 255, 255)
    CC(3) = D3DColorARGB(CurMapAmbient.Fog, 255, 255, 255)
    CC(0) = D3DColorARGB(CurMapAmbient.Fog, 255, 255, 255)

    For i = 1 To WeatherFogCount
        DDrawTransGrhIndextoSurface 27300, (X * 512) + WeatherFogX2, (Y * 512) + WeatherFogY2, 0, CC(), 0, False
        X = X + 1

        If X > (1 + (ScreenWidth \ 512)) Then
            X = 0
            Y = Y + 1

        End If

    Next i
            
    'Render fog 1
    X = 0
    Y = 0
    CC(1) = D3DColorARGB(CurMapAmbient.Fog / 2, 255, 255, 255)
    CC(2) = D3DColorARGB(CurMapAmbient.Fog / 2, 255, 255, 255)
    CC(3) = D3DColorARGB(CurMapAmbient.Fog / 2, 255, 255, 255)
    CC(0) = D3DColorARGB(CurMapAmbient.Fog / 2, 255, 255, 255)

    For i = 1 To WeatherFogCount
        DDrawTransGrhIndextoSurface 27301, (X * 512) + WeatherFogX1, (Y * 512) + WeatherFogY1, 0, CC(), 0, False
        X = X + 1

        If X > (2 + (ScreenWidth \ 512)) Then
            X = 0
            Y = Y + 1

        End If

    Next i

End Sub

Function Effect_Smoke_Begin(ByVal X As Single, _
                            ByVal Y As Single, _
                            ByVal Gfx As Integer, _
                            ByVal Particles As Integer, _
                            Optional ByVal Direction As Integer = 180, _
                            Optional ByVal Progression As Single = 1) As Integer
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Smoke_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Smoke_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Smoke      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(30)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Smoke_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Smoke_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Smoke_Reset
    '*****************************************************************

    'Reset the particle
    'Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
    'Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2, 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 5, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 2, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0.2, 0.2, 0.2, 0.2 + (Rnd * 0.2), 0.03 + (Rnd * 0.01)

End Sub

Private Sub Effect_Smoke_Update(ByVal EffectIndex As Integer)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Smoke_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then

                    'Reset the particle
                    Effect_Smoke_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Public Sub Engine_CreateEffect(ByVal CharIndex As Integer, _
                               ByVal Effect As Byte, _
                               ByVal X As Single, _
                               ByVal Y As Single, _
                               ByVal Loops As Integer)

    If Not ValidNumber(Effect, eNumber_Types.ent_Byte) Then
        'Call Log_Engine("Error in Engine_CreateEffect, The effect " & Effect & " is not valid.")
        Exit Sub

    End If
        
    Select Case Effect

        Case EffectNum_Fire

            'Case EffectNum_Snow
        Case EffectNum_Heal

        Case EffectNum_Bless

        Case EffectNum_Protection
        
        Case EffectNum_Strengthen
            charlist(CharIndex).ParticleIndex = Effect_Strengthen_Begin(X, Y, 2, 100, 30, CSng(Loops))
            
            'Case EffectNum_Rain
        Case EffectNum_EquationTemplate
        
            'Case EffectNum_Waterfall
        
        Case EffectNum_Summon
            charlist(CharIndex).ParticleIndex = Effect_Summon_Begin(X, Y, 1, 500, 0.1)
            
        Case EffectNum_Rayo
            charlist(CharIndex).ParticleIndex = Effect_Rayo_Begin(X, Y, 2, 100, -1)

    End Select

End Sub

Public Function Speed_Particle() As Single
    '***************************************************
    'Author: Standelf
    'Last Modify Date: 06/01/2011
    '***************************************************

    Dim ElapsedTime As Single
    ElapsedTime = Engine_ElapsedTime
    Speed_Particle = ElapsedTime * 0.2

End Function

