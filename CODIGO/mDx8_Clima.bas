Attribute VB_Name = "mDx8_Clima"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 15/05/10
'Blisse-AO | Set the Roof Color and Render _
    the Lights.
'***************************************************

Enum e_estados
    AMANECER = 1
    MEDIODIA = 2
    DIA = 3
    ATARDECER = 4
    NOCHE = 5
    LLUVIA = 6
End Enum

Enum e_render
    BEGINING = 1
    RUNNING = 2
    STOPPING = 3
    STOPED = 4
End Enum

Public Estados(1 To 5) As D3DCOLORVALUE
Public Estado_Actual As D3DCOLORVALUE
Public Estado_Actual_Date As Byte
Public current_State As Byte
Public blendSteps(1 To 3) As Double
Public blendAmount(1 To 3) As Double
Public greaterDif, greaterStep As Integer

Public Sub Init_MeteoEngine()
'***************************************************
'Author: Standelf
'Last Modification: 15/05/10
'Initializate
'***************************************************
    With Estados(e_estados.AMANECER)
        .a = 255
        .r = 255
        .g = 200
        .b = 200
    End With
    
    With Estados(e_estados.MEDIODIA)
        .a = 255
        .r = 240
        .g = 250
        .b = 210
    End With
    
    With Estados(e_estados.DIA)
        .a = 255
        .r = 255
        .g = 255
        .b = 255
    End With
    
    With Estados(e_estados.ATARDECER)
        .a = 255
        .r = 150
        .g = 120
        .b = 120
    End With
  
    With Estados(e_estados.NOCHE)
        .a = 255
        .r = 100
        .g = 100
        .b = 100
    End With
    
    With Estados(e_estados.LLUVIA)
        .a = 255
        .r = 230
        .g = 230
        .b = 230
    End With
    
    Estado_Actual_Date = 3
    current_State = e_render.STOPED
    
End Sub

Public Sub Set_AmbientColor()
    Estado_Actual.a = 255
    Estado_Actual.b = CurMapAmbient.OwnAmbientLight.b
    Estado_Actual.g = CurMapAmbient.OwnAmbientLight.g
    Estado_Actual.r = CurMapAmbient.OwnAmbientLight.r
End Sub

Public Sub Actualizar_Estado(ByVal Estado As Byte)
'***************************************************
'Author: Standelf
'Last Modification: 15/05/10
'Update State and RenderLights
'***************************************************
    If Estado < 0 Or Estado > 6 Then Exit Sub
    If current_State = e_render.STOPED Then Exit Sub
    If CurMapAmbient.UseDayAmbient = False Then Exit Sub
    
        If Estado = 0 Then Estado = e_estados.DIA
    
    If Estado_Actual_Date <> Estado Then
        Call calculateBlendSteps(Estado_Actual_Date, Estado)
        current_State = e_render.BEGINING
        
        Estado_Actual = Estados(Estado)
        Estado_Actual_Date = Estado
    End If
    
    If current_State = e_render.BEGINING Then
        'hacer animacion
    ElseIf current_State = e_render.RUNNING Then
        Dim X As Byte, Y As Byte
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
            Next Y
        Next X
        
        Call LightRenderAll
        current_State = e_render.STOPED
    End If
    
    
End Sub

Public Sub Start_Rampage()
'***************************************************
'Author: Standelf
'Last Modification: 27/05/2010
'Init Rampage
'***************************************************
    Dim X As Byte, Y As Byte, TempColor As D3DCOLORVALUE
    TempColor.a = 255: TempColor.b = 255: TempColor.r = 255: TempColor.g = 255
    
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), TempColor)
            Next Y
        Next X
End Sub

Public Sub End_Rampage()
'***************************************************
'Author: Standelf
'Last Modification: 27/05/2010
'End Rampage
'***************************************************
    OnRampageImgGrh = 0
    OnRampageImg = 0
    
    Dim X As Byte, Y As Byte
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
            Next Y
        Next X
        Call LightRenderAll
End Sub

Public Sub calculateBlendSteps(ByVal origState As Integer, ByVal destState As Integer)
'***************************************************
'Author: Cucsifae
'Last Modification:14/09/2018
'
'***************************************************
    'erase last blendsteps
    blendSteps(1) = 0
    blendSteps(2) = 0
    blendSteps(3) = 0
    greaterDif = 0
    greaterStep = 0
    
    With Estados(destState)
        blendSteps(1) = .r - Estados(origState).r
        blendSteps(2) = .g - Estados(origState).g
        blendSteps(3) = .b - Estados(origState).b
    End With

    If blendSteps(1) >= blendSteps(2) Then
        greaterDif = difSteps(1)
        greaterStep = 1
        If blendSteps(1) < blendSteps(3) Then
            greaterDif = blendSteps(3)
            greaterStep = 3
        End If
    ElseIf blendSteps(2) >= blendSteps(3) Then
        greaterDif = blendSteps(2)
        greaterStep = 2
    Else
        greaterDif = blendSteps(3)
        greaterStep = 3
    End If
    
    blendAmount(greaterStep) = 1
    
    If greaterStep = 1 Then
        blendAmount(2) = blendSteps(1) / blendSteps(2)
        blendAmount(3) = blendSteps(1) / blendSteps(3)
    ElseIf greaterStep = 2 Then
        blendAmount(1) = blendSteps(2) / blendSteps(1)
        blendAmount(3) = blendSteps(2) / blendSteps(3)
    Else
        blendAmount(1) = blendSteps(3) / blendSteps(1)
        blendAmount(2) = blendSteps(3) / blendSteps(2)
    End If
    
End Sub

Public Sub BlendStates()
'***************************************************
'Author: Cucsifae
'Last Modification:14/09/2018
'
'***************************************************
    'erase last blendsteps
    blendSteps(1) = 0
    blendSteps(2) = 0
    blendSteps(3) = 0
    greaterDif = 0
    greaterStep = 0
    
   
    
    
End Sub
