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

End Enum

Public Estados(1 To 5)    As D3DCOLORVALUE
Public Estado_Actual      As D3DCOLORVALUE
Public Estado_Actual_Date As Byte

Public Sub Init_MeteoEngine()
    
    On Error GoTo Init_MeteoEngine_Err
    

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
    
    Estado_Actual_Date = 3
    
    
    Exit Sub

Init_MeteoEngine_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Clima" & "->" & "Init_MeteoEngine"
    End If
Resume Next
    
End Sub

Public Sub Set_AmbientColor()
    
    On Error GoTo Set_AmbientColor_Err
    
    Estado_Actual.a = 255
    Estado_Actual.b = CurMapAmbient.OwnAmbientLight.b
    Estado_Actual.g = CurMapAmbient.OwnAmbientLight.g
    Estado_Actual.r = CurMapAmbient.OwnAmbientLight.r

    
    Exit Sub

Set_AmbientColor_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Clima" & "->" & "Set_AmbientColor"
    End If
Resume Next
    
End Sub

Public Sub Actualizar_Estado(ByVal Estado As Byte)
    
    On Error GoTo Actualizar_Estado_Err
    

    '***************************************************
    'Author: Standelf
    'Last Modification: 15/05/10
    'Update State and RenderLights
    '***************************************************
    If Estado < 0 Or Estado > 5 Then Exit Sub
    If CurMapAmbient.UseDayAmbient = False Then Exit Sub
    
    If Estado = 0 Then Estado = e_estados.DIA
        
    Estado_Actual = Estados(Estado)
    Estado_Actual_Date = Estado
        
    Dim X As Byte, Y As Byte

    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
        Next Y
    Next X
        
    Call LightRenderAll

    
    Exit Sub

Actualizar_Estado_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Clima" & "->" & "Actualizar_Estado"
    End If
Resume Next
    
End Sub

Public Sub Start_Rampage()
    '***************************************************
    'Author: Standelf
    'Last Modification: 27/05/2010
    'Init Rampage
    '***************************************************
    
    On Error GoTo Start_Rampage_Err
    
    Dim X As Byte, Y As Byte, TempColor As D3DCOLORVALUE
    TempColor.a = 255: TempColor.b = 255: TempColor.r = 255: TempColor.g = 255
    
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), TempColor)
        Next Y
    Next X

    
    Exit Sub

Start_Rampage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Clima" & "->" & "Start_Rampage"
    End If
Resume Next
    
End Sub

Public Sub End_Rampage()
    '***************************************************
    'Author: Standelf
    'Last Modification: 27/05/2010
    'End Rampage
    '***************************************************
    
    On Error GoTo End_Rampage_Err
    
    OnRampageImgGrh = 0
    OnRampageImg = 0
    
    Dim X As Byte, Y As Byte

    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
        Next Y
    Next X

    Call LightRenderAll

    
    Exit Sub

End_Rampage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Clima" & "->" & "End_Rampage"
    End If
Resume Next
    
End Sub

