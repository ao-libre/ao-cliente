Attribute VB_Name = "mDx8_Clima"
Option Explicit

'***************************************************
'Author: Ezequiel Juarez (Standelf)
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

Public Estados(1 To 5) As D3DCOLORVALUE
Public Estado_Actual As D3DCOLORVALUE
Public Estado_Actual_Date As Byte

Public Sub Init_MeteoEngine()
'***************************************************
'Author: Standelf
'Last Modification: 15/05/10
'Initializate
'***************************************************
    With Estados(e_estados.AMANECER)
        .A = 255
        .r = 255
        .g = 200
        .B = 200
    End With
    
    With Estados(e_estados.MEDIODIA)
        .A = 255
        .r = 240
        .g = 250
        .B = 210
    End With
    
    With Estados(e_estados.DIA)
        .A = 255
        .r = 255
        .g = 255
        .B = 255
    End With
    
    With Estados(e_estados.ATARDECER)
        .A = 255
        .r = 150
        .g = 120
        .B = 120
    End With
  
    With Estados(e_estados.NOCHE)
        .A = 255
        .r = 100
        .g = 100
        .B = 100
    End With
    
    Estado_Actual_Date = 3
    
End Sub

Public Sub Set_AmbientColor()
    Estado_Actual.A = 255
    Estado_Actual.B = CurMapAmbient.OwnAmbientLight.B
    Estado_Actual.g = CurMapAmbient.OwnAmbientLight.g
    Estado_Actual.r = CurMapAmbient.OwnAmbientLight.r
End Sub

Public Sub Actualizar_Estado(ByVal Estado As Byte)
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
End Sub

Public Sub Start_Rampage()
'***************************************************
'Author: Standelf
'Last Modification: 27/05/2010
'Init Rampage
'***************************************************
    Dim X As Byte, Y As Byte, TempColor As D3DCOLORVALUE
    TempColor.A = 255: TempColor.B = 255: TempColor.r = 255: TempColor.g = 255
    
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
    Dim X As Byte, Y As Byte

    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
        Next Y
    Next X

    Call LightRenderAll

End Sub

