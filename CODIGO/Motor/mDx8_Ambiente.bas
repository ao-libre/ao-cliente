Attribute VB_Name = "mDx8_Ambiente"
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: ??/??/10
'Blisse-AO | Sistema de Ambientes
'***************************************************

Option Explicit

Type A_Light
    range As Byte
    r As Integer
    g As Integer
    b As Integer
End Type

Type MapAmbientBlock
    Light As A_Light
    Particle As Byte
End Type

Type MapAmbient
    MapBlocks() As MapAmbientBlock
    UseDayAmbient As Boolean
    OwnAmbientLight As D3DCOLORVALUE
    Fog As Integer
    Snow As Boolean
    Rain As Boolean
End Type

Public CurMapAmbient As MapAmbient
    
Public Sub Apply_OwnAmbient()

    If CurMapAmbient.UseDayAmbient = False Then
        Estado_Actual = CurMapAmbient.OwnAmbientLight
    Else
        Call Actualizar_Estado(Estado_Actual_Date)
    End If
    
    Dim Xx As Integer, Yy As Integer

    For Xx = XMinMapSize To XMaxMapSize
        For Yy = YMinMapSize To YMaxMapSize
            
            If CurMapAmbient.UseDayAmbient = False Then
                Call Engine_D3DColor_To_RGB_List(MapData(Xx, Yy).Engine_Light(), CurMapAmbient.OwnAmbientLight)
            End If
            
        Next Yy
    Next Xx
            
    Call LightRenderAll

End Sub

Public Sub Init_Ambient(ByVal Map As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 15/10/10
'***************************************************
    With CurMapAmbient
        .Fog = -1
        .UseDayAmbient = True
        .OwnAmbientLight.a = 255
        .OwnAmbientLight.r = 0
        .OwnAmbientLight.g = 0
        .OwnAmbientLight.b = 0
        
        .Rain = True
        .Snow = False
        
        ReDim .MapBlocks(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapAmbientBlock
        
        If FileExist(App.path & "\Ambiente\" & Map & ".amb", vbNormal) Then
            Dim N As Integer
            N = FreeFile
                Open App.path & "\Ambiente\" & Map & ".amb" For Binary As #N
                    Get #N, , CurMapAmbient
                Close #N
        End If
        
        If .UseDayAmbient = False Then
            Estado_Actual = .OwnAmbientLight
        Else
            Call Actualizar_Estado(Estado_Actual_Date)
        End If
                    
        Dim Xx As Integer, Yy As Integer
        
            For Xx = XMinMapSize To XMaxMapSize
                For Yy = YMinMapSize To YMaxMapSize
                    If .UseDayAmbient = False Then Call Engine_D3DColor_To_RGB_List(MapData(Xx, Yy).Engine_Light(), .OwnAmbientLight)
                    
                    If .MapBlocks(Xx, Yy).Light.range <> 0 Then
                        Create_Light_To_Map Xx, Yy, .MapBlocks(Xx, Yy).Light.range, .MapBlocks(Xx, Yy).Light.r, .MapBlocks(Xx, Yy).Light.g, .MapBlocks(Xx, Yy).Light.b
                    End If
                Next Yy
            Next Xx
            
        Call LightRenderAll
            
            If .UseDayAmbient = True Then
                frmAmbientEditor.Option1(0).Value = True
            Else
                frmAmbientEditor.Option1(1).Value = True
                frmAmbientEditor.Text1(0).Text = .OwnAmbientLight.r
                frmAmbientEditor.Text1(1).Text = .OwnAmbientLight.g
                frmAmbientEditor.Text1(2).Text = .OwnAmbientLight.b
            End If
                                        
            If .Fog <> -1 Then
                frmAmbientEditor.Check1.Value = Checked
                frmAmbientEditor.HScroll1.Value = .Fog
            Else
                frmAmbientEditor.Check1.Value = Unchecked
            End If
            
            If .Rain = True Then frmAmbientEditor.Check3.Value = Checked
            If .Snow = True Then frmAmbientEditor.Check2.Value = Checked
            
            
    End With
End Sub

Public Sub Save_Ambient(ByVal Map As Integer)
    '***************************************************
    'Author: Standelf
    'Last Modification: 15/10/10
    '***************************************************
    Debug.Print CurMapAmbient.UseDayAmbient
    
    Dim File As Integer: File = FreeFile

    Open App.path & "\Ambiente\" & Map & ".amb" For Binary Access Write As File
        Put File, , CurMapAmbient
    Close #File

End Sub

Public Sub DiaNoche()

    '*****************************************************************
    'Author: Pablo D. (DISCORD: Abusivo#1215)
    '*****************************************************************
    Static lastmovement As Long
    
    'Se ejecuta cada 30min
    If GetTickCount - lastmovement > 30000 Then
    
        lastmovement = GetTickCount + 30000

        Dim Hora As Integer: Hora = Hour(Now)

        With CurMapAmbient
        
            .UseDayAmbient = False
            
            With .OwnAmbientLight
                
                Select Case Hora
                    
                    Case 0
                        Call ShowConsoleMsg("Es media noche.")
                        
                    Case Is >= 1
                        .a = 181.6
                        .r = 181.6
                        .g = 181.6
                        .b = 181.6
                        
                    Case Is >= 6
                        .a = 100 + (Hora * 12.9)
                        .r = 100 + (Hora * 12.9)
                        .g = 100 + (Hora * 12.9)
                        .b = 100 + (Hora * 12.9)
                        
                        Call ShowConsoleMsg("Ha comenzado a amanercer.")
                    
                    Case Is >= 12
                        .a = 255
                        .r = 255
                        .g = 255
                        .b = 255
                        
                        Call ShowConsoleMsg("Es medio dia.")
                        
                    Case Is >= 18
                        .a = 255 - (Hora * 3.4)
                        .r = 255 - (Hora * 3.4)
                        .g = 255 - (Hora * 3.4)
                        .b = 255 - (Hora * 3.4)
                        
                        Call ShowConsoleMsg("Ha comenzado a anochecer.")
                        
                End Select
                
            End With
        
        End With

        'setear luz
        Call Apply_OwnAmbient
       
    End If
    
End Sub
