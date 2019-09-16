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
    G As Integer
    B As Integer
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
                    If CurMapAmbient.UseDayAmbient = False Then Call Engine_D3DColor_To_RGB_List(MapData(Xx, Yy).Engine_Light(), CurMapAmbient.OwnAmbientLight)
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
        .OwnAmbientLight.A = 255
        .OwnAmbientLight.r = 0
        .OwnAmbientLight.G = 0
        .OwnAmbientLight.B = 0
        
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
                    Call Create_Light_To_Map(Xx, Yy, .MapBlocks(Xx, Yy).Light.range, .MapBlocks(Xx, Yy).Light.r, .MapBlocks(Xx, Yy).Light.G, .MapBlocks(Xx, Yy).Light.B)

                End If

            Next Yy
        Next Xx
            
        Call LightRenderAll
            
        If .UseDayAmbient = True Then
            frmAmbientEditor.Option1(0).Value = True
        Else
            frmAmbientEditor.Option1(1).Value = True
            frmAmbientEditor.Text1(0).Text = .OwnAmbientLight.r
            frmAmbientEditor.Text1(1).Text = .OwnAmbientLight.G
            frmAmbientEditor.Text1(2).Text = .OwnAmbientLight.B

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
