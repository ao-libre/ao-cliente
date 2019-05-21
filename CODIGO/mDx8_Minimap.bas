Attribute VB_Name = "mDx8_Minimap"
'/////////////////////////////Motor Grafico en DirectX 8///////////////////////////////
'////////////////////////Extraccion de varios motores por ShaFTeR//////////////////////
'***************************************GoDKeR*****************************************
Option Explicit

Public AlphaMiniMap As Byte

'Dim MMC_Blocked     As Long
'Dim MMC_Mountain    As Long
'Dim MMC_Exit        As Long
'Dim MMC_Sign        As Long

Private MMC_Char        As Long

'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A

    Width As Long
    Height As Long

End Type

Private Type POINTAPI

    X As Long
    Y As Long

End Type

Public Type tMinMap

    color(0 To 3) As Long
    X As Integer
    Y As Integer
    Texture As Direct3DTexture8 'Holds the texture of the text
    TextureSize As POINTAPI     'Size of the texture

End Type

Public Minimap As tMinMap ' _Default2 As CustomFont

Public Sub MiniMap_Init()
    'MMC_Blocked = D3DColorARGB(75, 255, 255, 255)   'Blocked tiles
    'MMC_Exit = D3DColorARGB(150, 255, 0, 0)         'Exit tiles (warps)
    'MMC_Sign = D3DColorARGB(125, 255, 255, 0)       'Tiles with a sign
    MMC_Char = D3DColorARGB(150, 255, 0, 0)
    'MMC_Mountain = D3DColorARGB(150, 206, 130, 72)

End Sub

Public Sub MiniMap_Render(ByVal X As Long, ByVal Y As Long)
    
    If Not FileExist(App.path & "\Graficos\MiniMapa\" & UserMap & ".bmp", vbArchive) Then Exit Sub
    
    Dim VertexArray(0 To 3) As TLVERTEX
    Dim SrcWidth            As Integer
    Dim Width               As Integer
    Dim SrcHeight           As Integer
    Dim Height              As Integer
    Dim SrcBitmapWidth      As Long
    Dim SrcBitmapHeight     As Long
    Dim SRDesc              As D3DSURFACE_DESC
        
    With Minimap
    
        If Not .Texture Is Nothing Then
            .Texture.GetLevelDesc 0, SRDesc
                        
            SrcWidth = 100 'd3dtextures.texwidth
            Width = 100 'd3dtextures.texwidth
            Height = 100 'd3dtextures.texheight
            SrcHeight = 100 'd3dtextures.texheight
                        
            SrcBitmapWidth = SRDesc.Width
            SrcBitmapHeight = SRDesc.Height
                        
            'Set the RHWs (must always be 1)
            VertexArray(0).rhw = 1
            VertexArray(1).rhw = 1
            VertexArray(2).rhw = 1
            VertexArray(3).rhw = 1
                        
            'Find the left side of the rectangle
            VertexArray(0).X = X
            VertexArray(0).tu = (.TextureSize.X / SrcBitmapWidth)
                        
            'Find the top side of the rectangle
            VertexArray(0).Y = Y
            VertexArray(0).tv = (.TextureSize.Y / SrcBitmapHeight)
                       
            'Find the right side of the rectangle
            VertexArray(1).X = X + Width
            VertexArray(1).tu = (.TextureSize.X + SrcWidth) / SrcBitmapWidth
                       
            'These values will only equal each other when not a shadow
            VertexArray(2).X = VertexArray(0).X
            VertexArray(3).X = VertexArray(1).X
                        
            'Find the bottom of the rectangle
            VertexArray(2).Y = Y + Height
            VertexArray(2).tv = (.TextureSize.Y + SrcHeight) / SrcBitmapHeight
                        
            'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
            VertexArray(1).Y = VertexArray(0).Y
            VertexArray(1).tv = VertexArray(0).tv
            VertexArray(2).tu = VertexArray(0).tu
            VertexArray(3).Y = VertexArray(2).Y
            VertexArray(3).tu = VertexArray(1).tu
            VertexArray(3).tv = VertexArray(2).tv
            VertexArray(0).color = .color(0)
            VertexArray(1).color = .color(1)
            VertexArray(2).color = .color(2)
            VertexArray(3).color = .color(3)
                        
            'Set the texture
            DirectDevice.SetTexture 0, .Texture
            DirectDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(AlphaMiniMap, 0, 0, 0)
                        
            'faster
            DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), LenB(VertexArray(0))

            Call MiniMap_ColorSet

        End If

    End With

End Sub

Public Sub MiniMap_UserPos()
    
    If AlphaMiniMap > 50 Then
        Call mDx8_Engine.Engine_Draw_Box(UserPos.X, UserPos.Y, 4, 4, MMC_Char)
    End If
    
End Sub


Private Sub MiniMap_ColorSet()

    On Error GoTo Err

    Dim Colorsitus As D3DCOLORVALUE

    If frmMain.MouseX > Minimap.X And frmMain.MouseY > Minimap.Y And frmMain.MouseX < Minimap.X + 100 And frmMain.MouseY < Minimap.Y + 100 Then

        If AlphaMiniMap <> 0 Then
            AlphaMiniMap = AlphaMiniMap - timerTicksPerFrame * 25

            If AlphaMiniMap < 10 Then AlphaMiniMap = 0

        End If

    Else

        If AlphaMiniMap <> 205 Then
            AlphaMiniMap = AlphaMiniMap + timerTicksPerFrame * 25

            If AlphaMiniMap > 195 Then AlphaMiniMap = 205

        End If

    End If

    With Colorsitus
    
        .r = 255
        .g = 255
        .b = 255
        .a = AlphaMiniMap
    
    End With
    
    mDx8_Engine.Engine_D3DColor_To_RGB_List Minimap.color(), Colorsitus

    Exit Sub

Err:

    With Colorsitus
        .a = 205
        .r = 255
        .g = 255
        .b = 255

    End With
    
    mDx8_Engine.Engine_D3DColor_To_RGB_List Minimap.color(), Colorsitus

End Sub

Public Sub MiniMap_ChangeTex(UserMap As Integer)

    Dim mapInfo As D3DXIMAGE_INFO_A

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
    With Minimap
    
        'Set the texture
        Set .Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, App.path & "\Graficos\MiniMapa\" & UserMap & ".bmp", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
        
        'Store the size of the texture
        With .TextureSize
            .X = mapInfo.Width
            .Y = mapInfo.Height
        End With
    
    End With

    Exit Sub

End Sub

