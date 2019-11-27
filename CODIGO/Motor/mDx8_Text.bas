Attribute VB_Name = "mDx8_Text"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, source As Any, ByVal Length As Long)
    
Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type

Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type

Private cfonts(1 To 2) As CustomFont ' _Default2 As CustomFont

Public Function ColorToDX8(ByVal long_color As Long) As Long
    Dim temp_color As String
    Dim Red As Integer, Blue As Integer, Green As Integer
    
    temp_color = Hex$(long_color)
    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String$(6 - Len(temp_color), "0") + temp_color
    End If
    
    Red = CLng("&H" + mid$(temp_color, 1, 2))
    Green = CLng("&H" + mid$(temp_color, 3, 2))
    Blue = CLng("&H" + mid$(temp_color, 5, 2))
    
    ColorToDX8 = D3DColorXRGB(Red, Green, Blue)

End Function

Public Sub Text_Render_Special(ByVal intX As Integer, ByVal intY As Integer, ByRef strText As String, ByVal lngColor As Long, Optional bolCentred As Boolean = False, Optional Font As Integer = 1)  ' GSZAO
'*****************************************************************
'Text_Render_Special by ^[GS]^
'*****************************************************************
    Dim i As Long
    If LenB(strText) <> 0 Then
        lngColor = ColorToDX8(lngColor)
        Call Engine_Render_Text(cfonts(1), strText, intX, intY, lngColor, bolCentred) 
    End If
    
End Sub ' GSZAO

Private Function Es_Emoticon(ByVal ascii As Byte) As Boolean ' GSZAO
'*****************************************************************
'Emoticones by ^[GS]^
'*****************************************************************
    Es_Emoticon = False
    If (ascii = 129 Or ascii = 137 Or ascii = 141 Or ascii = 143 Or ascii = 144 Or ascii = 157 Or ascii = 160) Then
        Es_Emoticon = True
    End If
End Function ' GSZAO

Private Sub Engine_Render_Text(ByRef UseFont As CustomFont, 
                                ByVal Text As String, 
                                ByVal X As Long, 
                                ByVal Y As Long, 
                                ByVal Color As Long, Optional 
                                ByVal Center As Boolean = False, Optional 
                                ByVal Alpha As Byte = 255, Optional 
                                ByVal ParseEmoticons As Boolean = True)
                                
'*****************************************************************
'Render text with a custom font
'*****************************************************************
    Dim TempVA(0 To 3) As TLVERTEX
    Dim tempstr() As String
    Dim Count As Integer
    Dim ascii() As Byte
    Dim i As Long
    Dim J As Long
    Dim yOffset As Single

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    
     'WyroX: Agregado para evitar dibujar emojis en los nombres de los personajes
    If ParseEmoticons Then
        'Analizar mensaje, palabra por palabra... GSZAO
        Dim NewText As String
        
        tempstr = Split(Text, Chr$(32))
        NewText = Text
        Text = vbNullString

        Dim Upper_tempstr As Long
        Upper_tempstr = UBound(tempstr)

        For i = 0 To Upper_tempstr
            If tempstr(i) = ":)" Or tempstr(i) = "=)" Then
                tempstr(i) = Chr$(129)
            ElseIf tempstr(i) = ":@" Or tempstr(i) = "=@" Then
                tempstr(i) = Chr$(137)
            ElseIf tempstr(i) = ":(" Or tempstr(i) = "=(" Then
                tempstr(i) = Chr$(141)
            ElseIf tempstr(i) = "^^" Or tempstr(i) = "^_^" Then
                tempstr(i) = Chr$(143)
            ElseIf tempstr(i) = ":D" Or tempstr(i) = "=D" Then
                tempstr(i) = Chr$(144)
            ElseIf tempstr(i) = "xD" Or tempstr(i) = "XD" Then
                tempstr(i) = Chr$(157)
            ElseIf tempstr(i) = ":S" Or tempstr(i) = "=S" Then
                tempstr(i) = Chr$(160)
            End If
            Text = Text & Chr$(32) & tempstr(i)
        Next
        ' Made by ^[GS]^ for GSZAO
    End If
    
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)

    'Set the texture
    ' Call Batch.SetTexture(UseFont.Texture)
    DirectDevice.SetTexture 0, UseFont.Texture 

    If Center Then
        X = X - Engine_GetTextWidth(cfonts(Font), Text) * 0.5
    End If

    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To Upper_tempstr
        ' Dim Len_tempstr As Long
        ' Len_tempstr = UBound(tempstr)

        If Len(tempstr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            Count = 0
        
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
            
            'Loop through the characters
            For J = 1 To Len(tempstr(i))

                Call CopyMemory(TempVA, UseFont.HeaderInfo.CharVA(ascii(J - 1)), 24) 'this number represents the size of "CharVA" struct
                
                TempVA.X = X + Count
                TempVA.Y = Y + yOffset
                
                'Set the colors
                If Es_Emoticon(ascii(J - 1)) Then ' GSZAO los colores no afectan a los emoticones!
                    
                    If (ascii(J - 1) <> 157) Then
                        Count = Count + 5   ' Los emoticones tienen tamano propio (despues hay que cargarlos "correctamente" para evitar hacer esto)
                    End If
                    
                End If
            
                ' Call Batch.Draw(TempVA.X, TempVA.Y, TempVA.W, TempVA.H, Color, TempVA.Tx1, TempVA.Ty1, TempVA.Tx2, TempVA.Ty2)
                DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0)) 
                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(ascii(J - 1))
                
            Next J
            
        End If
    Next i

End Sub

Public Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef Color As D3DCOLORVALUE)
Dim dest(3) As Byte
CopyMemory dest(0), ARGB, 4
Color.A = dest(3)
Color.r = dest(2)
Color.g = dest(1)
Color.B = dest(0)
End Function

Public Function ARGB(ByVal r As Long, ByVal g As Long, ByVal B As Long, ByVal A As Long) As Long
        
    Dim c As Long
        
    If A > 127 Then
        A = A - 128
        c = A * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or B
    Else
        c = A * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or B
    End If
    
    ARGB = c

End Function

Private Function Engine_GetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth
'***************************************************
Dim i As Integer
Dim Len_text As Long

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    Len_text = Len(Text)
    
    'Loop through the text
    For i = 1 To Len_text
        
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
        
    Next i

End Function

Sub Engine_Init_FontTextures()
On Error GoTo eDebug:
'*****************************************************************
'Init the custom font textures
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures
'*****************************************************************
    Dim i As Long
    Dim TexInfo As D3DXIMAGE_INFO_A

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
    For i = 1 To UBound(cfonts)
    'Set the texture
    Set cfonts(i).Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, Game.path(Graficos) & "font" & i & ".bmp", D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, _
            &HFF000000, ByVal 0, ByVal 0)
    'Store the size of the texture
    cfonts(i).TextureSize.X = TexInfo.Width
    cfonts(i).TextureSize.Y = TexInfo.Height
    Next
    Exit Sub
eDebug:
    If Err.number = "-2005529767" Then
        MsgBox "Error en la textura de fuente utilizada " & Game.path(Graficos) & "Font.png.", vbCritical
        End
    End If
    End

End Sub

Sub Engine_Init_FontSettings()
    '*****************************************************************
    'Init the custom font settings
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontSettings
    '*****************************************************************
    Dim FileNum  As Byte
    Dim LoopChar As Long
    Dim Row      As Single
    Dim u        As Single
    Dim v        As Single
    Dim i As Long
    '*** Default font ***

    'Load the header information
    FileNum = FreeFile
    For i = 1 To UBound(cfonts)
        Open Game.path(Extras) & "Font" & i & ".dat" For Binary As #FileNum
        Get #FileNum, , cfonts(i).HeaderInfo
        Close #FileNum
        
        'Calculate some common values
        cfonts(i).CharHeight = cfonts(i).HeaderInfo.CellHeight - 4
        cfonts(i).RowPitch = cfonts(i).HeaderInfo.BitmapWidth \ cfonts(i).HeaderInfo.CellWidth
        cfonts(i).ColFactor = cfonts(i).HeaderInfo.CellWidth / cfonts(i).HeaderInfo.BitmapWidth
        cfonts(i).RowFactor = cfonts(i).HeaderInfo.CellHeight / cfonts(i).HeaderInfo.BitmapHeight
        
        'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
        For LoopChar = 0 To 255
            
            'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
            Row = (LoopChar - cfonts(i).HeaderInfo.BaseCharOffset) \ cfonts(i).RowPitch
            u = ((LoopChar - cfonts(i).HeaderInfo.BaseCharOffset) - (Row * cfonts(i).RowPitch)) * cfonts(i).ColFactor
            v = Row * cfonts(i).RowFactor
    
            'Set the verticies
            With cfonts(i).HeaderInfo.CharVA(LoopChar)
                .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
                .Vertex(0).rhw = 1
                .Vertex(0).tu = u
                .Vertex(0).tv = v
                .Vertex(0).X = 0
                .Vertex(0).Y = 0
                .Vertex(0).z = 0

                .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
                .Vertex(1).rhw = 1
                .Vertex(1).tu = u + cfonts(1).ColFactor
                .Vertex(1).tv = v
                .Vertex(1).X = cfonts(1).HeaderInfo.CellWidth
                .Vertex(1).Y = 0
                .Vertex(1).z = 0

                .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
                .Vertex(2).rhw = 1
                .Vertex(2).tu = u
                .Vertex(2).tv = v + cfonts(1).RowFactor
                .Vertex(2).X = 0
                .Vertex(2).Y = cfonts(1).HeaderInfo.CellHeight
                .Vertex(2).z = 0

                .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
                .Vertex(3).rhw = 1
                .Vertex(3).tu = u + cfonts(1).ColFactor
                .Vertex(3).tv = v + cfonts(1).RowFactor
                .Vertex(3).X = cfonts(1).HeaderInfo.CellWidth
                .Vertex(3).Y = cfonts(1).HeaderInfo.CellHeight
                .Vertex(3).z = 0
            End With
            
        Next LoopChar
    Next i
End Sub

Public Sub DrawText(ByVal X As Integer, _
                    ByVal Y As Integer, _
                    ByVal Text As String, _
                    ByVal Color As Long, _
                    Optional Center As Boolean = False, 
                    Optional Font As Integer = 1)

    Dim aux As D3DCOLORVALUE

    'Obtener_RGB Color, r, g, b
    Call ARGBtoD3DCOLORVALUE(Color, aux)
    
    Color = D3DColorARGB(255, aux.r, aux.g, aux.B) 

    Call Engine_Render_Text(cfonts(1), Text, X, Y, Color, Center, 255, False) 

End Sub

