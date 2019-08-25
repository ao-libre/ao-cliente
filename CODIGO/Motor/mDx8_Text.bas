Attribute VB_Name = "mDx8_Text"
Option Explicit

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       source As Any, _
                                       ByVal Length As Long)
    
Private Type CharVA
    X As Integer
    Y As Integer
    W As Integer
    H As Integer
        
    Tx1 As Single
    Tx2 As Single
    Ty1 As Single
    Ty2 As Single
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
    Dim Red        As Integer, Blue As Integer, Green As Integer
    
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

Public Sub Text_Render_Special(ByVal intX As Integer, _
                               ByVal intY As Integer, _
                               ByRef strText As String, _
                               ByRef lngColor() As Long, _
                               Optional bolCentred As Boolean = False)  ' GSZAO
    '*****************************************************************
    'Text_Render_Special by ^[GS]^
    '*****************************************************************
    
    If LenB(strText) <> 0 Then
        
        #If SpriteBatch = 1 Then
            Call Engine_Render_Text(cfonts(1), strText, intX, intY, lngColor, bolCentred)
        #Else
            Call Engine_Render_Text(cfonts(1), strText, intX, intY, lngColor, bolCentred, , , SpriteBatch)
        #End If
        
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

Private Sub Engine_Render_Text(ByRef UseFont As CustomFont, _
                               ByVal Text As String, _
                               ByVal X As Long, _
                               ByVal Y As Long, _
                               ByRef Color() As Long, _
                               Optional ByVal Center As Boolean = False, _
                               Optional ByVal Alpha As Byte = 255, _
                               Optional ByVal ParseEmoticons As Boolean = True, _
                               Optional ByRef Batch As clsBatch)
    '*****************************************************************
    'Render text with a custom font
    '*****************************************************************

    Dim TempVA As CharVA
    Dim tempstr()     As String
    Dim Count         As Integer
    Dim ascii()       As Byte
    Dim i             As Long
    Dim J             As Long
    Dim TempColor     As Long
    Dim ResetColor    As Byte
    Dim YOffset       As Single
    
    Dim Upper_tempstr As Long, Len_tempstr As Long
    
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
        
        'pre-calculate tempstr's upperbound to improve performance
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
    Call Batch.SetTexture(UseFont.Texture)
    
    If Center Then
        X = X - Engine_GetTextWidth(cfonts(1), Text) * 0.5
    End If
    
    'pre-calculate tempstr's upperbound to improve performance
    Upper_tempstr = UBound(tempstr)
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To Upper_tempstr

        If Len(tempstr(i)) > 0 Then
            YOffset = i * UseFont.CharHeight
            Count = 0
        
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
        
            'Loop through the characters
            For J = 1 To Len(tempstr(i))

                CopyMemory TempVA, UseFont.HeaderInfo.CharVA(ascii(J - 1)), 24 'this number represents the size of "CharVA" struct
                
                TempVA.X = X + Count
                TempVA.Y = Y + YOffset
            
                Call SpriteBatch.Draw(TempVA.X, TempVA.Y, TempVA.W, TempVA.H, Color(), TempVA.Tx1, TempVA.Ty1, TempVA.Tx2, TempVA.Ty2)

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

Public Function ARGB(ByVal r As Long, _
                     ByVal g As Long, _
                     ByVal B As Long, _
                     ByVal A As Long) As Long
        
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

Private Function Engine_GetTextWidth(ByRef UseFont As CustomFont, _
                                     ByVal Text As String) As Integer
    '***************************************************
    'Returns the width of text
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth
    '***************************************************
    Dim i        As Integer
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
    Dim TexInfo As D3DXIMAGE_INFO_A

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
    
    'Set the texture
    Set cfonts(1).Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, Game.path(Graficos) & "Font.png", D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, 0, TexInfo, ByVal 0)
    
    'Store the size of the texture
    cfonts(1).TextureSize.X = TexInfo.Width
    cfonts(1).TextureSize.Y = TexInfo.Height
    
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

    '*** Default font ***

    'Load the header information
    FileNum = FreeFile
    Open App.path & "\Extras\" & "Font.dat" For Binary As #FileNum
    Get #FileNum, , cfonts(1).HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    cfonts(1).CharHeight = cfonts(1).HeaderInfo.CellHeight - 4
    cfonts(1).RowPitch = cfonts(1).HeaderInfo.BitmapWidth \ cfonts(1).HeaderInfo.CellWidth
    cfonts(1).ColFactor = cfonts(1).HeaderInfo.CellWidth / cfonts(1).HeaderInfo.BitmapWidth
    cfonts(1).RowFactor = cfonts(1).HeaderInfo.CellHeight / cfonts(1).HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) \ cfonts(1).RowPitch
        u = ((LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) - (Row * cfonts(1).RowPitch)) * cfonts(1).ColFactor
        v = Row * cfonts(1).RowFactor
            
        'Set the verticies
        With cfonts(1).HeaderInfo.CharVA(LoopChar)
            .X = 0
            .Y = 0
            .W = cfonts(1).HeaderInfo.CellWidth
            .H = cfonts(1).HeaderInfo.CellHeight
            .Tx1 = u
            .Ty1 = v
            .Tx2 = u + cfonts(1).ColFactor
            .Ty2 = v + cfonts(1).RowFactor
        End With
        
    Next LoopChar

End Sub

Public Sub DrawText(ByVal X As Integer, _
                    ByVal Y As Integer, _
                    ByVal Text As String, _
                    ByRef Color() As Long, _
                    Optional ByVal Center As Boolean = False)

    Call Engine_Render_Text(cfonts(1), Text, X, Y, Color, Center, 255, False, SpriteBatch)

End Sub
