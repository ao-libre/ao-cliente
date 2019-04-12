Attribute VB_Name = "mDx8_Cuenta"
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 23/12/10
'Blisse-AO | Sistema de Cuenta en Screen
'***************************************************
Option Explicit

Public Type Count
    Min As Byte
    
    TickCount As Long
    DoIt As Boolean
End Type

Public DX_Count As Count

'Fuente de Cuenta
'Public Const Count_Font As Byte = 8


Public Sub RenderCount()
    '   Si no hay cuenta no rompemos mas el render
    If DX_Count.DoIt = False Then Exit Sub

    If DX_Count.Min <> 0 Then
        '   Si no es 0 Dibujamos normal
        'Call Fonts_Render_String(DX_Count.min, (ScreenWidth - Fonts_Render_String_Width(DX_Count.min, Count_Font)) / 2, (ScreenHeight - Fuentes(Count_Font).CharactersHeight) / 2, -1, Count_Font)
    Else
        'Si es 0, Dibujamos el "@" que en la fuente esta puesto como el "YA!"
        'Call Fonts_Render_String("@", (ScreenWidth - Fonts_Render_String_Width("@", Count_Font)) / 2, (ScreenHeight - Fuentes(Count_Font).CharactersHeight) / 2, -1, Count_Font)
    End If
    
    'Checkeamos la cuenta, si es necesario restamos valor
    Call CheckCount
End Sub

Public Sub CheckCount()
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 23/12/10
'Check the count
'***************************************************

    '   Si no hay cuenta no rompemos mas
    If DX_Count.DoIt = False Then Exit Sub
    
        If GetTickCount - DX_Count.TickCount > 1000 Then
        '   Nos fijamos que haya pasado el tiempo
            If DX_Count.Min > 0 Then
            '   Si es mayor a 0 restamos
                DX_Count.Min = DX_Count.Min - 1
                DX_Count.TickCount = GetTickCount
            ElseIf DX_Count.Min = 0 Then
                '   Si es 0 quitamos la cuenta
                DX_Count.Min = 0
                DX_Count.DoIt = False
            End If
        End If
End Sub

Public Sub InitCount(ByVal Max As Byte)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 23/12/10
'Check the count
'***************************************************

    '   Si hay cuenta no rompemos a la actual
    If DX_Count.DoIt = True Then Exit Sub
    
    With DX_Count
        '   Seteamos el Min
        .Min = Max
        
        '   Seteamos el tiempo
        .TickCount = GetTickCount
        
        '   Y entonces... DO IT!
        .DoIt = True
    End With
End Sub
