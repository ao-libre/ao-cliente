Attribute VB_Name = "Carteles"
Option Explicit

Const XPosCartel = 100
Const YPosCartel = 100
Const MAXLONG = 30

Public Cartel              As Boolean
Public Leyenda             As String
Public LeyendaFormateada() As String
Public textura             As Integer

Sub InitCartel(Ley As String, Grh As Integer)
    
    On Error GoTo InitCartel_Err
    

    If Not Cartel Then
        Leyenda = Ley
        textura = Grh
        Cartel = True
        ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                
        Dim i As Integer, k As Integer, anti As Integer
        anti = 1
        k = 0
        i = 0
        Call DarFormato(Leyenda, i, k, anti)
        i = 0

        Do While LeyendaFormateada(i) <> "" And i < UBound(LeyendaFormateada)
        
            i = i + 1
        Loop
        ReDim Preserve LeyendaFormateada(0 To i)
    Else
        Exit Sub

    End If

    
    Exit Sub

InitCartel_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Carteles" & "->" & "InitCartel"
    End If
Resume Next
    
End Sub

Private Function DarFormato(s As String, i As Integer, k As Integer, anti As Integer)
    
    On Error GoTo DarFormato_Err
    

    If anti + i <= Len(s) + 1 Then
        If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
            LeyendaFormateada(k) = mid$(s, anti, i + 1)
            k = k + 1
            anti = anti + i + 1
            i = 0
        Else
            i = i + 1

        End If

        Call DarFormato(s, i, k, anti)

    End If

    
    Exit Function

DarFormato_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Carteles" & "->" & "DarFormato"
    End If
Resume Next
    
End Function

Sub DibujarCartel()
    
    On Error GoTo DibujarCartel_Err
    
    Dim X As Integer, Y As Integer

    If Not Cartel Then Exit Sub
    
    X = XPosCartel + 20
    Y = YPosCartel + 20
    
    Call DDrawTransGrhIndextoSurface(textura, XPosCartel, YPosCartel, 0, Normal_RGBList(), 0, False)
    Dim j As Integer, desp As Integer
    
    For j = 0 To UBound(LeyendaFormateada)
        'Fonts_Render_String LeyendaFormateada(j), X, Y + desp, -1, Settings.Engine_Font
        DrawText X, Y + desp, LeyendaFormateada(j), -1
        desp = desp + (frmMain.Font.Size) + 5
    Next

    
    Exit Sub

DibujarCartel_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Carteles" & "->" & "DibujarCartel"
    End If
Resume Next
    
End Sub

