Attribute VB_Name = "Carteles"
Option Explicit

Const XPosCartel = 100
Const YPosCartel = 100
Const MAXLONG = 30

Public Cartel As Boolean
Public Leyenda As String
Public LeyendaFormateada() As String
Public textura As Long

Sub InitCartel(Ley As String, Grh As Long)
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
End Sub

Private Function DarFormato(s As String, i As Integer, k As Integer, anti As Integer)
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
End Function

Sub DibujarCartel()
    If Not Cartel Then Exit Sub

    Dim X As Integer, Y As Integer
    
    X = XPosCartel + 20
    Y = YPosCartel + 20
    
    Call Draw_GrhIndex(textura, XPosCartel, YPosCartel, 0, Normal_RGBList(), 0, False)
    Dim J As Integer, desp As Integer, Upper_leyendaFormateada As Long
    
    Upper_leyendaFormateada = UBound(LeyendaFormateada)
    
    For J = 0 To Upper_leyendaFormateada
        'Fonts_Render_String LeyendaFormateada(j), X, Y + desp, -1, Settings.Engine_Font
        Call DrawText(X, Y + desp, LeyendaFormateada(J), -1)
        desp = desp + (frmMain.Font.Size) + 5
    Next
End Sub



