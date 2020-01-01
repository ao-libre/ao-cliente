Attribute VB_Name = "mDx8_Dibujado"
Option Explicit

' Dano en Render
Private Const DAMAGE_TIME As Integer = 500
Private Const DAMAGE_OFFSET As Integer = 20
Private Const DAMAGE_FONT_S As Byte = 12
 
Private Enum EDType
     edPunal = 1    'Apunalo.
     edNormal = 2   'Hechizo o golpe com�n.
     edCritico = 3  'Golpe Critico
     edFallo = 4    'Fallo el ataque
     edCurar = 5    'Curacion a usuario
     edTrabajo = 6  'Cantidad de items obtenidas a partir del trabajo realizado
End Enum
 
Private DNormalFont    As New StdFont
 
Type DList
     DamageVal      As Integer      'Cantidad de da�o.
     ColorRGB       As Long         'Color.
     DamageType     As EDType       'Tipo, se usa para saber si es apu o no.
     DamageFont     As New StdFont  'Efecto del apu.
     StartedTime    As Long         'Cuando fue creado.
     Downloading    As Byte         'Contador para la posicion Y.
     Activated      As Boolean      'Si esta activado..
End Type

Sub DrawGrhtoHdc(ByRef Pic As PictureBox, _
                 ByVal GrhIndex As Integer, _
                 ByRef DestRect As RECT)

    '*****************************************************************
    'Draws a Grh's portion to the given area of any Device Context
    '*****************************************************************
         
    DoEvents
        
    'Clear the inventory window
    Call Engine_BeginScene
        
    Call Draw_GrhIndex(GrhIndex, 0, 0, 0, Normal_RGBList())
        
    Call Engine_EndScene(DestRect, Pic.hWnd)
        
End Sub

Sub Damage_Initialize()

    ' Inicializamos el dano en render
    With DNormalFont
        .Size = 20
        .italic = False
        .bold = False
        .name = "Tahoma"
    End With

End Sub

Sub Damage_Create(ByVal X As Byte, _
                  ByVal Y As Byte, _
                  ByVal ColorRGB As Long, _
                  ByVal DamageValue As Integer, _
                  ByVal edMode As Byte)
 
    ' @ Agrega un nuevo dano.
 
    With MapData(X, Y).Damage
     
        .Activated = True
        .ColorRGB = ColorRGB
        .DamageType = edMode
        .DamageVal = DamageValue
        .StartedTime = GetTickCount
        .Downloading = 0
     
        Select Case .DamageType
        
            Case EDType.edPunal

                With .DamageFont
                    .Size = Val(DAMAGE_FONT_S)
                    .name = "Tahoma"
                    .bold = False
                    Exit Sub

                End With
            
        End Select
     
        .DamageFont = DNormalFont
        .DamageFont.Size = 14
     
    End With
 
End Sub

Private Function EaseOutCubic(Time As Double)
    Time = Time - 1
    EaseOutCubic = Time * Time * Time + 1
End Function
 
Sub Damage_Draw(ByVal X As Byte, _
                ByVal Y As Byte, _
                ByVal PixelX As Integer, _
                ByVal PixelY As Integer)
 
    ' @ Dibuja un dano
 
    With MapData(X, Y).Damage
     
        If (Not .Activated) Or (Not .DamageVal <> 0) Then Exit Sub
        
        Dim ElapsedTime As Long
        ElapsedTime = GetTickCount - .StartedTime
        
        If ElapsedTime < DAMAGE_TIME Then
           
            .Downloading = EaseOutCubic(ElapsedTime / DAMAGE_TIME) * DAMAGE_OFFSET
           
            .ColorRGB = Damage_ModifyColour(.DamageType)
           
            'Efectito para el apu
            If .DamageType = EDType.edPunal Then
                .DamageFont.Size = Damage_NewSize(ElapsedTime)

            End If
               
            'Dibujo
            Select Case .DamageType
            
                Case EDType.edCritico
                    DrawText PixelX, PixelY - .Downloading, .DamageVal & "!!", .ColorRGB
                
                Case EDType.edCurar
                    DrawText PixelX, PixelY - .Downloading, "+" & .DamageVal, .ColorRGB
                
                Case EDType.edTrabajo
                    DrawText PixelX, PixelY - .Downloading, "+" & .DamageVal, .ColorRGB
                    
                Case EDType.edFallo
                    DrawText PixelX, PixelY - .Downloading, "Fallo", .ColorRGB
                    
                Case Else 'EDType.edNormal
                    DrawText PixelX, PixelY - .Downloading, "-" & .DamageVal, .ColorRGB
                    
            End Select
            
        'Si llego al tiempo lo limpio
        Else
            Damage_Clear X, Y
           
        End If
       
    End With
 
End Sub
 
Sub Damage_Clear(ByVal X As Byte, ByVal Y As Byte)
 
    ' @ Limpia todo.
 
    With MapData(X, Y).Damage
        .Activated = False
        .ColorRGB = 0
        .DamageVal = 0
        .StartedTime = 0

    End With
 
End Sub
 
Function Damage_ModifyColour(ByVal DamageType As Byte) As Long
 
    ' @ Se usa para el "efecto" de desvanecimiento.
 
    Select Case DamageType
                   
        Case EDType.edPunal
            Damage_ModifyColour = ColoresDano(52)
            
        Case EDType.edFallo
            Damage_ModifyColour = ColoresDano(54)
            
        Case EDType.edCurar
            Damage_ModifyColour = ColoresDano(55)
        
        Case EDType.edTrabajo
            Damage_ModifyColour = ColoresDano(56)
            
        Case Else 'EDType.edNormal
            Damage_ModifyColour = ColoresDano(51)
            
    End Select
 
End Function
 
Function Damage_NewSize(ByVal ElapsedTime As Byte) As Byte
 
    ' @ Se usa para el "efecto" del apu.

    ' Nos basamos en la constante DAMAGE_TIME
    Select Case ElapsedTime
 
        Case Is <= DAMAGE_TIME / 5
            Damage_NewSize = 14
       
        Case Is <= DAMAGE_TIME * 2 / 5
            Damage_NewSize = 13
           
        Case Is <= DAMAGE_TIME * 3 / 5
            Damage_NewSize = 12
           
        Case Else
            Damage_NewSize = 11
       
    End Select
 
End Function
