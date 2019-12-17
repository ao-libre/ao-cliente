Attribute VB_Name = "mDx8_Dibujado"
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef source As Any, ByVal Length As Long)

Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

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

Public Sub ArrayToPicturePNG(ByRef byteArray() As Byte, ByRef imgDest As IPicture) ' GSZAO
    Call SetBitmapBits(imgDest.handle, UBound(byteArray), byteArray(0))
End Sub

Public Function ArrayToPicture(inArray() As Byte, Offset As Long, Size As Long) As IPicture
    
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(Offset), Size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                  Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), ArrayToPicture)
            End If
        End If
    End If
End Function

Sub DrawGrhtoHdc(ByVal desthDC As Long, ByVal grh_index As Integer, ByRef SourceRect As RECT, ByRef destRect As RECT)
    On Error Resume Next
    
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim PrevObj As Long
    Dim screen_x As Integer
    Dim screen_y As Integer
    
    screen_x = destRect.Left
    screen_y = destRect.Top
    
    If grh_index <= 0 Then Exit Sub

    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If
    
    Dim data() As Byte
    Dim bmpData As StdPicture
    
    'get Picture
    If Get_Image(Game.path(Graficos), CStr(GrhData(grh_index).FileNum), data, True) Then  ' GSZAO
        Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1)
        
        src_x = GrhData(grh_index).SX
        src_y = GrhData(grh_index).SY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
        
        hdcsrc = CreateCompatibleDC(desthDC)
        PrevObj = SelectObject(hdcsrc, bmpData)
        
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
        DeleteDC hdcsrc
        
        Set bmpData = Nothing
    End If

End Sub

Sub Damage_Initialize()

    ' Inicializamos el dano en render
    With DNormalFont
        .Size = 20
        .italic = False
        .bold = False
        .Name = "Tahoma"
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
                    .Name = "Tahoma"
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
