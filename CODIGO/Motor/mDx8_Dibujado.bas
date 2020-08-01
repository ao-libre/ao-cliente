Attribute VB_Name = "mDx8_Dibujado"
Option Explicit

' Dano en Render
Private Const DAMAGE_TIME As Integer = 1000
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

Private DrawBuffer As cDIBSection

Sub DrawGrhtoHdc(ByRef Pic As PictureBox, _
                 ByVal GrhIndex As Long, _
                 ByRef DestRect As RECT)

    '*****************************************************************
    'Draws a Grh's portion to the given area of any Device Context
    '*****************************************************************
         
    DoEvents
    
    Pic.AutoRedraw = False
        
    'Clear the inventory window
    Call Engine_BeginScene
        
    Call Draw_GrhIndex(GrhIndex, 0, 0, 0, Normal_RGBList())
        
    Call Engine_EndScene(DestRect, Pic.hWnd)
    
    Call DrawBuffer.LoadPictureBlt(Pic.hdc)

    Pic.AutoRedraw = True

    Call DrawBuffer.PaintPicture(Pic.hdc, 0, 0, Pic.Width, Pic.Height, 0, 0, vbSrcCopy)

    Pic.Picture = Pic.Image
        
End Sub

Public Sub PrepareDrawBuffer()
    Set DrawBuffer = New cDIBSection
    'El tamanio del buffer es arbitrario = 1024 x 1024
    Call DrawBuffer.Create(1024, 1024)
End Sub

Public Sub CleanDrawBuffer()
    Set DrawBuffer = Nothing
End Sub

Public Sub DrawPJ(ByVal Index As Byte)

    If LenB(cPJ(Index).Nombre) = 0 Then Exit Sub
    DoEvents
    
    Dim cColor       As Long
    Dim Head_OffSet  As Integer
    Dim PixelOffsetX As Integer
    Dim PixelOffsetY As Integer
    Dim RE           As RECT
    
    If cPJ(Index).GameMaster Then
        cColor = 2004510
    Else
        cColor = IIf(cPJ(Index).Criminal, 255, 16744448)
    End If
    
    With frmPanelAccount.lblAccData(Index)
        .Caption = cPJ(Index).Nombre
        .ForeColor = cColor
    End With
    
    With frmPanelAccount.picChar(Index - 1)
        RE.Left = 0
        RE.Top = 0
        RE.Bottom = .Height
        RE.Right = .Width
    End With

    PixelOffsetX = RE.Right \ 2 - 16
    PixelOffsetY = RE.Bottom \ 2
    
    Call Engine_BeginScene
    
    With cPJ(Index)
    
        If .Body <> 0 Then

            Call Draw_Grh(BodyData(.Body).Walk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)

            If .Head <> 0 Then
                Call Draw_Grh(HeadData(.Head).Head(3), PixelOffsetX + BodyData(.Body).HeadOffset.X, PixelOffsetY + BodyData(.Body).HeadOffset.Y, 1, Normal_RGBList(), 0)
            End If

            If .helmet <> 0 Then
                Call Draw_Grh(CascoAnimData(.helmet).Head(3), PixelOffsetX + BodyData(.Body).HeadOffset.X, PixelOffsetY + BodyData(.Body).HeadOffset.Y + OFFSET_HEAD, 1, Normal_RGBList(), 0)
            End If

            If .weapon <> 0 Then
                Call Draw_Grh(WeaponAnimData(.weapon).WeaponWalk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)
            End If

            If .shield <> 0 Then
                Call Draw_Grh(ShieldAnimData(.shield).ShieldWalk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)
            End If
        
        End If
    
    End With

    Call Engine_EndScene(RE, frmPanelAccount.picChar(Index - 1).hWnd)

    Call DrawBuffer.LoadPictureBlt(frmPanelAccount.picChar(Index - 1).hdc)

    frmPanelAccount.picChar(Index - 1).AutoRedraw = True

    Call DrawBuffer.PaintPicture(frmPanelAccount.picChar(Index - 1).hdc, 0, 0, RE.Right, RE.Bottom, 0, 0, vbSrcCopy)

    frmPanelAccount.picChar(Index - 1).Picture = frmPanelAccount.picChar(Index - 1).Image
    
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
                    Call DrawText(PixelX, PixelY - .Downloading, .DamageVal & "!!", .ColorRGB)
                
                Case EDType.edCurar
                    Call DrawText(PixelX, PixelY - .Downloading, "+" & .DamageVal, .ColorRGB)
                
                Case EDType.edTrabajo
                    Call DrawText(PixelX, PixelY - .Downloading, "+" & .DamageVal, .ColorRGB)
                    
                Case EDType.edFallo
                    Call DrawText(PixelX, PixelY - .Downloading, "Fallo", .ColorRGB)
                    
                Case Else 'EDType.edNormal
                    Call DrawText(PixelX, PixelY - .Downloading, "-" & .DamageVal, .ColorRGB)
                    
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
 
Function Damage_NewSize(ByVal ElapsedTime As Long) As Byte
 
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
