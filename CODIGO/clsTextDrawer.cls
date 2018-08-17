VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECTO ' xD
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, _
    ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, _
    ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, _
    ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As size) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function InvalidateRect Lib "user32" ( _
    ByVal hWnd As Long, lpRect As RECTO, ByVal bErase As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long

Const WM_PAINT = &HF

Const TEXT_TRANSPARENT = 1
Const TEXT_OPAQUE = 2

Const FW_DONTCARE = 0
Const FW_THIN = 100
Const FW_EXTRALIGHT = 200
Const FW_ULTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_NORMAL = 400
Const FW_REGULAR = 400
Const FW_MEDIUM = 500
Const FW_SEMIBOLD = 600
Const FW_DEMIBOLD = 600
Const FW_BOLD = 700
Const FW_EXTRABOLD = 800
Const FW_ULTRABOLD = 800
Const FW_BLACK = 900
Const FW_HEAVY = 900

Const ANSI_CHARSET = 0
Const DEFAULT_CHARSET = 1
Const SYMBOL_CHARSET = 2
Const SHIFTJIS_CHARSET = 128
Const OEM_CHARSET = 255

Const OUT_TT_ONLY_PRECIS = 7
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_LH_ANGLES = &H10
Const PROOF_QUALITY = 2
Const TRUETYPE_FONTTYPE = &H4

Const TEXTBUFFER_SIZE = 1024

Private Type TextInfo
    X As Long
    Y As Long
    color As Long
    centered As Boolean
    Text As String
End Type

Private Type size
    cx As Long
    cy As Long
End Type

Private TextBuffer(TEXTBUFFER_SIZE) As TextInfo
Private textCount As Long

Private lngFont As Long
Private oldFont As Long
Private angle As Long
Private TeTireUnPaint As Boolean

Const MAX_COLORS = 255

Private DirectDraw As DirectDraw7
Private TextSurfaces(MAX_COLORS) As Long
Private Surface(MAX_COLORS) As DirectDrawSurface7
Private CantColors As Long
Private surfacesize As size
Private letters As String
Private lettersrect(255) As RECT

Public Function InitText(DD As DirectDraw7)
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

Dim ret As size
Dim tsize As Integer
Dim i As Long
Set DirectDraw = DD
Dim hdc As Long

hdc = GetDC(frmMain.hWnd)
lngFont = CreateFont(13, 0, 0, 0, FW_BOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "Tahoma") ' Hardocodeo la que usan siempre, la de frmMain.
oldFont = SelectObject(hdc, lngFont)
tsize = 0
letters = ""

For i = 32 To 255
        letters = letters & Chr(i)
        Call GetTextExtentPoint32(hdc, Chr(i), 1, ret)
        lettersrect(i).Left = tsize
        lettersrect(i).Right = tsize + ret.cx
        lettersrect(i).Top = 0
        lettersrect(i).Bottom = 13
        tsize = tsize + ret.cx
        If i = 126 Then i = 160
Next i

    surfacesize.cy = ret.cy
    surfacesize.cx = tsize
    If oldFont <> 0 Then
        Call SelectObject(hdc, oldFont)
    End If
    
DeleteObject lngFont
Call ReleaseDC(frmMain.hWnd, hdc)

Set Surface(0) = CreateSurface(0) ' Creo el color Negro
CantColors = 1
End Function

Private Function GetTextSurface(color As Long) As DirectDrawSurface7
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************
Dim i As Long

For i = 0 To CantColors
    If TextSurfaces(i) = color Then
        Set GetTextSurface = Surface(i)
        Exit Function
    End If
Next i

If CantColors < MAX_COLORS Then 'creo la nueva surface
    Set Surface(CantColors) = CreateSurface(color)
    TextSurfaces(CantColors) = color
    Set GetTextSurface = Surface(CantColors)
    CantColors = CantColors + 1
End If

End Function

Private Function CreateSurface(ByVal color As Long) As DirectDrawSurface7
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************
On Error GoTo ErrHandler

    Dim newSurface As DirectDrawSurface7
    Dim ddsd As DDSURFACEDESC2
    Dim ddck As DDCOLORKEY
    Dim bkcolor As Long
    Dim hdc As Long
    
     'Set up the surface desc
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    ddsd.lWidth = surfacesize.cx
    ddsd.lHeight = surfacesize.cy
        
    Set newSurface = DirectDraw.CreateSurface(ddsd)
    bkcolor = &H0
    If color = vbBlack Then bkcolor = &H10101
    
    'Dibujo el texto
    hdc = newSurface.GetDC
    lngFont = CreateFont(13, 0, 0, 0, FW_BOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "Tahoma") ' Hardocodeo la que usan siempre, la de frmMain.
    oldFont = SelectObject(hdc, lngFont)
    SetBkMode hdc, TEXT_OPAQUE
    SetBkColor hdc, bkcolor
    SetTextColor hdc, color
    TextOut hdc, 0, 0, letters, Len(letters)
    If oldFont <> 0 Then Call SelectObject(hdc, oldFont)
    DeleteObject lngFont
    newSurface.ReleaseDC hdc
    'Podria haber usado el DrawText del dx aca, pero soy jodido
    
    ddck.high = bkcolor
    ddck.low = bkcolor
    Call newSurface.SetColorKey(DDCKEY_SRCBLT, ddck)
    Set CreateSurface = newSurface
    
Exit Function

ErrHandler:
    If Err.number = DDERR_OUTOFMEMORY Or Err.number = DDERR_OUTOFVIDEOMEMORY Then
            MsgBox "No hay memoria disponible! El programa abortará. Cierra algunos programas e intenta de nuevo"
            End
    Else
        MsgBox "Un error inesperado ocurrió al intentar crear el texto" & ". " & vbCrLf & _
                "El código de error es " & CStr(Err.number) & " - " & Err.Description & vbCrLf & vbCrLf & "Copia este mensaje y notifica a los administradores.", _
                vbOKOnly Or vbCritical Or vbExclamation, "Error"
        End
    End If
End Function

Public Function DrawText(ByVal X As Long, ByVal Y As Long, ByVal Text As String, ByVal color As Long, ByVal Surface As DirectDrawSurface7)
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************
    Dim i As Integer
    Dim tx As Long
    Dim textSurface As DirectDrawSurface7
    Dim bytestring() As Byte
    If X > 120 And X < 799 And Y > 246 And Y < 668 Then
        bytestring = StrConv(Text, vbFromUnicode)
        tx = X
        Set textSurface = GetTextSurface(color)
        For i = 0 To UBound(bytestring) ' - 1
            Call Surface.BltFast(tx, Y, textSurface, lettersrect(bytestring(i)), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            tx = tx + lettersrect(bytestring(i)).Right - lettersrect(bytestring(i)).Left
        Next i
    End If
End Function

Public Function DrawTextToDC(hdc As Long, Optional FontSize As Long = 13) As Long
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

'Forma alternativa de dibujar texto con GDI
'La forma de usarla es llamando a la funcion AddText de ahi abajo
Dim i As Long
Dim lastcolor As Long
Dim ret As size
Dim tRect As RECTO

'Aplico mi font
lngFont = CreateFont(FontSize, 0, 0, 0, FW_BOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "Tahoma") ' Hardocodeo la que usan siempre, la de frmMain.
oldFont = SelectObject(hdc, lngFont)

'Le doy fondo transparente
SetBkMode hdc, TEXT_TRANSPARENT

'Primero dibujo toda la capa en negro
SetTextColor hdc, vbBlack
For i = 0 To textCount - 1
    With (TextBuffer(i))
        If .centered Then ' ya lo dejo acomodado
            Call GetTextExtentPoint32(hdc, .Text, Len(.Text), ret)
            .X = .X - ret.cx \ 2
        End If
        TextOut hdc, .X - 2, .Y - 1, .Text, Len(.Text)
    End With
Next i

'Y ahora toda la capa en color
For i = 0 To textCount - 1
    With (TextBuffer(i))
        If .color <> lastcolor Then ' si es igual me ahorro una llamada a un api, costo beneficio me va
            SetTextColor hdc, .color
            lastcolor = .color
        End If
        TextOut hdc, .X, .Y, .Text, Len(.Text)
    End With
Next i


textCount = 0 ' Se resetea
If oldFont <> 0 Then Call SelectObject(hdc, oldFont)
DeleteObject lngFont

DrawTextToDC = 0
End Function

Public Function AddText(xPos As Integer, yPos As Integer, fontcolor As Long, strText As String, Optional centered As Boolean = False) As Long
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

If strText <> "" Then
    If xPos > 120 And xPos < 799 And yPos > 246 And yPos < 668 Then 'te la hago corta, mas alla de esto no lo veo.
        If textCount < TEXTBUFFER_SIZE Then
            With TextBuffer(textCount)
                .X = xPos
                .Y = yPos
                .color = fontcolor
                .Text = strText
                .centered = centered
            End With
            textCount = textCount + 1
            AddText = textCount
        Else
            AddText = -1
        End If
    End If
End If

End Function

Private Sub Class_Terminate()
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

Dim i As Long

For i = 0 To CantColors
    Set Surface(i) = Nothing
Next i

End Sub

Public Function DrawDice(Optional refresh As Boolean = True) As Long
'**************************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modify Date: 07/31/2010
'
'**************************************************************

    Dim i As Long
    Dim hdc As Long
    Dim tesmpstring As String
    
    Dim thdc As Long
    Dim TempRect As RECTO

    If refresh = False Then
        angle = RandomNumber(0, 1) * (3440) + RandomNumber(0, 160)
    End If

    TempRect.Top = 218
    TempRect.Bottom = 338
    TempRect.Left = 295
    TempRect.Right = 380
    
    Call InvalidateRect(frmCrearPersonaje.hWnd, TempRect, 0)
    TeTireUnPaint = True
    
    Call PostMessage(frmCrearPersonaje.hWnd, WM_PAINT, 0, 0)
    DoEvents ' Como es esto que usa un doevents afuera de un loop? jajaja
    
    hdc = GetDC(frmCrearPersonaje.hWnd)
    
    lngFont = CreateFont(13, 0, angle, 0, FW_BOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "MS Sans Serif")
    oldFont = SelectObject(hdc, lngFont)


    SetBkMode hdc, TEXT_TRANSPARENT
    SetBkColor hdc, vbBlack
    SetTextColor hdc, vbWhite
    
    For i = 1 To NUMATRIBUTOS
        TextOut hdc, frmCrearPersonaje.lblAtributos(i).Left, frmCrearPersonaje.lblAtributos(i).Top, frmCrearPersonaje.lblAtributos(i).Caption, Len(frmCrearPersonaje.lblAtributos(i).Caption)
        TextOut hdc, frmCrearPersonaje.lblAtributoFinal(i).Left, frmCrearPersonaje.lblAtributoFinal(i).Top, frmCrearPersonaje.lblAtributoFinal(i).Caption, Len(frmCrearPersonaje.lblAtributoFinal(i).Caption)
    Next i

    If oldFont <> 0 Then
        Call SelectObject(hdc, oldFont)
    End If
    
    DeleteObject lngFont
    ReleaseDC frmCrearPersonaje.hWnd, hdc
  
End Function

Public Property Get SentPaint() As Boolean
    SentPaint = TeTireUnPaint
    TeTireUnPaint = False
End Property
