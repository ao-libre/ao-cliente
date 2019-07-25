VERSION 5.00
Begin VB.UserControl uAOButton 
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1140
   ScaleWidth      =   2460
   ToolboxBitmap   =   "uAOButton.ctx":0000
   Begin VB.Timer MouseO 
      Interval        =   3
      Left            =   1920
      Top             =   720
   End
   Begin VB.PictureBox bButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Image iFON 
      Height          =   15
      Left            =   0
      Picture         =   "uAOButton.ctx":0312
      Top             =   1220
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image iVER 
      Height          =   15
      Left            =   960
      Picture         =   "uAOButton.ctx":0360
      Top             =   1200
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image iHOR 
      Height          =   15
      Left            =   0
      Picture         =   "uAOButton.ctx":03AE
      Top             =   1200
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image iESQ 
      Height          =   15
      Left            =   960
      Picture         =   "uAOButton.ctx":03FC
      Top             =   1220
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "uAOButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%                                                     %
'%                    AO BUTTON v1.6                   %
'%               Copyright © 2012 by ^[GS]^            %
'%                    www.GS-ZONE.org                  %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%  Este control permite hacer botones fácilmente      %
'%  sin utilizar grandes cantidades de imágenes o      %
'%  clases complicadas.                                %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%  Changelog:                                         %
'%   20/10/2012 - Se corrigió un bug al pasar el       %
'%                mouse sobre el control que perdia    %
'%                el foco de otros controles. (^[GS]^) %
'%   03/08/2012 - Hice una corrección para que         %
'%                funcione correctamente el refresco   %
'%                del control. (^[GS]^)                %
'%   31/07/2012 - Le agregue 3 colores por defecto     %
'%                al botón para cuando no tiene las    %
'%                imágenes cargadas. (^[GS]^)          %
'%   27/07/2012 - Se removió la textura del control    %
'%                para poder cargarla desde fuera      %
'%                y así no pesar al programa. (^[GS]^) %
'%   18/07/2012 - Se mejoro el Caption, ahora con      %
'%                el API DrawText, que permite usar    %
'%                autowarp en textos grandes. (^[GS]^) %
'%   15/07/2012 - Se agrego la función para cargar     %
'%                la textura de los botones de         %
'%                manera más sencilla que editando     %
'%                el control. (^[GS]^)                 %
'%              - Se redujo la cantidad de imágenes    %
'%                por botón de 12 para los 3 estados   %
'%                a solo 4 partes con los 3 estados    %
'%                auto-mapeados. (^[GS]^)              %
'%              - Se corrigieron varios                %
'%                detalles menores. (^[GS]^)           %
'%   14/07/2012 - Se termino la primera versión        %
'%                totalmente funcional. (^[GS]^)       %
'%                - Botón de 3 estados.                %
'%                - Caption configurable.              %
'%                - Funcionalidad mediante teclado.    %
'%   13/07/2012 - Se inicio el proyecto. (^[GS]^)      %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Option Explicit

Private szEsquinasW As Integer
Private szEsquinasH As Integer
Private szLineaHW   As Integer
Private szLineaHH   As Integer
Private szLineaVW   As Integer
Private szLineaVH   As Integer
Private szFondoW    As Integer
Private szFondoH    As Integer

Private iEsquinaLoaded  As Boolean
Private iLineaVLoaded   As Boolean
Private iLineaHLoaded   As Boolean
Private iFondoLoaded    As Boolean

Private iLineaH     As IPictureDisp
Private iLineaV     As IPictureDisp
Private iFondo      As IPictureDisp
Private iEsquina    As IPictureDisp

Private isOver      As Boolean
Private lastStat    As Byte
Private lastHwnd    As Long
Private lastButton  As Byte
Private lastKeyDown As Byte
Private isFocus     As Boolean

Private CaptionButton As String
Private ForeC       As Long
Private ForeCo      As Long
Private isEnabled   As Boolean
Private rc          As RECT

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOver()
Public Event MouseOut()

Private Sub bButton_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    If isEnabled Then
        RaiseEvent Click
    End If
End Sub

Private Sub bButton_DblClick()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    If isEnabled Then
        Call bButton_MouseDown(1, 0, 0, 0)
        RaiseEvent DblClick
    End If
End Sub

Private Sub bButton_GotFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    If isOver = False And isEnabled Then
        isFocus = True
        Call Redraw(1)
        'Debug.Print "bButton_GotFocus1"
    End If
End Sub

Private Sub bButton_KeyDown(KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent KeyDown(KeyCode, Shift)
    lastKeyDown = KeyCode
    Select Case KeyCode
    Case 32
        Call Redraw(2)
        'Debug.Print "bButton_KeyDown2"
    Case 39, 40
        SendKeys "{Tab}"
    Case 37, 38
        SendKeys "+{Tab}"
    End Select
End Sub

Private Sub bButton_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub bButton_KeyUp(KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent KeyUp(KeyCode, Shift)
    If (KeyCode = 32) And (lastKeyDown = 32) Then
        Call Redraw(1)
        'Debug.Print "bButton_KeyUp1"
        UserControl.Refresh
        RaiseEvent Click
    End If
End Sub

Private Sub bButton_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    If isOver = False And isEnabled Then
        isFocus = False
        Call Redraw(0)
        'Debug.Print "bButton_LostFocus0"
    End If
End Sub

Private Sub bButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    lastButton = Button
    If lastButton <> 2 And isEnabled Then
        Call Redraw(2)
        'Debug.Print "bButton_MouseDown2"
    End If
End Sub

Private Sub bButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/10/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If lastButton < 2 And isEnabled Then
        lastHwnd = bButton.hWnd
        If Not isMouseOver Then
            Call Redraw(0)
            'Debug.Print "bButton_MouseMove0"
        Else
            If Button = 0 And Not isOver Then
                MouseO.Enabled = True
                isOver = True
                Call Redraw(1)
                'Debug.Print "bButton_MouseMove1"
                RaiseEvent MouseOver
            End If
        End If
    End If
End Sub

Private Sub bButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If lastButton <> 2 And isEnabled Then
        If isOver = True Then
            Call Redraw(1)
            'Debug.Print "bButton_MouseUp1"
        Else
            Call Redraw(0)
            'Debug.Print "bButton_MouseUp0"
        End If
    End If
    lastButton = 0
End Sub

Private Sub MouseO_Timer()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    If Not isMouseOver Then
        Call Redraw(0)
        'Debug.Print "MouseO0"
        isOver = False
        isFocus = False
        RaiseEvent MouseOut
        MouseO.Enabled = False
    End If
End Sub

Private Function isMouseOver() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    Dim pt As POINTAPI
    GetCursorPos pt
    isMouseOver = (WindowFromPoint(pt.X, pt.Y) = lastHwnd)
End Function

Private Sub ReloadTextures()
'*************************************************
'Author: ^[GS]^
'Last modified: 31/07/2012
'*************************************************

On Error Resume Next
    
    Set iEsquina = iESQ.Picture
    Set iFondo = iFON.Picture
    Set iLineaH = iHOR.Picture
    Set iLineaV = iVER.Picture
    
    If iESQ.Picture.Height <> iESQ.Picture.Width Then
        szEsquinasW = iESQ.Width / 6
        szEsquinasH = iESQ.Height / 2
        iEsquinaLoaded = True
    Else
        iEsquinaLoaded = False
    End If
    If iFON.Picture.Height <> iFON.Picture.Width Then
        szFondoW = iFON.Width
        szFondoH = iFON.Height / 3
        iFondoLoaded = True
    Else
        iFondoLoaded = False
    End If
    If iHOR.Picture.Height <> iHOR.Picture.Width Then
        szLineaHW = iHOR.Width
        szLineaHH = iHOR.Height / 6
        iLineaHLoaded = True
    Else
        iLineaHLoaded = False
    End If
    If iVER.Picture.Height <> iVER.Picture.Width Then
        szLineaVW = iVER.Width / 6
        szLineaVH = iVER.Height
        iLineaVLoaded = True
    Else
        iLineaVLoaded = False
    End If
    
End Sub

Private Sub Redraw(ByVal Estado As Byte, Optional Force As Boolean = False)
'*************************************************
'Author: ^[GS]^
'Last modified: 31/07/2012
'*************************************************

On Error Resume Next
    
    If lastStat = Estado And Force = False Then
        'Debug.Print "<Cancel" & Estado
        Exit Sub
    Else
        'Debug.Print ">Redraw" & Estado
    End If
    
    Dim szEsquinasX As Integer
    Dim szLineaHY   As Integer
    Dim szLineaVX   As Integer
    Dim szFondoY    As Integer
    
    Dim rI          As Integer
    Dim rY          As Integer
    
    lastStat = Estado
    If Estado = 0 Then
        bButton.ForeColor = ForeC
        bButton.BackColor = RGB(32, 32, 32)
        szFondoY = 0
        szLineaHY = 0
        szLineaVX = 0
        szEsquinasX = 0
        MouseO.Enabled = False
    ElseIf Estado = 1 Then
        bButton.ForeColor = ForeCo
        bButton.BackColor = RGB(64, 64, 64)
        szFondoY = szFondoH
        szLineaHY = szLineaHH * 2
        szLineaVX = szLineaVW * 2
        szEsquinasX = szEsquinasW * 2
        If isFocus = False Then MouseO.Enabled = True
    ElseIf Estado = 2 Then
        bButton.ForeColor = ForeCo
        bButton.BackColor = RGB(16, 16, 16)
        szFondoY = szFondoH * 2
        szLineaHY = szLineaHH * 4
        szLineaVX = szLineaVW * 4
        szEsquinasX = szEsquinasW * 4
        If isFocus = False Then MouseO.Enabled = True
    End If

    bButton.Cls
    ' fondo!
    If iFondoLoaded Then
        For rI = 0 To bButton.ScaleWidth Step szFondoW
            For rY = 0 To bButton.ScaleHeight Step szFondoH
                bButton.PaintPicture iFondo, rI, rY, szFondoW, szFondoH, 0, szFondoY, szFondoW, szFondoH
           Next
        Next
    End If
    ' lineas
    If iLineaHLoaded Then
        For rI = szEsquinasW To bButton.ScaleWidth - szEsquinasW Step szLineaHW ' arriba
            bButton.PaintPicture iLineaH, rI, 0, szLineaHW, szLineaHH, 0, szLineaHY, szLineaHW, szLineaHH
        Next
        For rI = szEsquinasW To bButton.ScaleWidth - szEsquinasW Step szLineaHW ' abajo
            bButton.PaintPicture iLineaH, rI, bButton.ScaleHeight - szLineaHH, szLineaHW, szLineaHH, 0, szLineaHY + szLineaHH, szLineaHW, szLineaHH
        Next
    End If
    If iLineaVLoaded Then
        For rI = szEsquinasH To bButton.ScaleHeight - szEsquinasH Step szLineaVH ' derecha
            bButton.PaintPicture iLineaV, bButton.ScaleWidth - szLineaVW, rI, szLineaVW, szLineaVH, szLineaVX + szLineaVW, 0, szLineaVW, szLineaVH
        Next
        For rI = szEsquinasH To bButton.ScaleHeight - szEsquinasH Step szLineaVH ' izquierda
            bButton.PaintPicture iLineaV, 0, rI, szLineaVW, szLineaVH, szLineaVX, 0, szLineaVW, szLineaVH
        Next
    End If
    ' esquinas
    If iEsquinaLoaded Then
        bButton.PaintPicture iEsquina, 0, 0, szEsquinasW, szEsquinasH, szEsquinasX, 0, szEsquinasW, szEsquinasH ' arriba izq
        bButton.PaintPicture iEsquina, 0, bButton.ScaleHeight - szEsquinasH, szEsquinasW, szEsquinasH, szEsquinasX, szEsquinasH, szEsquinasW, szEsquinasH ' abajo izq
        bButton.PaintPicture iEsquina, bButton.ScaleWidth - szEsquinasW, bButton.ScaleHeight - szEsquinasH, szEsquinasW, szEsquinasH, szEsquinasX + szEsquinasW, szEsquinasH, szEsquinasW, szEsquinasH ' abajo der
        bButton.PaintPicture iEsquina, bButton.ScaleWidth - szEsquinasW, 0, szEsquinasW, szEsquinasH, szEsquinasX + szEsquinasW, 0, szEsquinasW, szEsquinasH ' arriba der
    End If
    ' dibujamos el texto!
    If Force = True Then Call UpdateCaption(False)
    Call DrawCaption

End Sub


Private Sub DrawCaption()
'*************************************************
'Author: ^[GS]^
'Last modified: 27/07/2012
'*************************************************

On Error Resume Next
    
    With bButton
        Dim TempC As Long
        TempC = .ForeColor
        If (isEnabled = False) Then
            .ForeColor = RGB(128, 128, 128)
        Else
            ' Efectos
            Dim tempR As RECT
            CopyRect tempR, rc
            OffsetRect tempR, -1, -1
            If iFondoLoaded = True Then ' solo si hay fondo!
                .ForeColor = RGB(40, 10, 10)
                Call DrawText(.hdc, CaptionButton, Len(CaptionButton), tempR, DT_CENTER Or DT_WORDBREAK)
            End If
            OffsetRect tempR, 2, 2
            .ForeColor = RGB(90, 90, 90)
            Call DrawText(.hdc, CaptionButton, Len(CaptionButton), tempR, DT_CENTER Or DT_WORDBREAK)
            .ForeColor = TempC
        End If
        ' Texto
        Call DrawText(.hdc, CaptionButton, Len(CaptionButton), rc, DT_CENTER Or DT_WORDBREAK)
    End With
End Sub



Private Sub UpdateCaption(Optional bRedraw As Boolean = True)
'*************************************************
'Author: ^[GS]^
'Last modified: 18/07/2012
'*************************************************

On Error Resume Next
    
    bButton.Font = UserControl.Font
    bButton.FontBold = UserControl.FontBold
    bButton.FontItalic = UserControl.FontItalic
    bButton.FontUnderline = UserControl.FontUnderline
    bButton.FontSize = UserControl.FontSize
    bButton.FontName = UserControl.FontName
    bButton.ForeColor = ForeC
    bButton.FontSize = UserControl.FontSize
    
    Dim FixTop As Integer
    FixTop = 0
    If (bButton.TextWidth(CaptionButton) < bButton.ScaleWidth) Then
        FixTop = (bButton.TextHeight(CaptionButton) / 15) / 2.3
    End If
    With rc
        .Top = ((bButton.ScaleHeight / 15) / 2) - ((bButton.TextHeight(CaptionButton) / 15)) + FixTop
        .Left = 0
        .Bottom = bButton.ScaleHeight / 15
        .Right = bButton.ScaleWidth / 15
    End With
    
    If bRedraw Then
        Call Redraw(lastStat, True)
    End If
    
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    lastButton = 1
    Call UserControl_Click
    'Debug.Print "UserControl_AccessKeyPress2"
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next

    Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_Initialize()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    Call ReloadTextures
    isEnabled = True
    'Debug.Print "UserControl_Initialize"
    
End Sub

Private Sub UserControl_InitProperties()
'*************************************************
'Author: ^[GS]^
'Last modified: 32/07/2012
'*************************************************

On Error Resume Next
    
    lastStat = 0
    ForeC = RGB(178, 155, 111)
    ForeCo = vbWhite
    isEnabled = True
    
    Set iESQ.Picture = Nothing
    Set iFON.Picture = Nothing
    Set iHOR.Picture = Nothing
    Set iVER.Picture = Nothing
    
    iEsquinaLoaded = False
    iFondoLoaded = False
    iLineaVLoaded = False
    iLineaHLoaded = False
    bButton.BackColor = RGB(32, 32, 32)
    
    CaptionButton = Ambient.DisplayName
    Set UserControl.Font = Ambient.Font
    'Debug.Print "UserControl_InitProperties"

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
'*************************************************
'Author: ^[GS]^
'Last modified: 03/08/2012
'*************************************************

On Error Resume Next
    
    Call Redraw(0, True)
    isOver = False
End Sub

Private Sub UserControl_Show()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    Call Redraw(0)
    'Debug.Print "UserControl_Show0"
    isOver = False
End Sub

Private Sub UserControl_Resize()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    bButton.Width = UserControl.Width
    bButton.Height = UserControl.Height
    bButton.ScaleWidth = bButton.Width
    bButton.ScaleHeight = bButton.Height
    Call Redraw(0, True)
    'Debug.Print "UserControl_Resize0"
    isOver = False
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next

    With PropBag
        CaptionButton = .ReadProperty("TX", "")
        isEnabled = .ReadProperty("ENAB", True)
        ForeC = .ReadProperty("FCOL", vbWhite)
        ForeCo = .ReadProperty("OCOL", vbWhite)
        iESQ.Picture = .ReadProperty("PICE", iESQ.Picture)
        iFON.Picture = .ReadProperty("PICF", iFON.Picture)
        iHOR.Picture = .ReadProperty("PICH", iHOR.Picture)
        iVER.Picture = .ReadProperty("PICV", iVER.Picture)
        Set UserControl.Font = .ReadProperty("FONT", UserControl.Font)
    End With

    UserControl.Enabled = isEnabled
    Call ReloadTextures
    Call UpdateCaption
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    With PropBag
        Call .WriteProperty("TX", CaptionButton)
        Call .WriteProperty("ENAB", isEnabled)
        Call .WriteProperty("FCOL", ForeC)
        Call .WriteProperty("OCOL", ForeCo)
        Call .WriteProperty("PICE", iESQ.Picture)
        Call .WriteProperty("PICF", iFON.Picture)
        Call .WriteProperty("PICH", iHOR.Picture)
        Call .WriteProperty("PICV", iVER.Picture)
        Call .WriteProperty("FONT", UserControl.Font)
    End With
End Sub

Public Property Get Enabled() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    Enabled = isEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    isEnabled = NewValue
    isOver = False
    Call Redraw(0, True)
    UserControl.Enabled = isEnabled
    PropertyChanged "ENAB"
End Property

Public Property Get Caption() As String
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    Caption = CaptionButton
End Property

Public Property Let Caption(ByVal NewValue As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    CaptionButton = NewValue
    Call UpdateCaption
    Call Redraw(0, True)
    isOver = False
    PropertyChanged "TX"
End Property

Public Property Get ForeColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    ForeColor = ForeC
End Property

Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    ForeC = theCol
    Call UpdateCaption
    Call Redraw(0)
    PropertyChanged "FCOL"
End Property

Public Property Get ForeColorOver() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    ForeColorOver = ForeCo
End Property

Public Property Let ForeColorOver(ByVal theCol As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    ForeCo = theCol
    Call UpdateCaption
    Call Redraw(0)
    PropertyChanged "OCOL"
End Property

Public Property Get Font() As Font
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByRef newFont As Font)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    Set UserControl.Font = newFont
    Call UpdateCaption
    Call Redraw(0)
    isOver = False
    PropertyChanged "FONT"
End Property

Public Property Get FontBold() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    UserControl.FontBold = NewValue
    Call UpdateCaption
    Call Redraw(0)
End Property

Public Property Get FontItalic() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    UserControl.FontItalic = NewValue
    Call UpdateCaption
    Call Redraw(0)
    isOver = False
End Property

Public Property Get FontUnderline() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    UserControl.FontUnderline = NewValue
    Call UpdateCaption
    Call Redraw(0)
    isOver = False
End Property

Public Property Get FontSize() As Integer
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal NewValue As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    UserControl.FontSize = NewValue
    Call UpdateCaption
    Call Redraw(0)
    isOver = False
End Property

Public Property Get FontName() As String
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal NewValue As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    UserControl.FontName = NewValue
    Call UpdateCaption
    Call Redraw(0)
    isOver = False
End Property


Public Property Get PictureEsquina() As StdPicture
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    Set PictureEsquina = iESQ.Picture

End Property

Public Property Set PictureEsquina(ByVal newPic As StdPicture)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next
    
    iESQ.Picture = newPic
    Call ReloadTextures
    Call Redraw(0, True)
    PropertyChanged "PICE"

End Property

Public Property Get PictureFondo() As StdPicture
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next

    Set PictureFondo = iFON.Picture

End Property

Public Property Set PictureFondo(ByVal newPic As StdPicture)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next

    iFON.Picture = newPic
    Call ReloadTextures
    Call Redraw(0, True)
    PropertyChanged "PICF"

End Property


Public Property Get PictureHorizontal() As StdPicture
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next

    Set PictureHorizontal = iHOR.Picture

End Property

Public Property Set PictureHorizontal(ByVal newPic As StdPicture)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next

    iHOR.Picture = newPic
    Call ReloadTextures
    Call Redraw(0, True)
    PropertyChanged "PICH"

End Property

Public Property Get PictureVertical() As StdPicture
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next

    Set PictureVertical = iVER.Picture

End Property

Public Property Set PictureVertical(ByVal newPic As StdPicture)
'*************************************************
'Author: ^[GS]^
'Last modified: 15/07/2012
'*************************************************

On Error Resume Next

    iVER.Picture = newPic
    Call ReloadTextures
    Call Redraw(0, True)
    PropertyChanged "PICV"

End Property

