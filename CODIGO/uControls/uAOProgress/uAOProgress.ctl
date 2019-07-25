VERSION 5.00
Begin VB.UserControl uAOProgress 
   BackStyle       =   0  'Transparent
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   ScaleHeight     =   945
   ScaleWidth      =   2295
   ToolboxBitmap   =   "uAOProgress.ctx":0000
   Begin VB.Timer tTimer 
      Interval        =   10
      Left            =   1680
      Top             =   360
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Shape shpStat 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   1365
   End
   Begin VB.Shape shpBack 
      BorderColor     =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2085
   End
End
Attribute VB_Name = "uAOProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%                                                     %
'%                   AO PROGRESS v1.4                  %
'%               Copyright © 2013 by ^[GS]^            %
'%                    www.GS-ZONE.org                  %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%  Este control permite realizar barras de            %
'%  progreso facilmente.                               %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%  Changelog:                                         %
'%   25/04/2013 - Se agrego la posibilidad de usar     %
'%                un color de fondo solido. (^[GS]^)   %
'%   03/09/2012 - Mejora de rendimiento.               %
'%                Se agrego el % al mantener el        %
'%                sobre el valor. (^[GS]^)             %
'%   25/08/2012 - Se finalizo una primera versión,     %
'%                sencilla, son Shapes y animación     %
'%                de cambio de valor. (^[GS]^)         %
'%   23/07/2012 - Se inicio el proyecto. (^[GS]^)      %
'%                                                     %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Option Explicit

Private iMax As Long
Private iMin As Long
Private iValue As Long
Private iNewValue As Long
Private bAnimate As Boolean
Private bEnabled As Boolean
Private bUseBackground As Boolean
Private lForeColor As Long
Private lBackColor As Long
Private lBorderColor As Long
Private lBackgroundColor As Long
Private fTextFont As Font

Private MouseOverText As String
Private bAnimating As Boolean

Private Sub DrawStat()
'*************************************************
'Author: ^[GS]^
'Last modified: 03/09/2012
'*************************************************

On Error Resume Next
    
    If bEnabled = False Then Exit Sub
    
    If LenB(MouseOverText) <> 0 Then MouseOverText = vbNullString
    If bAnimate = False Then
        iNewValue = iValue
        lblStat.Caption = iNewValue & "/" & iMax
        shpStat.Width = (((iNewValue / 100) / (iMax / 100)) * UserControl.Width)
    Else
        If iNewValue = iValue Then
            tTimer.Enabled = False
        Else
            tTimer.Enabled = True
        End If
        Dim lDif As Long
        lDif = Abs(iValue - iNewValue)
        If iNewValue < iValue Then
            iNewValue = iNewValue + 1
            Select Case lDif
                Case Is > 500
                    iNewValue = iNewValue + (lDif / 8)
                Case Is > 100
                    iNewValue = iNewValue + (lDif / 14)
                Case Is > 10
                    iNewValue = iNewValue + (lDif / 18)
            End Select
            If iNewValue > iValue Then iNewValue = iValue
            bAnimating = True
        ElseIf iNewValue > iValue Then
            iNewValue = iNewValue - 1
            Select Case lDif
                Case Is > 500
                    iNewValue = iNewValue - (lDif / 8)
                Case Is > 100
                    iNewValue = iNewValue - (lDif / 14)
                Case Is > 10
                    iNewValue = iNewValue - (lDif / 18)
            End Select
            If iNewValue < iValue Then iNewValue = iValue
            bAnimating = True
        Else
            iNewValue = iValue
            bAnimating = False
        End If
        If lDif > (iMax / 10) Or iValue < (iMax / 10) Then
            tTimer.Interval = 1
        Else
            tTimer.Interval = 30
        End If
        lblStat.Caption = iNewValue & "/" & iMax
        shpStat.Width = (((iNewValue / 100) / (iMax / 100)) * UserControl.Width)
    End If
    
End Sub




Private Sub lblStat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 03/09/2012
'*************************************************

On Error Resume Next

    If LenB(MouseOverText) = 0 Then
        MouseOverText = Round(CDbl(iNewValue) * CDbl(100) / CDbl(iMax), 2) & "%"
    End If
    
    lblStat.Caption = MouseOverText
   
End Sub

Private Sub lblStat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 03/09/2012
'*************************************************

On Error Resume Next

    lblStat.Caption = iNewValue & "/" & iMax
    
End Sub

Private Sub tTimer_Timer()
'*************************************************
'Author: ^[GS]^
'Last modified: 03/09/2012
'*************************************************

On Error Resume Next
    
    If bEnabled = False Then
        tTimer.Enabled = False
        Exit Sub
    End If
    If bAnimating = True Then
        Call DrawStat
    End If
    
End Sub

Private Sub ResizeLabel()
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.Left = 0
    lblStat.Width = UserControl.Width
    lblStat.Top = (UserControl.Height / 2) - ((lblStat.Height / 2))
    Call DrawStat
    
End Sub

Private Sub UserControl_InitProperties()
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next

    iMax = 100
    iMin = 1
    iValue = 1
    bEnabled = True
    bAnimate = True
    bUseBackground = False
    lBackgroundColor = RGB(0, 0, 0)
    lForeColor = RGB(255, 255, 255)
    lBackColor = RGB(100, 100, 100)
    lBorderColor = RGB(200, 200, 200)
    
End Sub

Private Sub UserControl_Resize()
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    shpStat.Left = 0
    shpStat.Height = UserControl.Height
    shpBack.Height = UserControl.Height
    shpBack.Width = UserControl.Width
    Call ResizeLabel
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    With PropBag
        iMax = .ReadProperty("Max", 100)
        iMin = .ReadProperty("Min", 0)
        iValue = .ReadProperty("Value", 50)
        bEnabled = .ReadProperty("Enabled", True)
        bAnimate = .ReadProperty("Animate", True)
        bUseBackground = .ReadProperty("UseBackground", True)
        lBackgroundColor = .ReadProperty("BackgroundColor", RGB(0, 0, 0))
        lForeColor = .ReadProperty("ForeColor", RGB(255, 255, 255))
        lBackColor = .ReadProperty("BackColor", RGB(100, 100, 100))
        lBorderColor = .ReadProperty("BorderColor", RGB(200, 200, 200))
        Set lblStat.Font = .ReadProperty("FONT", lblStat.Font)
    End With
    
    lblStat.ForeColor = lForeColor
    shpStat.FillColor = lBackColor
    shpStat.BorderColor = lBorderColor
    shpBack.BorderColor = lBorderColor
    shpBack.BackColor = lBackgroundColor
    shpBack.Visible = bUseBackground
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    With PropBag
        .WriteProperty "Max", iMax, 100
        .WriteProperty "Min", iMin, 0
        .WriteProperty "Value", iValue, 50
        .WriteProperty "Enabled", bEnabled, True
        .WriteProperty "Animate", bAnimate, True
        .WriteProperty "UseBackground", bUseBackground, True
        .WriteProperty "BackgroundColor", lBackgroundColor, RGB(0, 0, 0)
        .WriteProperty "ForeColor", lForeColor, RGB(255, 255, 255)
        .WriteProperty "BackColor", lBackColor, RGB(100, 100, 100)
        .WriteProperty "BorderColor", lBorderColor, RGB(200, 200, 200)
        Call .WriteProperty("FONT", lblStat.Font)
    End With
    
End Sub

Public Property Get Enabled() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Enabled = bEnabled
    
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    bEnabled = NewValue
    PropertyChanged "Enabled"
    
    UserControl.Enabled = False
    
End Property

Public Property Get Animado() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Animado = bAnimate
    
End Property

Public Property Let Animado(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    bAnimate = NewValue
    PropertyChanged "Animate"
    
    Call DrawStat
    
End Property

Public Property Get UseBackground() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    UseBackground = bUseBackground
    
End Property

Public Property Let UseBackground(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    bUseBackground = NewValue
    PropertyChanged "UseBackground"
    
    shpBack.Visible = bUseBackground
    
End Property

Public Property Get Font() As Font
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Set Font = lblStat.Font
    
End Property

Public Property Set Font(ByRef newFont As Font)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Set lblStat.Font = newFont

    Call ResizeLabel

    PropertyChanged "FONT"
    
End Property

Public Property Get FontBold() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    FontBold = lblStat.FontBold
    
End Property

Public Property Let FontBold(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.FontBold = NewValue
    
    Call ResizeLabel
    
End Property

Public Property Get FontItalic() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    FontItalic = lblStat.FontItalic
    
End Property

Public Property Let FontItalic(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.FontItalic = NewValue

    Call ResizeLabel
    
End Property

Public Property Get FontUnderline() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    FontUnderline = lblStat.FontUnderline
    
End Property

Public Property Let FontUnderline(ByVal NewValue As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.FontUnderline = NewValue

    Call ResizeLabel
    
End Property

Public Property Get FontSize() As Integer
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    FontSize = lblStat.FontSize
    
End Property

Public Property Let FontSize(ByVal NewValue As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.FontSize = NewValue

    Call ResizeLabel
    
End Property

Public Property Get FontName() As String
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    FontName = lblStat.FontName
    
End Property

Public Property Let FontName(ByVal NewValue As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lblStat.FontName = NewValue
    
    Call ResizeLabel
    
End Property

Public Property Get ForeColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    ForeColor = lForeColor
    
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lForeColor = NewValue
    PropertyChanged "ForeColor"
    
    lblStat.ForeColor = lForeColor
    
End Property

Public Property Get BackgroundColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    BackgroundColor = lBackgroundColor
    
End Property

Public Property Let BackgroundColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    lBackgroundColor = NewValue
    PropertyChanged "BackgroundColor"
    
    shpBack.FillColor = lBackgroundColor
    
End Property

Public Property Get BackColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    BackColor = lBackColor
    
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    lBackColor = NewValue
    PropertyChanged "BackColor"
    
    shpStat.FillColor = lBackColor
    
End Property

Public Property Get BorderColor() As OLE_COLOR
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    BorderColor = lBorderColor
    
End Property

Public Property Let BorderColor(ByVal NewValue As OLE_COLOR)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/04/2013
'*************************************************

On Error Resume Next
    
    lBorderColor = NewValue
    PropertyChanged "BorderColor"
    
    shpStat.BorderColor = lBorderColor
    shpBack.BackColor = lBorderColor
    
End Property

Public Property Let Value(ByVal NewValue As Long)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    If NewValue > iMax Then NewValue = iMax
    If NewValue < iMin Then NewValue = iMin
    iValue = NewValue
    
    PropertyChanged "Value"
    
    Call DrawStat
    
End Property

Public Property Get Value() As Long
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Value = iValue
    
End Property

Public Property Let Max(ByVal NewValue As Long)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    If NewValue < 1 Then NewValue = 1
    If NewValue <= iMin Then NewValue = iMin + 1
    iMax = NewValue
    
    If Value > iMax Then Value = iMax
    PropertyChanged "Max"
    
    Call DrawStat
    
End Property

Public Property Get Max() As Long
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Max = iMax
    
End Property

Public Property Let Min(ByVal NewValue As Long)
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    If NewValue >= iMax Then NewValue = Max - 1
    If NewValue < 0 Then NewValue = 0
    iMin = NewValue
    If Value < iMin Then Value = iMin
    
    PropertyChanged "Min"
    
    Call DrawStat
    
End Property

Public Property Get Min() As Long
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012
'*************************************************

On Error Resume Next
    
    Min = iMin
    
End Property
