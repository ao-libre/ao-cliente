VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Por Javier Podavini (El Yind)
'09/07/2009
' Ultima actualizacion: 01/10/2020 WyroX

Private ListPic As PictureBox
Private lista() As String
Private p_ListIndex As Integer
Private p_ListCount As Integer
Private p_Visible As Boolean
Private p_ForeColor As Long
Private p_Pixel_Alto As Integer
Private p_Scroll As Integer
Private p_OldY As Integer
Private p_Padding As Integer
Private p_ScrollWidth As Integer

Private Sub Class_Initialize()
    p_Visible = False
    Clear
End Sub

Public Sub Initialize(ByRef LPic As PictureBox, ByVal ForeColor As Long, ByVal Padding As Integer, ByVal ScrollBarWidth As Integer)
    Set ListPic = LPic
    
    p_ForeColor = ForeColor
    p_Pixel_Alto = 13
    p_Padding = Padding
    p_ScrollWidth = ScrollBarWidth
    
    ListPic.ForeColor = p_ForeColor
    ListPic.Cls
End Sub

Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_UserMemId = 0
    List = lista(Index)
End Property
Public Property Let List(ByVal Index As Integer, ByVal Valor As String)
    lista(Index) = CutStringLength(Valor)
    Dibujar
End Property

Public Property Get Visible() As Integer
    Visible = p_Visible
End Property
Public Property Let Visible(ByVal Valor As Integer)
    p_Visible = Valor
    ListPic.Visible = p_Visible
End Property

Public Property Get ListIndex() As Integer
    ListIndex = p_ListIndex
End Property
Public Property Let ListIndex(ByVal Valor As Integer)
    p_ListIndex = Valor
    If p_ListIndex >= p_ListCount Then p_ListIndex = p_ListCount - 1
    
    ' Si seleccionamos un item fuera de vista, scrolleamos hasta verlo
    If p_ListIndex >= 0 Then
        If p_ListIndex < p_Scroll Then
            ' Scroll para arriba
            p_Scroll = p_ListIndex
        
        ElseIf p_ListIndex >= p_Scroll + ListPic.Height \ p_Pixel_Alto Then
            ' Scroll para abajo
            p_Scroll = p_ListIndex - ListPic.Height \ p_Pixel_Alto + 1
        End If
    End If
    
    Dibujar
End Property

Public Property Get OldY() As Integer
    OldY = p_OldY
End Property
Public Property Let OldY(ByVal Valor As Integer)
    p_OldY = Valor
End Property

Public Property Get Pixel_Alto() As Integer
    Pixel_Alto = p_Pixel_Alto
End Property
Public Property Let Pixel_Alto(ByVal Valor As Integer)
    p_Pixel_Alto = Valor
End Property

Public Property Get Scroll() As Integer
    Scroll = p_Scroll
End Property
Public Property Let Scroll(ByVal Valor As Integer)
    p_Scroll = Valor
    Dibujar
End Property

Public Property Get ForeColor() As Long
    ForeColor = p_ForeColor
End Property
Public Property Let ForeColor(ByVal Valor As Long)
    p_ForeColor = Valor
    Dibujar
End Property

Public Property Get ListCount() As Integer
    ListCount = p_ListCount
End Property

Public Sub AddItem(item As String)
    ReDim Preserve lista(p_ListCount)
    lista(p_ListCount) = CutStringLength(item)
    p_ListCount = p_ListCount + 1
    If p_ListCount = 1 Then p_ListIndex = 0
    Dibujar
End Sub

Public Sub Clear()
    ReDim lista(0)
    p_ListCount = 0
    p_ListIndex = -1
    p_Scroll = 0
    If Not ListPic Is Nothing Then Dibujar
End Sub

Private Sub Dibujar()

    Dim i As Integer
    Dim PosY As Integer
    
    ListPic.Cls
    
    For i = p_Scroll To p_Scroll + ListPic.Height \ p_Pixel_Alto - 1
        If i < p_ListCount Then
            PosY = (i - p_Scroll) * p_Pixel_Alto + p_Padding
    
            If i = p_ListIndex Then
                ListPic.ForeColor = RGB(0, 120, 255)
                ListPic.Line (1, PosY)-(ListPic.Width - p_ScrollWidth - p_Padding, PosY + p_Pixel_Alto - 1), , B
                ListPic.ForeColor = RGB(180, 180, 180)
            Else
                ListPic.ForeColor = p_ForeColor
            End If
            ListPic.CurrentX = p_Padding
            ListPic.CurrentY = PosY
           
            ListPic.Print lista(i)
        End If
    Next i

End Sub

Private Function CutStringLength(Text As String) As String
    Dim LastEsp As Integer
    Dim i As Integer
    Dim MaxWidth As Integer
    
    MaxWidth = ListPic.Width - p_ScrollWidth - p_Padding * 2

    If ListPic.TextWidth(Text) > MaxWidth Then
        LastEsp = 0
        
        For i = 1 To Len(Text)
            If mid(Text, i, 1) = " " Then LastEsp = i
            If ListPic.TextWidth(Left$(Text, i)) > MaxWidth Then Exit For
        Next i
        
        If LastEsp = 0 Then LastEsp = i - 1

        CutStringLength = Left$(Text, LastEsp)
        
    Else
        CutStringLength = Text
    End If
End Function
