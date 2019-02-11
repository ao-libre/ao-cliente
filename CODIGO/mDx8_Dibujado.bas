Attribute VB_Name = "mDx8_Dibujado"
Option Explicit

Public MinEleccion As Integer
Public MaxEleccion As Integer
Public Actual      As Integer

Private Declare Function BitBlt _
                Lib "gdi32" (ByVal hDestDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal dwRop As Long) As Long
Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateStreamOnHGlobal _
                Lib "ole32" (ByVal hGlobal As Long, _
                             ByVal fDeleteOnRelease As Long, _
                             ppstm As Any) As Long
Private Declare Function GlobalAlloc _
                Lib "kernel32" (ByVal uFlags As Long, _
                                ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleLoadPicture _
                Lib "olepro32" (pStream As Any, _
                                ByVal lSize As Long, _
                                ByVal fRunmode As Long, _
                                riid As Any, _
                                ppvObj As Any) As Long
Private Declare Sub CopyMemory _
                Lib "kernel32.dll" _
                Alias "RtlMoveMemory" (ByRef destination As Any, _
                                       ByRef source As Any, _
                                       ByVal length As Long)

Private Declare Function SetBitmapBits _
                Lib "gdi32" (ByVal hBitmap As Long, _
                             ByVal dwCount As Long, _
                             lpBits As Any) As Long

Public Sub ArrayToPicturePNG(ByRef byteArray() As Byte, ByRef imgDest As IPicture) ' GSZAO
    
    On Error GoTo ArrayToPicturePNG_Err
    
    Call SetBitmapBits(imgDest.handle, UBound(byteArray), byteArray(0))

    
    Exit Sub

ArrayToPicturePNG_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Dibujado" & "->" & "ArrayToPicturePNG"
    End If
Resume Next
    
End Sub

Public Function ArrayToPicture(inArray() As Byte, _
                               offset As Long, _
                               Size As Long) As IPicture
    
    On Error GoTo ArrayToPicture_Err
    
    
    Dim o_hMem        As Long
    Dim o_lpMem       As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream      As IUnknown
    
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)

    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)

        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(offset), Size
            Call GlobalUnlock(o_hMem)

            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), ArrayToPicture)

            End If

        End If

    End If

    
    Exit Function

ArrayToPicture_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Dibujado" & "->" & "ArrayToPicture"
    End If
Resume Next
    
End Function

Sub DrawGrhtoHdc(ByVal desthDC As Long, _
                 ByVal grh_index As Integer, _
                 ByRef SourceRect As RECT, _
                 ByRef destRect As RECT)

    On Error Resume Next
    
    Dim file_path  As String
    Dim src_x      As Integer
    Dim src_y      As Integer
    Dim src_width  As Integer
    Dim src_height As Integer
    Dim hdcsrc     As Long
    Dim MaskDC     As Long
    Dim PrevObj    As Long
    Dim PrevObj2   As Long
    Dim screen_x   As Integer
    Dim screen_y   As Integer
    
    screen_x = destRect.Left
    screen_y = destRect.Top
    
    If grh_index <= 0 Then Exit Sub

    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)

    End If
    
    Dim data()  As Byte
    Dim bmpData As StdPicture
    
    'get Picture
    If Get_Image(DirGraficos, CStr(GrhData(grh_index).FileNum), data, True) Then  ' GSZAO
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
