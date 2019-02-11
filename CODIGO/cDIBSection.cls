VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (lpvDest As Any, _
                                       lpvSource As Any, _
                                       ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND

    cElements As Long
    lLbound As Long

End Type

Private Type SAFEARRAY2D

    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND

End Type

Private Declare Function VarPtrArray _
                Lib "msvbvm50.dll" _
                Alias "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD

    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte

End Type

Private Type BITMAPINFOHEADER '40 bytes

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long

End Type

Private Type BITMAPINFO

    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD

End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection _
                Lib "gdi32" (ByVal hdc As Long, _
                             pBitmapInfo As BITMAPINFO, _
                             ByVal un As Long, _
                             lplpVoid As Long, _
                             ByVal handle As Long, _
                             ByVal dw As Long) As Long
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
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hdc As Long) As Long
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Type BITMAP

    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long

End Type

Private Declare Function GetObjectAPI _
                Lib "gdi32" _
                Alias "GetObjectA" (ByVal hObject As Long, _
                                    ByVal nCount As Long, _
                                    lpObject As Any) As Long

' Handle to the current DIBSection:
Private m_hDIb    As Long
' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long
' Handle to the Device context holding the DIBSection:
Private m_hDC     As Long
' Address of memory pointing to the DIBSection's bits:
Private m_lPtr    As Long
' Type containing the Bitmap information:
Private m_tBI     As BITMAPINFO

Public Function CreateDIB(ByVal lhDC As Long, _
                          ByVal lWidth As Long, _
                          ByVal lHeight As Long, _
                          ByRef hDib As Long) As Boolean
    
    On Error GoTo CreateDIB_Err
    

    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight

    End With

    hDib = CreateDIBSection(lhDC, m_tBI, DIB_RGB_COLORS, m_lPtr, 0, 0)
    CreateDIB = (hDib <> 0)

    
    Exit Function

CreateDIB_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "CreateDIB"
    End If
Resume Next
    
End Function

Public Function CreateFromPicture(ByRef picThis As StdPicture)
    
    On Error GoTo CreateFromPicture_Err
    
    Dim lhDC        As Long
    Dim lhDCDesktop As Long
    Dim lhBmpOld    As Long
    Dim tBMP        As BITMAP
    Dim lhWnd       As Long
    
    GetObjectAPI picThis.handle, Len(tBMP), tBMP

    If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
        lhWnd = GetDesktopWindow()
        lhDCDesktop = GetDC(lhWnd)

        If (lhDCDesktop <> 0) Then
            lhDC = CreateCompatibleDC(lhDCDesktop)
            ReleaseDC lhWnd, lhDCDesktop

            If (lhDC <> 0) Then
                lhBmpOld = SelectObject(lhDC, picThis.handle)
                LoadPictureBlt lhDC
                SelectObject lhDC, lhBmpOld
                DeleteDC lhDC

            End If

        End If

    End If

    
    Exit Function

CreateFromPicture_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "CreateFromPicture"
    End If
Resume Next
    
End Function

Public Function Create(ByVal lWidth As Long, ByVal lHeight As Long) As Boolean
    
    On Error GoTo Create_Err
    
    ClearUp
    m_hDC = CreateCompatibleDC(0)

    If (m_hDC <> 0) Then
        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIb)
            Create = True
        Else
            DeleteDC m_hDC
            m_hDC = 0

        End If

    End If

    
    Exit Function

Create_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "Create"
    End If
Resume Next
    
End Function

Public Property Get BytesPerScanLine() As Long
    ' Scans must align on dword boundaries:
    
    On Error GoTo BytesPerScanLine_Err
    
    BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC

    
    Exit Property

BytesPerScanLine_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "BytesPerScanLine"
    End If
Resume Next
    
End Property

Public Property Get Width() As Long
    
    On Error GoTo Width_Err
    
    Width = m_tBI.bmiHeader.biWidth

    
    Exit Property

Width_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "Width"
    End If
Resume Next
    
End Property

Public Property Get Height() As Long
    
    On Error GoTo Height_Err
    
    Height = m_tBI.bmiHeader.biHeight

    
    Exit Property

Height_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "Height"
    End If
Resume Next
    
End Property

Public Sub LoadPictureBlt(ByVal lhDC As Long, _
                          Optional ByVal lSrcLeft As Long = 0, _
                          Optional ByVal lSrcTop As Long = 0, _
                          Optional ByVal lSrcWidth As Long = -1, _
                          Optional ByVal lSrcHeight As Long = -1, _
                          Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
    
    On Error GoTo LoadPictureBlt_Err
    

    If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lhDC, lSrcLeft, lSrcTop, eRop

    
    Exit Sub

LoadPictureBlt_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "LoadPictureBlt"
    End If
Resume Next
    
End Sub

Public Sub PaintPicture(ByVal lhDC As Long, _
                        Optional ByVal lDestLeft As Long = 0, _
                        Optional ByVal lDestTop As Long = 0, _
                        Optional ByVal lDestWidth As Long = -1, _
                        Optional ByVal lDestHeight As Long = -1, _
                        Optional ByVal lSrcLeft As Long = 0, _
                        Optional ByVal lSrcTop As Long = 0, _
                        Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
    
    On Error GoTo PaintPicture_Err
    

    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
    BitBlt lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop

    
    Exit Sub

PaintPicture_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "PaintPicture"
    End If
Resume Next
    
End Sub

Public Property Get DIBSectionBitsPtr() As Long
    
    On Error GoTo DIBSectionBitsPtr_Err
    
    DIBSectionBitsPtr = m_lPtr

    
    Exit Property

DIBSectionBitsPtr_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "DIBSectionBitsPtr"
    End If
Resume Next
    
End Property

Public Sub ClearUp()
    
    On Error GoTo ClearUp_Err
    

    If (m_hDC <> 0) Then
        If (m_hDIb <> 0) Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDIb

        End If

        DeleteDC m_hDC

    End If

    m_hDC = 0: m_hDIb = 0: m_hBmpOld = 0: m_lPtr = 0

    
    Exit Sub

ClearUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "ClearUp"
    End If
Resume Next
    
End Sub

Private Function ResampleDib(ByRef cDibTo As cDIBSection) As Boolean
    
    On Error GoTo ResampleDib_Err
    
    Dim bDibFrom() As Byte
    Dim bDibTo()   As Byte

    Dim tSAFrom    As SAFEARRAY2D
    Dim tSATo      As SAFEARRAY2D

    ' Get the bits in the from DIB section:
    With tSAFrom
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr

    End With

    CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4

    ' Get the bits in the to DIB section:
    With tSATo
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibTo.BytesPerScanLine()
        .pvData = cDibTo.DIBSectionBitsPtr

    End With

    CopyMemory ByVal VarPtrArray(bDibTo()), VarPtr(tSATo), 4

    Dim xScale As Single
    Dim yScale As Single

    Dim X      As Long, Y As Long, xEnd As Long, xOut As Long

    Dim fX     As Single, fY As Single
    Dim ifY    As Long, ifX As Long
    Dim dX     As Single, dy As Single
    Dim r      As Long, r1 As Single, r2 As Single, r3 As Single, r4 As Single
    Dim g      As Long, g1 As Single, g2 As Single, g3 As Single, g4 As Single
    Dim b      As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single
    Dim ir1    As Long, ig1 As Long, ib1 As Long
    Dim ir2    As Long, ig2 As Long, ib2 As Long

    xScale = (Width - 1) / cDibTo.Width
    yScale = (Height - 1) / cDibTo.Height
    
    xEnd = cDibTo.Width - 1
        
    For Y = 0 To cDibTo.Height - 1
        
        fY = Y * yScale
        ifY = Int(fY)
        dy = fY - ifY
        
        For X = 0 To xEnd
            fX = X * xScale
            ifX = Int(fX)
            dX = fX - ifX
            
            ifX = ifX * 3
            ' Interpolate using the four nearest pixels in the source
            b1 = bDibFrom(ifX, ifY): g1 = bDibFrom(ifX + 1, ifY): r1 = bDibFrom(ifX + 2, ifY)
            b2 = bDibFrom(ifX + 3, ifY): g2 = bDibFrom(ifX + 4, ifY): r2 = bDibFrom(ifX + 5, ifY)
            b3 = bDibFrom(ifX, ifY + 1): g3 = bDibFrom(ifX + 1, ifY + 1): r3 = bDibFrom(ifX + 2, ifY + 1)
            b4 = bDibFrom(ifX + 3, ifY + 1): g4 = bDibFrom(ifX + 4, ifY + 1): r4 = bDibFrom(ifX + 5, ifY + 1)
            
            ' Interplate in x direction:
            ir1 = r1 * (1 - dy) + r3 * dy: ig1 = g1 * (1 - dy) + g3 * dy: ib1 = b1 * (1 - dy) + b3 * dy
            ir2 = r2 * (1 - dy) + r4 * dy: ig2 = g2 * (1 - dy) + g4 * dy: ib2 = b2 * (1 - dy) + b4 * dy
            ' Interpolate in y:
            r = ir1 * (1 - dX) + ir2 * dX: g = ig1 * (1 - dX) + ig2 * dX: b = ib1 * (1 - dX) + ib2 * dX
            
            ' Set output:
            If (r < 0) Then r = 0
            If (r > 255) Then r = 255
            If (g < 0) Then g = 0
            If (g > 255) Then g = 255
            If (b < 0) Then b = 0
            If (b > 255) Then
                b = 255

            End If

            xOut = X * 3
            bDibTo(xOut, Y) = b
            bDibTo(xOut + 1, Y) = g
            bDibTo(xOut + 2, Y) = r
            
        Next X
        
    Next Y

    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
    CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4

    
    Exit Function

ResampleDib_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "ResampleDib"
    End If
Resume Next
    
End Function

Private Sub Class_Terminate()
    
    On Error GoTo Class_Terminate_Err
    
    ClearUp

    
    Exit Sub

Class_Terminate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cDIBSection" & "->" & "Class_Terminate"
    End If
Resume Next
    
End Sub

