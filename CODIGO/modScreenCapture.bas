Attribute VB_Name = "modScreenCapture"
Option Explicit

' ==================================================================================
' Author:      Steve McMahon
' Date:        15 March 1999
' Requires:    cDIBSection.cls
'              IJL11.DLL (Intel)
'
' An interface to Intel's IJL (Intel JPG Library) for use in VB.
'
' Modifications
'  Author: Alejandro Salvo
'  Date: 5 February 2007
'  Description: Added ScreenCapture Method and removed usless things.
'
'  Author: Juan Martín Sotuyo Dodero (Maraxus)
'  Date: 28 Febraury 2007
'  Description: Changed ScreenCapture to use the Screenshots directory.
'               Fixed a bug that caused the DC not to be allways released (added INVALID_HANDLE constant)
'
'  Author: Torres Patricio (Pato)
'  Date: 25 August 2009
'  Description: Added FullScreenCapture Function.
'
'
' Copyright.
' IJL.DLL is a copyright © Intel, which is a registered trade mark of the Intel
' Corporation.
'
'
' Note.
' Intel are not responsible for any errors in this code and should not be
' mentioned in any Help, About or support in any product using the Intel library.
'
'
'
' ==================================================================================

' IJL Declares:

Private Enum IJLERR
    '// The following "error" values indicate an "OK" condition.
    IJL_OK = 0
End Enum

Private Enum IJLIOTYPE

    ''// Read JPEG parameters (i.e., height, width, channels,
    ''// sampling, etc.) from a JPEG bit stream.
    IJL_JFILE_READPARAMS = 0&
    IJL_JBUFF_READPARAMS = 1&
    ''// Read a JPEG Interchange Format image.
    IJL_JFILE_READWHOLEIMAGE = 2&
    IJL_JBUFF_READWHOLEIMAGE = 3&
    ''// Write an entire JFIF bit stream.
    IJL_JFILE_WRITEWHOLEIMAGE = 8&
    IJL_JBUFF_WRITEWHOLEIMAGE = 9&

End Enum

Private Type JPEG_CORE_PROPERTIES_VB ' Sadly, due to a limitation in VB (UDT variable count)

    ' we can't encode the full JPEG_CORE_PROPERTIES structure
    UseJPEGPROPERTIES As Long                      '// default = 0

    '// DIB specific I/O data specifiers.
    DIBBytes As Long ';                  '// default = NULL 4
    DIBWidth As Long ';                  '// default = 0 8
    DIBHeight As Long ';                 '// default = 0 12
    DIBPadBytes As Long ';               '// default = 0 16
    DIBChannels As Long ';               '// default = 3 20
    DIBColor As Long ';                  '// default = IJL_BGR 24
    DIBSubsampling As Long  ';            '// default = IJL_NONE 28

    '// JPEG specific I/O data specifiers.
    JPGFile As Long 'LPTSTR              JPGFile;                32   '// default = NULL
    JPGBytes As Long ';                  '// default = NULL 36
    JPGSizeBytes As Long ';              '// default = 0 40
    JPGWidth As Long ';                  '// default = 0 44
    JPGHeight As Long ';                 '// default = 0 48
    JPGChannels As Long ';               '// default = 3
    JPGColor As Long           ';                  '// default = IJL_YCBCR
    JPGSubsampling As Long  ';            '// default = IJL_411
    JPGThumbWidth As Long ' ;             '// default = 0
    JPGThumbHeight As Long ';            '// default = 0

    '// JPEG conversion properties.
    cconversion_reqd As Long ';          '// default = TRUE
    upsampling_reqd As Long ';           '// default = TRUE
    jquality As Long ';                  '// default = 75.  100 is my preferred quality setting.

    '// Low-level properties - 20,000 bytes.  If the whole structure
    ' is written out then VB fails with an obscure error message
    ' "Too Many Local Variables" !
    '
    ' These all default if they are not otherwise specified so there
    ' is no trouble to just assign a sufficient buffer in memory:
    jprops(0 To 19999) As Byte

End Type

'
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hdc As Long) As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef dest As Any, _
                                       ByRef source As Any, _
                                       ByVal byteCount As Long)

'

Private Declare Function ijlInit Lib "ijl11.dll" (jcprops As Any) As Long
Private Declare Function ijlFree Lib "ijl11.dll" (jcprops As Any) As Long
Private Declare Function ijlRead _
                Lib "ijl11.dll" (jcprops As Any, _
                                 ByVal ioType As Long) As Long
Private Declare Function ijlWrite _
                Lib "ijl11.dll" (jcprops As Any, _
                                 ByVal ioType As Long) As Long

' Stuff for replacing a file when you have to Kill the original:
Private Const MAX_PATH = 260

Private Type FILETIME

    dwLowDateTime As Long
    dwHighDateTime As Long

End Type

Private Type WIN32_FIND_DATA

    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14

End Type

Private Declare Function FindFirstFile _
                Lib "kernel32" _
                Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                        lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function lopen _
                Lib "kernel32" _
                Alias "_lopen" (ByVal lpPathName As String, _
                                ByVal iReadWrite As Long) As Long
Private Declare Function lclose _
                Lib "kernel32" _
                Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function SetFileTime _
                Lib "kernel32" (ByVal hFile As Long, _
                                lpCreationTime As FILETIME, _
                                lpLastAccessTime As FILETIME, _
                                lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileAttributes _
                Lib "kernel32" _
                Alias "SetFileAttributesA" (ByVal lpFileName As String, _
                                            ByVal dwFileAttributes As Long) As Long
Private Const OF_WRITE = &H1
Private Const OF_SHARE_DENY_WRITE = &H20

Private Const INVALID_HANDLE As Long = -1

'bltbit constant
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

'Good old bitblt
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

Public Function SaveJPG(ByRef cDib As cDIBSection, _
                        ByVal sFile As String, _
                        Optional ByVal lQuality As Long = 90) As Boolean
    
    On Error GoTo SaveJPG_Err
    
    Dim tJ           As JPEG_CORE_PROPERTIES_VB
    Dim bFile()      As Byte
    Dim lPtr         As Long
    Dim lR           As Long
    Dim tFnd         As WIN32_FIND_DATA
    Dim hFile        As Long
    Dim bFileExisted As Boolean
    Dim lFileSize    As Long
   
    hFile = -1
   
    lR = ijlInit(tJ)

    If lR = IJL_OK Then
      
        ' Check if we're attempting to overwrite an existing file.
        ' If so hFile <> INVALID_FILE_HANDLE:
        bFileExisted = (FindFirstFile(sFile, tFnd) <> -1)

        If bFileExisted Then
            Kill sFile

        End If
      
        ' Set up the DIB information:
        ' Store DIBWidth:
        tJ.DIBWidth = cDib.Width
        ' Store DIBHeight:
        tJ.DIBHeight = -cDib.Height
        ' Store DIBBytes (pointer to uncompressed JPG data):
        tJ.DIBBytes = cDib.DIBSectionBitsPtr
        ' Very important: tell IJL how many bytes extra there
        ' are on each DIB scan line to pad to 32 bit boundaries:
        tJ.DIBPadBytes = cDib.BytesPerScanLine - cDib.Width * 3
      
        ' Set up the JPEG information:
      
        ' Store JPGFile:
        bFile = StrConv(sFile, vbFromUnicode)
        ReDim Preserve bFile(0 To UBound(bFile) + 1) As Byte
        bFile(UBound(bFile)) = 0
        lPtr = VarPtr(bFile(0))
        CopyMemory tJ.JPGFile, lPtr, 4
        ' Store JPGWidth:
        tJ.JPGWidth = cDib.Width
        ' .. & JPGHeight member values:
        tJ.JPGHeight = cDib.Height
        ' Set the quality/compression to save:
        tJ.jquality = lQuality
            
        ' Write the image:
        lR = ijlWrite(tJ, IJL_JFILE_WRITEWHOLEIMAGE)
      
        ' Check for success:
        If lR = IJL_OK Then
      
            ' Now if we are replacing an existing file, then we want to
            ' put the file creation and archive information back again:
            If bFileExisted Then
            
                hFile = lopen(sFile, OF_WRITE Or OF_SHARE_DENY_WRITE)

                If hFile = 0 Then
                    ' problem
                Else
                    SetFileTime hFile, tFnd.ftCreationTime, tFnd.ftLastAccessTime, tFnd.ftLastWriteTime
                    lclose hFile
                    SetFileAttributes sFile, tFnd.dwFileAttributes

                End If
            
            End If
         
            lFileSize = tJ.JPGSizeBytes - tJ.JPGBytes
         
            ' Success:
            SaveJPG = True
         
        Else
            ' Throw error
            Err.Raise 26001, "No se pudo Guarrdar el JPG" & lR, vbExclamation

        End If
      
        ' Ensure we have freed memory:
        ijlFree tJ
    Else
        ' Throw error:
        Err.Raise 26001, App.EXEName & ".mIntelJPEGLibrary", "No se pudo inicializar la Libreria " & lR

    End If

    
    Exit Function

SaveJPG_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modScreenCapture" & "->" & "SaveJPG"
    End If
Resume Next
    
End Function

Public Sub ScreenCapture(Optional ByVal Autofragshooter As Boolean = False)

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 11/16/2006
    '11/16/2010: Amraphen - Now the FragShooter screenshots are stored in different directories.
    '**************************************************************
    On Error GoTo Err:

    Dim hwnd As Long
    Dim File As String
    Dim sI   As String
    Dim c    As cDIBSection
    Set c = New cDIBSection
    Dim i       As Long
    Dim hdcc    As Long
    
    Dim dirFile As String
    
    hdcc = GetDC(frmMain.hwnd)
    
    frmScreenshots.Picture1.AutoRedraw = True
    frmScreenshots.Picture1.Width = 12090
    frmScreenshots.Picture1.Height = 9075

    Call BitBlt(frmScreenshots.Picture1.hdc, 0, 0, 800, 600, hdcc, 0, 0, SRCCOPY)
    Call ReleaseDC(frmMain.hwnd, hdcc)
    
    hdcc = INVALID_HANDLE
    
    ' Primero chequea si existe la carpeta Screenshots
    dirFile = App.path & "\Screenshots"

    If Not FileExist(dirFile, vbDirectory) Then Call MkDir(dirFile)
    
    ' Si es una imagen de Autofragshooter, se fija si existe la carpeta.
    If Autofragshooter Then
        dirFile = dirFile & "\FragShooter"

        If Not FileExist(dirFile, vbDirectory) Then Call MkDir(dirFile)
        
        'Nuevos directorios del FragShooter:
        If FragShooterKilledSomeone Then 'Si mató a alguien.
            dirFile = dirFile & "\Frags"
        Else 'Si nos mató alguien.
            dirFile = dirFile & "\Muertes"

        End If

        If Not FileExist(dirFile, vbDirectory) Then Call MkDir(dirFile)
        
        'Nuevo formato de las screenshots del FragShooter: "VICTIMA/ASESINO(DD-MM-YYYY hh-mm-ss).jpg"
        File = dirFile & "\" & FragShooterNickname & "(" & Format$(Now, "DD-MM-YYYY hh-mm-ss") & ").jpg"
    Else
        'Si no es screenshot del FragShooter, entonces se usa el formato "DD-MM-YYYY hh-mm-ss.jpg"
        File = dirFile & "\" & Format$(Now, "DD-MM-YYYY hh-mm-ss") & ".jpg"

    End If
    
    frmScreenshots.Picture1.Refresh
    frmScreenshots.Picture1.Picture = frmScreenshots.Picture1.Image
    
    c.CreateFromPicture frmScreenshots.Picture1.Picture
    
    SaveJPG c, File
    
    AddtoRichTextBox frmMain.RecTxt, "Screen Capturada!", 200, 200, 200, False, False, True
    Exit Sub

Err:
    Call AddtoRichTextBox(frmMain.RecTxt, Err.number & "-" & Err.Description, 200, 200, 200, False, False, True)
    
    If hdcc <> INVALID_HANDLE Then Call ReleaseDC(frmMain.hwnd, hdcc)

End Sub
