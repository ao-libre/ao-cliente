Attribute VB_Name = "modCompression"
Option Explicit

Public Const PNG_SOURCE_FILE_EXT  As String = ".png"
Public Const BMP_SOURCE_FILE_EXT  As String = ".bmp"
Public Const GRH_RESOURCE_FILE    As String = "Graphics.AO"
Public Const GRH_PATCH_FILE       As String = "Graficos.PATCH"
Public Const MAPS_SOURCE_FILE_EXT As String = ".map"
Public Const MAPS_RESOURCE_FILE   As String = "Mapas.AO"
Public Const MAPS_PATCH_FILE      As String = "Mapas.PATCH"

Public GrhDatContra()             As Byte ' Contrase�a
Public GrhUsaContra               As Boolean ' Usa Contrase�a?

Public MapsDatContra()            As Byte ' Contrase�a
Public MapsUsaContra              As Boolean  ' Usa Contrase�a?

'This structure will describe our binary file's
'size, number and version of contained files
Public Type FILEHEADER

    lngNumFiles As Long                 'How many files are inside?
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    lngFileVersion As Long              'The resource version (Used to patch)

End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER

    lngFileSize As Long             'How big is this chunk of stored data?
    lngFileStart As Long            'Where does the chunk start?
    strFileName As String * 16      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed

End Type

Private Enum PatchInstruction

    Delete_File
    Create_File
    Modify_File

End Enum

Private Declare Function compress _
                Lib "zlib.dll" (dest As Any, _
                                destlen As Any, _
                                src As Any, _
                                ByVal srclen As Long) As Long
Private Declare Function uncompress _
                Lib "zlib.dll" (dest As Any, _
                                destlen As Any, _
                                src As Any, _
                                ByVal srclen As Long) As Long

'BitMaps Strucures
Public Type BITMAPFILEHEADER

    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long

End Type

Public Type BITMAPINFOHEADER

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

Public Type RGBQUAD

    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte

End Type

Public Type BITMAPINFO

    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD

End Type

'To get free bytes in drive
Private Declare Function GetDiskFreeSpace _
                Lib "kernel32" _
                Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
                                             FreeBytesToCaller As Currency, _
                                             bytesTotal As Currency, _
                                             FreeBytesTotal As Currency) As Long

Public Sub GenerateContra(ByVal Contra As String, Optional Modo As Byte = 0)
    '***************************************************
    'Author: ^[GS]^
    'Last Modification: 17/06/2012 - ^[GS]^
    '
    '***************************************************

    On Error Resume Next

    Dim LoopC As Byte
    Dim Upper_grhDatContra As Long
    Dim Upper_mapsDatContra As Long
    
    If Modo = 0 Then
        Erase GrhDatContra
    ElseIf Modo = 1 Then
        Erase MapsDatContra

    End If
    
    If LenB(Contra) <> 0 Then
        If Modo = 0 Then
            ReDim GrhDatContra(Len(Contra) - 1)
            
            Upper_grhDatContra = UBound(GrhDatContra)
            
            For LoopC = 0 To Upper_grhDatContra
                GrhDatContra(LoopC) = Asc(mid$(Contra, LoopC + 1, 1))
            Next LoopC

            GrhUsaContra = True
        ElseIf Modo = 1 Then
            ReDim MapsDatContra(Len(Contra) - 1)
            
            Upper_mapsDatContra = UBound(MapsDatContra)
                            
            For LoopC = 0 To Upper_mapsDatContra
                MapsDatContra(LoopC) = Asc(mid$(Contra, LoopC + 1, 1))
            Next LoopC

            MapsUsaContra = True

        End If

    Else

        If Modo = 0 Then
            GrhUsaContra = False
        ElseIf Modo = 1 Then
            MapsUsaContra = False

        End If

    End If
    
End Sub

Private Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 6/07/2004
    '
    '**************************************************************
    
    On Error GoTo General_Drive_Get_Free_Bytes_Err
    
    Dim retval As Long
    Dim FB     As Currency
    Dim BT     As Currency
    Dim FBT    As Currency
    
    retval = GetDiskFreeSpace(Left$(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes

    
    Exit Function

General_Drive_Get_Free_Bytes_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "General_Drive_Get_Free_Bytes"
    End If
Resume Next
    
End Function

''
' Sorts the info headers by their file name. Uses QuickSort.
'
' @param    InfoHead() The array of headers to be ordered.
' @param    first The first index in the list.
' @param    last The last index in the list.

Private Sub Sort_Info_Headers(ByRef InfoHead() As INFOHEADER, _
                              ByVal first As Long, _
                              ByVal last As Long)
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/20/2007
    'Sorts the info headers by their file name using QuickSort.
    '*****************************************************************
    
    On Error GoTo Sort_Info_Headers_Err
    
    Dim aux  As INFOHEADER
    Dim Min  As Long
    Dim Max  As Long
    Dim comp As String
    
    Min = first
    Max = last
    
    comp = InfoHead((Min + Max) \ 2).strFileName
    
    Do While Min <= Max
        Do While InfoHead(Min).strFileName < comp And Min < last
            Min = Min + 1
        Loop

        Do While InfoHead(Max).strFileName > comp And Max > first
            Max = Max - 1
        Loop

        If Min <= Max Then
            aux = InfoHead(Min)
            InfoHead(Min) = InfoHead(Max)
            InfoHead(Max) = aux
            Min = Min + 1
            Max = Max - 1

        End If

    Loop
    
    If first < Max Then Call Sort_Info_Headers(InfoHead, first, Max)
    If Min < last Then Call Sort_Info_Headers(InfoHead, Min, last)

    
    Exit Sub

Sort_Info_Headers_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Sort_Info_Headers"
    End If
Resume Next
    
End Sub

''
' Searches for the specified InfoHeader.
'
' @param    ResourceFile A handler to the data file.
' @param    InfoHead The header searched.
' @param    FirstHead The first head to look.
' @param    LastHead The last head to look.
' @param    FileHeaderSize The bytes size of a FileHeader.
' @param    InfoHeaderSize The bytes size of a InfoHeader.
'
' @return   True if found.
'
' @remark   File must be already open.
' @remark   InfoHead must have set its file name to perform the search.

Private Function BinarySearch(ByRef ResourceFile As Integer, _
                              ByRef InfoHead As INFOHEADER, _
                              ByVal FirstHead As Long, _
                              ByVal LastHead As Long, _
                              ByVal FileHeaderSize As Long, _
                              ByVal InfoHeaderSize As Long) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/21/2007
    'Searches for the specified InfoHeader
    '*****************************************************************
    
    On Error GoTo BinarySearch_Err
    
    Dim ReadingHead  As Long
    Dim ReadInfoHead As INFOHEADER
    
    Do Until FirstHead > LastHead
        ReadingHead = (FirstHead + LastHead) \ 2

        Get ResourceFile, FileHeaderSize + InfoHeaderSize * (ReadingHead - 1) + 1, ReadInfoHead

        If InfoHead.strFileName = ReadInfoHead.strFileName Then
            InfoHead = ReadInfoHead
            BinarySearch = True
            Exit Function
        Else

            If InfoHead.strFileName < ReadInfoHead.strFileName Then
                LastHead = ReadingHead - 1
            Else
                FirstHead = ReadingHead + 1

            End If

        End If

    Loop

    
    Exit Function

BinarySearch_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "BinarySearch"
    End If
Resume Next
    
End Function

''
' Retrieves the InfoHead of the specified graphic file.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    InfoHead The InfoHead where data is returned.
'
' @return   True if found.

Private Function Get_InfoHeader(ByRef ResourcePath As String, _
                                ByRef FileName As String, _
                                ByRef InfoHead As INFOHEADER, _
                                Optional Modo As Byte = 0) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 16/07/2012 - ^[GS]^
    'Retrieves the InfoHead of the specified graphic file
    '*****************************************************************
    
    On Error GoTo Get_InfoHeader_Err
    
    Dim ResourceFile     As Integer
    Dim ResourceFilePath As String
    Dim FileHead         As FILEHEADER
    
<<<<<<< HEAD
    On Local Error GoTo ErrHandler
=======
    Dim ERROR_LEER_ARCHIVO As String
    
On Local Error GoTo ErrHandler
>>>>>>> origin/master

    If Modo = 0 Then
        ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    ElseIf Modo = 1 Then
        ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE

    End If
    
    'Set InfoHeader we are looking for
    InfoHead.strFileName = UCase$(FileName)
   
    'Open the binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
    'Extract the FILEHEADER
    Get ResourceFile, 1, FileHead
        
<<<<<<< HEAD
    'Check the file for validity
    If LOF(ResourceFile) <> FileHead.lngFileSize Then
        MsgBox "Archivo de recursos da�ado. " & ResourceFilePath, , "Error"
        Close ResourceFile
        Exit Function

    End If
=======
        'Check the file for validity
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            MsgBox JsonLanguage.Item("ERROR_ARCHIVO_CORRUPTO").Item("TEXTO") & ": " & ResourceFilePath, , JsonLanguage.Item("Error").Item("TEXTO")
            Close ResourceFile
            Exit Function
        End If
>>>>>>> origin/master
        
    'Search for it!
    If BinarySearch(ResourceFile, InfoHead, 1, FileHead.lngNumFiles, Len(FileHead), Len(InfoHead)) Then
        Get_InfoHeader = True

    End If
        
    Close ResourceFile
    Exit Function

ErrHandler:
    Close ResourceFile
    
<<<<<<< HEAD
    Call MsgBox("Error al intentar leer el archivo " & ResourceFilePath & ". Raz�n: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")

    
    Exit Function

Get_InfoHeader_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Get_InfoHeader"
    End If
Resume Next
    
=======
    ERROR_LEER_ARCHIVO = JsonLanguage.Item("ERROR_LEER_ARCHIVO").Item("TEXTO")
    ERROR_LEER_ARCHIVO = Replace$(ERROR_LEER_ARCHIVO, "VAR_ARCHIVO", ResourceFilePath)
    ERROR_LEER_ARCHIVO = Replace$(ERROR_LEER_ARCHIVO, "VAR_ERROR", Err.number & " : " & Err.Description)
    
    Call MsgBox(ERROR_LEER_ARCHIVO)
>>>>>>> origin/master
End Function

''
' Compresses binary data avoiding data loses.
'
' @param    data() The data array.

Private Sub Compress_Data(ByRef data() As Byte, Optional Modo As Byte = 0)
    '*****************************************************************
    'Author: Juan Mart�n Dotuyo Dodero
    'Last Modify Date: 17/07/2012 - ^[GS]^
    'Compresses binary data avoiding data loses
    '*****************************************************************
    
    On Error GoTo Compress_Data_Err
    
    Dim Dimensions         As Long
    Dim DimBuffer          As Long
    Dim BufTemp()          As Byte
    
    Dim LoopC              As Long
    Dim Upper_grhDatContra As Long
    Dim Upper_mapsDatContra As Long
    
    Dimensions = UBound(data) + 1
    
    ' The worst case scenario, compressed info is 1.06 times the original - see zlib's doc for more info.
    DimBuffer = Dimensions * 1.06
    
    ReDim BufTemp(DimBuffer)
    
    Call compress(BufTemp(0), DimBuffer, data(0), Dimensions)
    
    Erase data
    
    ReDim data(DimBuffer - 1)
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    data = BufTemp
    
    Erase BufTemp
    
    ' GSZAO - Seguridad
    If Modo = 0 And GrhUsaContra = True Then
        If UBound(GrhDatContra) <= UBound(data) And UBound(GrhDatContra) <> 0 Then
            
            Upper_grhDatContra = UBound(GrhDatContra)
            
            For LoopC = 0 To Upper_grhDatContra
                data(LoopC) = data(LoopC) Xor GrhDatContra(LoopC)
            Next LoopC

        End If

    ElseIf Modo = 1 And MapsUsaContra = True Then

        If UBound(MapsDatContra) <= UBound(data) And UBound(MapsDatContra) <> 0 Then
            
            Upper_mapsDatContra = UBound(MapsDatContra)
            
            For LoopC = 0 To Upper_mapsDatContra
                data(LoopC) = data(LoopC) Xor MapsDatContra(LoopC)
            Next LoopC

        End If

    End If

    ' GSZAO - Seguridad
    
    Exit Sub

Compress_Data_Err:

    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Compress_Data"

    End If

    Resume Next
    
End Sub

''
' Decompresses binary data.
'
' @param    data() The data array.
' @param    OrigSize The original data size.

Private Sub Decompress_Data(ByRef data() As Byte, _
                            ByVal OrigSize As Long, _
                            Optional Modo As Byte = 0)
    '*****************************************************************
    'Author: Juan Mart�n Dotuyo Dodero
    'Last Modify Date: 16/07/2012 - ^[GS]^
    'Decompresses binary data
    '*****************************************************************
    
    On Error GoTo Decompress_Data_Err
    
    Dim BufTemp()           As Byte
    
    Dim LoopC               As Integer
    Dim Upper_grhDatContra  As Long
    Dim Upper_mapDataContra As Long
    
    ReDim BufTemp(OrigSize - 1)
    
    ' GSZAO - Seguridad
    If Modo = 0 And GrhUsaContra = True Then
        If UBound(GrhDatContra) <= UBound(data) And UBound(GrhDatContra) <> 0 Then
            
            Upper_grhDatContra = UBound(GrhDatContra)
            
            For LoopC = 0 To Upper_grhDatContra
                data(LoopC) = data(LoopC) Xor GrhDatContra(LoopC)
            Next LoopC

        End If

    ElseIf Modo = 1 And MapsUsaContra = True Then

        If UBound(MapsDatContra) <= UBound(data) And UBound(MapsDatContra) <> 0 Then
            
            Upper_mapDataContra = UBound(MapsDatContra)
            
            For LoopC = 0 To Upper_mapDataContra
                data(LoopC) = data(LoopC) Xor MapsDatContra(LoopC)
            Next LoopC

        End If

    End If

    ' GSZAO - Seguridad
    
    Call uncompress(BufTemp(0), OrigSize, data(0), UBound(data) + 1)
    
    ReDim data(OrigSize - 1)
    
    data = BufTemp
    
    Erase BufTemp
    
    Exit Sub

Decompress_Data_Err:

    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Decompress_Data"

    End If

    Resume Next
    
End Sub

''
' Compresses all graphic files to a resource file.
'
' @param    SourcePath The graphic files folder.
' @param    OutputPath The resource file folder.
' @param    version The resource file version.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function Compress_Files(ByRef SourcePath As String, _
                               ByRef OutputPath As String, _
                               ByVal version As Long, _
                               ByRef prgBar As ProgressBar, _
                               Optional Modo As Byte = 0) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 14/09/2012 - ^[GS]^
    'Compresses all graphic files to a resource file
    '*****************************************************************
    
    On Error GoTo Compress_Files_Err
    
    Dim SourceFileName As String
    Dim OutputFilePath As String
<<<<<<< HEAD
    Dim SourceFile     As Long
    Dim OutputFile     As Long
    Dim SourceData()   As Byte
    Dim FileHead       As FILEHEADER
    Dim InfoHead()     As INFOHEADER
    Dim LoopC          As Long

    On Local Error GoTo ErrHandler

=======
    Dim SourceFile As Long
    Dim OutputFile As Long
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim LoopC As Long
    
    Dim ERROR_EXT_NO_ENCONTRADA As String

On Local Error GoTo ErrHandler
>>>>>>> origin/master
    If Modo = 0 Then
        OutputFilePath = OutputPath & GRH_RESOURCE_FILE
        'If GraficosPNG = False Then ' GSZAO
        SourceFileName = Dir$(SourcePath & "*" & BMP_SOURCE_FILE_EXT, vbNormal)
        'Else
        '   SourceFileName = Dir$(SourcePath & "*" & PNG_SOURCE_FILE_EXT, vbNormal)
        'End If
    ElseIf Modo = 1 Then
        OutputFilePath = OutputPath & MAPS_RESOURCE_FILE
        SourceFileName = Dir$(SourcePath & "*" & MAPS_SOURCE_FILE_EXT, vbNormal)

    End If
    
    ' Create list of all files to be compressed
    While LenB(SourceFileName) <> 0

        FileHead.lngNumFiles = FileHead.lngNumFiles + 1
        
        ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
        InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
        
        'Search new file
        SourceFileName = Dir$()
    Wend
    
    'If Mode = 0 And frmMain.cmdGrhPNG.Value = 1 Then ' Comprimimos tambien los Graficos .PNG
    SourceFileName = Dir$(SourcePath & "*" & PNG_SOURCE_FILE_EXT, vbNormal)

    ' Create list of all files to be compressed
    While LenB(SourceFileName) <> 0

        FileHead.lngNumFiles = FileHead.lngNumFiles + 1
            
        ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
        InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
            
        'Search new file
        SourceFileName = Dir$()
    Wend
    'End If
    
    'If Mode = 1 And frmMain.cmdMiniMap.Value = 1 Then ' agregamos tambien los BMP junto a los mapas
    SourceFileName = Dir$(SourcePath & "*" & BMP_SOURCE_FILE_EXT, vbNormal)  ' GSZAO

    ' Create list of all files to be compressed
    While LenB(SourceFileName) <> 0

        FileHead.lngNumFiles = FileHead.lngNumFiles + 1
            
        ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
        InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
            
        'Search new file
        SourceFileName = Dir$()
    Wend
    'End If
    
    If FileHead.lngNumFiles = 0 Then
        'If GraficosPNG = False Then ' GSZAO
<<<<<<< HEAD
        MsgBox "No se encontraron archivos de extensi�n " & BMP_SOURCE_FILE_EXT & " en " & SourcePath & ".", , "Error"
=======
            ERROR_EXT_NO_ENCONTRADA = JsonLanguage.Item("ERROR_EXT_NO_ENCONTRADA").Item("TEXTO")
            ERROR_EXT_NO_ENCONTRADA = Replace$(ERROR_EXT_NO_ENCONTRADA, "VAR_EXT", BMP_SOURCE_FILE_EXT)
            ERROR_EXT_NO_ENCONTRADA = Replace$(ERROR_EXT_NO_ENCONTRADA, "VAR_PATH", SourcePath)
            
            MsgBox ERROR_EXT_NO_ENCONTRADA, , JsonLanguage.Item("Error").Item("TEXTO")
>>>>>>> origin/master
        'Else
        '    MsgBox "No se encontraron archivos de extensi�n " & PNG_SOURCE_FILE_EXT & " en " & SourcePath & ".", , JsonLanguage.Item("Error").Item("TEXTO")
        'End If
        Exit Function

    End If
    
    If Not prgBar Is Nothing Then
        prgBar.value = 0
        prgBar.Max = FileHead.lngNumFiles + 1

    End If
    
    'Destroy file if it previuosly existed
    If LenB(Dir$(OutputFilePath, vbNormal)) <> 0 Then
        Kill OutputFilePath

    End If
    
    'Finish setting the FileHeader data
    FileHead.lngFileVersion = version
    FileHead.lngFileSize = Len(FileHead) + FileHead.lngNumFiles * Len(InfoHead(0))
    
    'Order the InfoHeads
    Call Sort_Info_Headers(InfoHead(), 0, FileHead.lngNumFiles - 1)
    
    'Open a new file
    OutputFile = FreeFile()
    Open OutputFilePath For Binary Access Read Write As OutputFile
    ' Move to the end of the headers, where the file data will actually start
    Seek OutputFile, FileHead.lngFileSize + 1
        
    ' Process every file!
    For LoopC = 0 To FileHead.lngNumFiles - 1
              
        SourceFile = FreeFile()
        Open SourcePath & InfoHead(LoopC).strFileName For Binary Access Read Lock Write As SourceFile
                
        'Find out how large the file is and resize the data array appropriately
        InfoHead(LoopC).lngFileSizeUncompressed = LOF(SourceFile)
        ReDim SourceData(LOF(SourceFile) - 1)
                
        'Get the data from the file
        Get SourceFile, , SourceData
                
        'Compress it
        Call Compress_Data(SourceData, Modo)
                
        'Store it in the resource file
        Put OutputFile, , SourceData
                
        With InfoHead(LoopC)
            'Set up the info headers
            .lngFileSize = UBound(SourceData) + 1
            .lngFileStart = FileHead.lngFileSize + 1
                    
            'Update the file header
            FileHead.lngFileSize = FileHead.lngFileSize + .lngFileSize

        End With
                
        Erase SourceData
            
        Close SourceFile
        
<<<<<<< HEAD
        'Update progress bar
        If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
        DoEvents
    Next LoopC
=======
            'Update progress bar
            If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
            DoEvents
        Next LoopC
>>>>>>> origin/master
        
    'Store the headers in the file
    Seek OutputFile, 1
    Put OutputFile, , FileHead
    Put OutputFile, , InfoHead
        
    'Close the file
    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
    
    Compress_Files = True
    Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead
    Close OutputFile
    
<<<<<<< HEAD
    Call MsgBox("No se pudo crear el archivo binario. Raz�n: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")

    
    Exit Function

Compress_Files_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Compress_Files"
    End If
Resume Next
    
=======
    Call MsgBox(Replace$(JsonLanguage.Item("ERROR_CREAR_BINARIO").Item("TEXTO"), "VAR_ERROR", Err.number & " : " & Err.Description), vbOKOnly, JsonLanguage.Item("Error").Item("TEXTO"))
>>>>>>> origin/master
End Function

''
' Retrieves a byte array with the compressed data from the specified file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   InfoHead must not be encrypted.
' @remark   Data is not desencrypted.

Public Function Get_File_RawData(ByRef ResourcePath As String, _
                                 ByRef InfoHead As INFOHEADER, _
                                 ByRef data() As Byte, _
                                 Optional Modo As Byte = 0) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 16/07/2012 - ^[GS]^
    'Retrieves a byte array with the compressed data from the specified file
    '*****************************************************************
    
    On Error GoTo Get_File_RawData_Err
    
    Dim ResourceFilePath As String
    Dim ResourceFile     As Integer
    
    On Local Error GoTo ErrHandler

    If Modo = 0 Then
        ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    ElseIf Modo = 1 Then
        ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE

    End If
    
    'Size the Data array
    ReDim data(InfoHead.lngFileSize - 1)
    
    'Open the binary file
    ResourceFile = FreeFile
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
    'Get the data
    Get ResourceFile, InfoHead.lngFileStart, data
    'Close the binary file
    Close ResourceFile
    
    Get_File_RawData = True
    Exit Function

ErrHandler:
    Close ResourceFile

    
    Exit Function

Get_File_RawData_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Get_File_RawData"
    End If
Resume Next
    
End Function

''
' Extract the specific file from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Extract_File(ByRef ResourcePath As String, _
                             ByRef InfoHead As INFOHEADER, _
                             ByRef data() As Byte, _
                             Optional Modo As Byte = 0) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 14/09/2012 - ^[GS]^
    'Extract the specific file from a resource file
    '*****************************************************************
    
    On Error GoTo Extract_File_Err
    
    On Local Error GoTo ErrHandler
    
    If Get_File_RawData(ResourcePath, InfoHead, data, Modo) Then
        'Decompress all data
        'If InfoHead.lngFileSize < InfoHead.lngFileSizeUncompressed Then ' GSZAO
        Call Decompress_Data(data, InfoHead.lngFileSizeUncompressed, Modo)
        'End If
        
        Extract_File = True

    End If

    Exit Function

ErrHandler:
<<<<<<< HEAD
    Call MsgBox("Error al intentar decodificar recursos. Raz�n: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")

    
    Exit Function

Extract_File_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Extract_File"
    End If
Resume Next
    
=======
    Call MsgBox(Replace$(JsonLanguage.Item("ERROR_DECODE_RECURSOS").Item("TEXTO"), "VAR_ERROR", Err.number & " : " & Err.Description), vbOKOnly, JsonLanguage.Item("Error").Item("TEXTO"))
>>>>>>> origin/master
End Function

''
' Extracts all files from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    OutputPath The folder where graphic files will be extracted.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function Extract_Files(ByRef ResourcePath As String, _
                              ByRef OutputPath As String, _
                              ByRef prgBar As ProgressBar, _
                              Optional Modo As Byte = 0) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 17/07/2012 - ^[GS]^
    'Extracts all files from a resource file
    '*****************************************************************
    
    On Error GoTo Extract_Files_Err
    
    Dim LoopC            As Long
    Dim ResourceFile     As Integer
    Dim ResourceFilePath As String
    Dim OutputFile       As Integer
    Dim SourceData()     As Byte
    Dim FileHead         As FILEHEADER
    Dim InfoHead()       As INFOHEADER
    Dim RequiredSpace    As Currency
    
    Dim Upper_infoHead As Long
    
    On Local Error GoTo ErrHandler

    If Modo = 0 Then
        ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    ElseIf Modo = 1 Then
        ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE

    End If
    
    'Open the binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
    'Extract the FILEHEADER
    Get ResourceFile, 1, FileHead
    
<<<<<<< HEAD
    'Check the file for validity
    If LOF(ResourceFile) <> FileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos da�ado. " & ResourceFilePath, , "Error")
        Close ResourceFile
        Exit Function

    End If
=======
        'Check the file for validity
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            Call MsgBox(JsonLanguage.Item("ERROR_ARCHIVO_CORRUPTO").Item("TEXTO") & ": " & ResourceFilePath, , JsonLanguage.Item("Error").Item("TEXTO"))
            Close ResourceFile
            Exit Function
        End If
>>>>>>> origin/master
        
    'Size the InfoHead array
    ReDim InfoHead(FileHead.lngNumFiles - 1)
        
    'Extract the INFOHEADER
    Get ResourceFile, , InfoHead
        
    'Check if there is enough hard drive space to extract all files
    Upper_infoHead = UBound(InfoHead)
    For LoopC = 0 To Upper_infoHead
            
        RequiredSpace = RequiredSpace + InfoHead(LoopC).lngFileSizeUncompressed
    Next LoopC
        
<<<<<<< HEAD
    If RequiredSpace >= General_Drive_Get_Free_Bytes(Left$(App.path, 3)) Then
        Erase InfoHead
        Close ResourceFile
        Call MsgBox("No hay suficiente espacio en el disco para extraer los archivos.", , "Error")
        Exit Function

    End If

=======
        If RequiredSpace >= General_Drive_Get_Free_Bytes(Left$(App.path, 3)) Then
            Erase InfoHead
            Close ResourceFile
            Call MsgBox(JsonLanguage.Item("ERROR_SIN_ESPACIO").Item("TEXTO"), , JsonLanguage.Item("Error").Item("TEXTO"))
            Exit Function
        End If
>>>>>>> origin/master
    Close ResourceFile
    
    'Update progress bar
    If Not prgBar Is Nothing Then
        prgBar.value = 0
        prgBar.Max = FileHead.lngNumFiles + 1

    End If
    
    'Extract all of the files from the binary file
    Upper_infoHead = UBound(InfoHead)
    For LoopC = 0 To Upper_infoHead

        'Extract this file
        If Extract_File(ResourcePath, InfoHead(LoopC), SourceData) Then

            'Destroy file if it previuosly existed
            If FileExist(OutputPath & InfoHead(LoopC).strFileName, vbNormal) Then
                Call Kill(OutputPath & InfoHead(LoopC).strFileName)

            End If
            
            'Save it!
            OutputFile = FreeFile()
            Open OutputPath & InfoHead(LoopC).strFileName For Binary As OutputFile
            Put OutputFile, , SourceData
            Close OutputFile
            
            Erase SourceData
        Else
            Erase SourceData
            Erase InfoHead
            
            Call MsgBox(JsonLanguage.Item("ERROR_EXTRAER_ARCHIVO").Item("TEXTO") & ": " & InfoHead(LoopC).strFileName, vbOKOnly, JsonLanguage.Item("Error").Item("TEXTO"))
            Exit Function

        End If
            
        'Update progress bar
        If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
        DoEvents
    Next LoopC
    
    Erase InfoHead
    Extract_Files = True
    Exit Function

ErrHandler:
    Close ResourceFile
    Erase SourceData
    Erase InfoHead
    
<<<<<<< HEAD
    Call MsgBox("No se pudo extraer el archivo binario correctamente. Raz�n: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")

    
    Exit Function

Extract_Files_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Extract_Files"
    End If
Resume Next
    
=======
    Call MsgBox(Replace$(JsonLanguage.Item("ERROR_EXTRAER_BINARIO").Item("TEXTO"), "VAR_ERROR", Err.number & " : " & Err.Description), vbOKOnly, JsonLanguage.Item("Error").Item("TEXTO"))
>>>>>>> origin/master
End Function

''
' Retrieves a byte array with the specified file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Get_File_Data(ByRef ResourcePath As String, _
                              ByRef FileName As String, _
                              ByRef data() As Byte, _
                              Optional Modo As Byte = 0) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 16/07/2012 - ^[GS]^
    'Retrieves a byte array with the specified file data
    '*****************************************************************
    
    On Error GoTo Get_File_Data_Err
    
    Dim InfoHead As INFOHEADER
    
    If Get_InfoHeader(ResourcePath, FileName, InfoHead, Modo) Then
        'Extract!
        Get_File_Data = Extract_File(ResourcePath, InfoHead, data, Modo)
    Else
        Get_File_Data = False
<<<<<<< HEAD

        'Call MsgBox("No se se encontro el recurso " & FileName)
=======
        'Call MsgBox(JsonLanguage("ERROR_404").Item("TEXTO") & ": " & FileName)
>>>>>>> origin/master
    End If

    
    Exit Function

Get_File_Data_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Get_File_Data"
    End If
Resume Next
    
End Function

''
' Retrieves image file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.

Public Function Get_Image(ByRef ResourcePath As String, _
                          ByRef FileName As String, _
                          ByRef data() As Byte, _
                          Optional SoloBMP As Boolean = False) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 09/10/2012 - ^[GS]^
    'Retrieves image file data
    '*****************************************************************
    
    On Error GoTo Get_Image_Err
    
    Dim InfoHead  As INFOHEADER
    Dim ExistFile As Boolean
    
    ExistFile = False
    
    If SoloBMP = True Then
        If Get_InfoHeader(ResourcePath, FileName & ".BMP", InfoHead, 0) Then ' �BMP?
            FileName = FileName & ".BMP"
            ExistFile = True

        End If

    Else

        If Get_InfoHeader(ResourcePath, FileName & ".BMP", InfoHead, 0) Then ' �BMP?
            FileName = FileName & ".BMP"
            ExistFile = True
        ElseIf Get_InfoHeader(ResourcePath, FileName & ".PNG", InfoHead, 0) Then ' Existe PNG?
            FileName = FileName & ".PNG" ' usamos el PNG
            ExistFile = True

        End If

    End If
    
    If ExistFile = True Then
        If Extract_File(ResourcePath, InfoHead, data, 0) Then Get_Image = True
    Else
<<<<<<< HEAD
        Call MsgBox("Get_Image::No se encontro el recurso " & FileName)

=======
        Call MsgBox(JsonLanguage("ERROR_404").Item("TEXTO") & ": " & FileName)
>>>>>>> origin/master
    End If

    
    Exit Function

Get_Image_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Get_Image"
    End If
Resume Next
    
End Function

''
' Compare two byte arrays to detect any difference.
'
' @param    data1() Byte array.
' @param    data2() Byte array.
'
' @return   True if are equals.

Private Function Compare_Datas(ByRef data1() As Byte, ByRef data2() As Byte) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 02/11/2007
    'Compare two byte arrays to detect any difference
    '*****************************************************************
    
    On Error GoTo Compare_Datas_Err
    
    Dim length As Long
    Dim act    As Long
    
    length = UBound(data1) + 1
    
    If (UBound(data2) + 1) = length Then

        While act < length

            If data1(act) Xor data2(act) Then Exit Function
            
            act = act + 1
        Wend
        
        Compare_Datas = True

    End If

    
    Exit Function

Compare_Datas_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Compare_Datas"
    End If
Resume Next
    
End Function

''
' Retrieves the next InfoHeader.
'
' @param    ResourceFile A handler to the resource file.
' @param    FileHead The reource file header.
' @param    InfoHead The returned header.
' @param    ReadFiles The number of headers that have already been read.
'
' @return   False if there are no more headers tu read.
'
' @remark   File must be already open.
' @remark   Used to walk through the resource file info headers.
' @remark   The number of read files will increase although there is nothing else to read.
' @remark   InfoHead is encrypted.

Private Function ReadNext_InfoHead(ByRef ResourceFile As Integer, _
                                   ByRef FileHead As FILEHEADER, _
                                   ByRef InfoHead As INFOHEADER, _
                                   ByRef ReadFiles As Long) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/24/2007
    'Reads the next InfoHeader
    '*****************************************************************
    
    On Error GoTo ReadNext_InfoHead_Err
    

    If ReadFiles < FileHead.lngNumFiles Then
        'Read header
        Get ResourceFile, Len(FileHead) + Len(InfoHead) * ReadFiles + 1, InfoHead
        
        'Update
        ReadNext_InfoHead = True

    End If
    
    ReadFiles = ReadFiles + 1

    
    Exit Function

ReadNext_InfoHead_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "ReadNext_InfoHead"
    End If
Resume Next
    
End Function

''
' Retrieves the next bitmap.
'
' @param    ResourcePath The resource file folder.
' @param    ReadFiles The number of bitmaps that have already been read.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   False if there are no more bitmaps tu get.
'
' @remark   Used to walk through the resource file bitmaps.

Public Function GetNext_Bitmap(ByRef ResourcePath As String, _
                               ByRef ReadFiles As Long, _
                               ByRef bmpInfo As BITMAPINFO, _
                               ByRef data() As Byte, _
                               ByRef fileIndex As Long) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 09/10/2012 - ^[GS]^
    'Reads the next InfoHeader
    '*****************************************************************
    On Error Resume Next

    Dim ResourceFile As Integer
    Dim FileHead     As FILEHEADER
    Dim InfoHead     As INFOHEADER
    Dim FileName     As String
    
    ResourceFile = FreeFile
    Open ResourcePath & GRH_RESOURCE_FILE For Binary Access Read Lock Write As ResourceFile
    Get ResourceFile, 1, FileHead
    
    If ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ReadFiles) Then
        Call Get_Image(ResourcePath, InfoHead.strFileName, data())
        FileName = Trim$(InfoHead.strFileName)
        fileIndex = CLng(Left$(FileName, Len(FileName) - 4))
        
        GetNext_Bitmap = True

    End If
    
    Close ResourceFile

End Function

''
' Compares two resource versions and makes a patch file.
'
' @param    NewResourcePath The actual reource file folder.
' @param    OldResourcePath The previous reource file folder.
' @param    OutputPath The patchs file folder.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function Make_Patch(ByRef NewResourcePath As String, _
                           ByRef OldResourcePath As String, _
                           ByRef OutputPath As String, _
                           ByRef prgBar As ProgressBar, _
                           Optional Modo As Byte = 0) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 17/07/2012 - ^[GS]^
    'Compares two resource versions and make a patch file
    '*****************************************************************
    
    On Error GoTo Make_Patch_Err
    
    Dim NewResourceFile     As Integer
    Dim NewResourceFilePath As String
    Dim NewFileHead         As FILEHEADER
    Dim NewInfoHead         As INFOHEADER
    Dim NewReadFiles        As Long
    Dim NewReadNext         As Boolean
    
    Dim OldResourceFile     As Integer
    Dim OldResourceFilePath As String
    Dim OldFileHead         As FILEHEADER
    Dim OldInfoHead         As INFOHEADER
    Dim OldReadFiles        As Long
    Dim OldReadNext         As Boolean
    
    Dim OutputFile          As Integer
    Dim OutputFilePath      As String
    Dim data()              As Byte
    Dim auxData()           As Byte
    Dim Instruction         As Byte
    
    'Set up the error handler
    On Local Error GoTo ErrHandler

    If Modo = 0 Then
        NewResourceFilePath = NewResourcePath & GRH_RESOURCE_FILE
        OldResourceFilePath = OldResourcePath & GRH_RESOURCE_FILE
        OutputFilePath = OutputPath & GRH_PATCH_FILE
    ElseIf Modo = 1 Then
        NewResourceFilePath = NewResourcePath & MAPS_RESOURCE_FILE
        OldResourceFilePath = OldResourcePath & MAPS_RESOURCE_FILE
        OutputFilePath = OutputPath & MAPS_PATCH_FILE

    End If
    
    'Open the old binary file
    OldResourceFile = FreeFile
    Open OldResourceFilePath For Binary Access Read Lock Write As OldResourceFile
        
<<<<<<< HEAD
    'Get the old FileHeader
    Get OldResourceFile, 1, OldFileHead

    'Check the file for validity
    If LOF(OldResourceFile) <> OldFileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos anterior da�ado. " & OldResourceFilePath, , "Error")
        Close OldResourceFile
        Exit Function

    End If
=======
        'Get the old FileHeader
        Get OldResourceFile, 1, OldFileHead
        'Check the file for validity
        If LOF(OldResourceFile) <> OldFileHead.lngFileSize Then
            Call MsgBox(JsonLanguage.Item("ERROR_ARCHIVO_CORRUPTO").Item("TEXTO") & ": " & OldResourceFilePath, , JsonLanguage.Item("Error").Item("TEXTO"))
            Close OldResourceFile
            Exit Function
        End If
>>>>>>> origin/master
        
    'Open the new binary file
    NewResourceFile = FreeFile()
    Open NewResourceFilePath For Binary Access Read Lock Write As NewResourceFile
            
<<<<<<< HEAD
    'Get the new FileHeader
    Get NewResourceFile, 1, NewFileHead

    'Check the file for validity
    If LOF(NewResourceFile) <> NewFileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos anterior da�ado. " & NewResourceFilePath, , "Error")
        Close NewResourceFile
        Close OldResourceFile
        Exit Function

    End If
=======
            'Get the new FileHeader
            Get NewResourceFile, 1, NewFileHead
            'Check the file for validity
            If LOF(NewResourceFile) <> NewFileHead.lngFileSize Then
                Call MsgBox(JsonLanguage.Item("ERROR_ARCHIVO_CORRUPTO").Item("TEXTO") & ": " & NewResourceFilePath, , JsonLanguage.Item("Error").Item("TEXTO"))
                Close NewResourceFile
                Close OldResourceFile
                Exit Function
            End If
>>>>>>> origin/master
            
    'Destroy file if it previuosly existed
    If LenB(Dir$(OutputFilePath, vbNormal)) <> 0 Then Kill OutputFilePath
            
    'Open the patch file
    OutputFile = FreeFile()
    Open OutputFilePath For Binary Access Read Write As OutputFile
                
<<<<<<< HEAD
    If Not prgBar Is Nothing Then
        prgBar.value = 0
        prgBar.Max = (OldFileHead.lngNumFiles + NewFileHead.lngNumFiles) + 1

    End If
=======
                If Not prgBar Is Nothing Then
                    prgBar.value = 0
                    prgBar.Max = (OldFileHead.lngNumFiles + NewFileHead.lngNumFiles) + 1
                End If
>>>>>>> origin/master
                
    'put previous file version (unencrypted)
    Put OutputFile, , OldFileHead.lngFileVersion
                
    'Put the new file header
    Put OutputFile, , NewFileHead

    'Try to read old and new first files
    If ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) And ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                    
<<<<<<< HEAD
        'Update
        prgBar.value = prgBar.value + 2
=======
                    'Update
                    prgBar.value = prgBar.value + 2
>>>>>>> origin/master
                    
        Do 'Main loop

            'Comparisons are between encrypted names, for ordering issues
            If OldInfoHead.strFileName = NewInfoHead.strFileName Then

                'Get old file data
                Call Get_File_RawData(OldResourcePath, OldInfoHead, auxData, Modo)
                            
                'Get new file data
                Call Get_File_RawData(NewResourcePath, NewInfoHead, data, Modo)
                            
                If Not Compare_Datas(data, auxData) Then
                    'File was modified
                    Instruction = PatchInstruction.Modify_File
                    Put OutputFile, , Instruction
                                
                    'Write header
                    Put OutputFile, , NewInfoHead
                                
                    'Write data
                    Put OutputFile, , data

                End If
                            
                'Read next OldResource
                If Not ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then
                    Exit Do

                End If
                            
                'Read next NewResource
                If Not ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                    'Reread last OldInfoHead
                    OldReadFiles = OldReadFiles - 1
                    Exit Do

                End If
                            
<<<<<<< HEAD
                'Update
                If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 2
=======
                            'Update
                            If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 2
>>>>>>> origin/master
                        
            ElseIf OldInfoHead.strFileName < NewInfoHead.strFileName Then
                            
                'File was deleted
                Instruction = PatchInstruction.Delete_File
                Put OutputFile, , Instruction
                Put OutputFile, , OldInfoHead
                            
                'Read next OldResource
                If Not ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then
                    'Reread last NewInfoHead
                    NewReadFiles = NewReadFiles - 1
                    Exit Do

                End If
                            
<<<<<<< HEAD
                'Update
                If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
=======
                            'Update
                            If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
>>>>>>> origin/master
                        
            Else
                            
                'New file
                Instruction = PatchInstruction.Create_File
                Put OutputFile, , Instruction
                Put OutputFile, , NewInfoHead
                                     
                'Get file data
                Call Get_File_RawData(NewResourcePath, NewInfoHead, data, Modo)
                            
                'Write data
                Put OutputFile, , data
                            
                'Read next NewResource
                If Not ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                    'Reread last OldInfoHead
                    OldReadFiles = OldReadFiles - 1
                    Exit Do

                End If
                            
<<<<<<< HEAD
                'Update
                If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1

            End If
=======
                            'Update
                            If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
                        End If
>>>>>>> origin/master
                        
            DoEvents
        Loop
                
    Else
        'if at least one is empty
        OldReadFiles = 0
        NewReadFiles = 0

    End If
                
    'Read everything?
    While ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles)

        'Delete file
        Instruction = PatchInstruction.Delete_File
        Put OutputFile, , Instruction
        Put OutputFile, , OldInfoHead
                    
<<<<<<< HEAD
        'Update
        If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
        DoEvents
    Wend
=======
                    'Update
                    If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
                    DoEvents
                Wend
>>>>>>> origin/master
                
    'Read everything?
    While ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles)

        'Create file
        Instruction = PatchInstruction.Create_File
        Put OutputFile, , Instruction
        Put OutputFile, , NewInfoHead
                    
        'Get file data
        Call Get_File_RawData(NewResourcePath, NewInfoHead, data, Modo)
        'Write data
        Put OutputFile, , data
                    
<<<<<<< HEAD
        'Update
        If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
        DoEvents
    Wend
=======
                    'Update
                    If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
                    DoEvents
                Wend
>>>>>>> origin/master
            
    'Close the patch file
    Close OutputFile
        
    'Close the new binary file
    Close NewResourceFile
    
    'Close the old binary file
    Close OldResourceFile
    
    Make_Patch = True
    Exit Function

ErrHandler:
    Close OutputFile
    Close NewResourceFile
    Close OldResourceFile
    
<<<<<<< HEAD
    Call MsgBox("No se pudo terminar de crear el parche. Raz�n: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")

    
    Exit Function

Make_Patch_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Make_Patch"
    End If
Resume Next
    
=======
    Call MsgBox(Replace$(JsonLanguage.Item("ERROR_CREAR_PARCHE").Item("TEXTO"), "VAR_ERROR", Err.number & " : " & Err.Description), vbOKOnly, JsonLanguage.Item("Error").Item("TEXTO"))
>>>>>>> origin/master
End Function

''
' Follows patches instructions to update a resource file.
'
' @param    ResourcePath The reource file folder.
' @param    PatchPath The patch file folder.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.
Public Function Apply_Patch(ByRef ResourcePath As String, _
                            ByRef PatchPath As String, _
                            ByRef prgBar As ProgressBar, _
                            Optional Modo As Byte = 0) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 17/07/2012 - ^[GS]^
    'Follows patches instructions to update a resource file
    '*****************************************************************
    
    On Error GoTo Apply_Patch_Err
    
    Dim ResourceFile       As Integer
    Dim ResourceFilePath   As String
    Dim FileHead           As FILEHEADER
    Dim InfoHead           As INFOHEADER
    Dim ResourceReadFiles  As Long
    Dim EOResource         As Boolean

    Dim PatchFile          As Integer
    Dim PatchFilePath      As String
    Dim PatchFileHead      As FILEHEADER
    Dim PatchInfoHead      As INFOHEADER
    Dim Instruction        As Byte
    Dim OldResourceVersion As Long

    Dim OutputFile         As Integer
    Dim OutputFilePath     As String
    Dim data()             As Byte
    Dim WrittenFiles       As Long
    Dim DataOutputPos      As Long

    On Local Error GoTo ErrHandler

    If Modo = 0 Then
        ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
        PatchFilePath = PatchPath & GRH_PATCH_FILE
        OutputFilePath = ResourcePath & GRH_RESOURCE_FILE & "tmp"
    ElseIf Modo = 1 Then
        ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE
        PatchFilePath = PatchPath & MAPS_PATCH_FILE
        OutputFilePath = ResourcePath & MAPS_RESOURCE_FILE & "tmp"

    End If
    
    'Open the old binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        
<<<<<<< HEAD
    'Read the old FileHeader
    Get ResourceFile, , FileHead

    'Check the file for validity
    If LOF(ResourceFile) <> FileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos anterior da�ado. " & ResourceFilePath, , "Error")
        Close ResourceFile
        Exit Function

    End If
=======
        'Read the old FileHeader
        Get ResourceFile, , FileHead
        'Check the file for validity
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            Call MsgBox(JsonLanguage.Item("ERROR_ARCHIVO_CORRUPTO").Item("TEXTO") & ": " & ResourceFilePath, , JsonLanguage.Item("Error").Item("TEXTO"))
            Close ResourceFile
            Exit Function
        End If
>>>>>>> origin/master
        
    'Open the patch file
    PatchFile = FreeFile()
    Open PatchFilePath For Binary Access Read Lock Write As PatchFile
            
    'Get previous file version
    Get PatchFile, , OldResourceVersion
            
<<<<<<< HEAD
    'Check the file version
    If OldResourceVersion <> FileHead.lngFileVersion Then
        Call MsgBox("Incongruencia en versiones.", , "Error")
        Close ResourceFile
        Close PatchFile
        Exit Function

    End If
=======
            'Check the file version
            If OldResourceVersion <> FileHead.lngFileVersion Then
                Call MsgBox(JsonLanguage.Item("ERROR_VERSIONES_RECURSOS").Item("TEXTO"), , JsonLanguage.Item("Error").Item("TEXTO"))
                Close ResourceFile
                Close PatchFile
                Exit Function
            End If
>>>>>>> origin/master
            
    'Read the new FileHeader
    Get PatchFile, , PatchFileHead
            
    'Destroy file if it previuosly existed
    If FileExist(OutputFilePath, vbNormal) Then Call Kill(OutputFilePath)
            
    'Open the patch file
    OutputFile = FreeFile()
    Open OutputFilePath For Binary Access Read Write As OutputFile
                
    'Save the file header
    Put OutputFile, , PatchFileHead
  
<<<<<<< HEAD
    If Not prgBar Is Nothing Then
        prgBar.value = 0
        prgBar.Max = PatchFileHead.lngNumFiles + 1

    End If
=======
                If Not prgBar Is Nothing Then
                    prgBar.value = 0
                    prgBar.Max = PatchFileHead.lngNumFiles + 1
                End If
>>>>>>> origin/master
                
    'Update
    DataOutputPos = Len(FileHead) + Len(InfoHead) * PatchFileHead.lngNumFiles + 1
                
    'Process loop
    While Loc(PatchFile) < LOF(PatchFile)
                    
        'Get the instruction
        Get PatchFile, , Instruction
        'Get the InfoHead
        Get PatchFile, , PatchInfoHead
                    
        Do
            EOResource = Not ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)
                        
            'Comparison is performed among encrypted names for ordering issues
            If Not EOResource And InfoHead.strFileName < PatchInfoHead.strFileName Then
                            
                'GetData and update InfoHead
                Call Get_File_RawData(ResourcePath, InfoHead, data, Modo)
                InfoHead.lngFileStart = DataOutputPos
                                           
                'Save file!
                Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
                Put OutputFile, DataOutputPos, data
                            
<<<<<<< HEAD
                'Update
                DataOutputPos = DataOutputPos + UBound(data) + 1
                WrittenFiles = WrittenFiles + 1

                If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
            Else
                Exit Do

            End If

        Loop
                    
        Select Case Instruction

                'Delete
            Case PatchInstruction.Delete_File

                If InfoHead.strFileName <> PatchInfoHead.strFileName Then
                    Err.Description = "Incongruencia en archivos de recurso"
                    GoTo ErrHandler

                End If
=======
                            'Update
                            DataOutputPos = DataOutputPos + UBound(data) + 1
                            WrittenFiles = WrittenFiles + 1
                            If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    Select Case Instruction
                        'Delete
                        Case PatchInstruction.Delete_File
                            If InfoHead.strFileName <> PatchInfoHead.strFileName Then
                                Err.Description = JsonLanguage.Item("ERROR_VERSIONES_RECURSOS").Item("TEXTO")
                                GoTo ErrHandler
                            End If
>>>>>>> origin/master
                        
                'Create
            Case PatchInstruction.Create_File

                If (InfoHead.strFileName > PatchInfoHead.strFileName) Or EOResource Then
                                
                    'Get file data
                    ReDim data(PatchInfoHead.lngFileSize - 1)
                    Get PatchFile, , data
                                
                    'Save it
                    Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                    Put OutputFile, DataOutputPos, data
                                
                    'Reanalize last Resource InfoHead
                    EOResource = False
                    ResourceReadFiles = ResourceReadFiles - 1
                                
<<<<<<< HEAD
                    'Update
                    DataOutputPos = DataOutputPos + UBound(data) + 1
                    WrittenFiles = WrittenFiles + 1

                    If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
                Else
                    Err.Description = "Incongruencia en archivos de recurso"
                    GoTo ErrHandler

                End If
=======
                                'Update
                                DataOutputPos = DataOutputPos + UBound(data) + 1
                                WrittenFiles = WrittenFiles + 1
                                If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
                            Else
                                Err.Description = JsonLanguage.Item("ERROR_VERSIONES_RECURSOS").Item("TEXTO")
                                GoTo ErrHandler
                            End If
>>>>>>> origin/master
                        
                'Modify
            Case PatchInstruction.Modify_File

                If InfoHead.strFileName = PatchInfoHead.strFileName Then

                    'Get file data
                    ReDim data(PatchInfoHead.lngFileSize - 1)
                    Get PatchFile, , data
                                             
                    'Save it
                    Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                    Put OutputFile, DataOutputPos, data
                                
<<<<<<< HEAD
                    'Update
                    DataOutputPos = DataOutputPos + UBound(data) + 1
                    WrittenFiles = WrittenFiles + 1

                    If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
                Else
                    Err.Description = "Incongruencia en archivos de recurso"
                    GoTo ErrHandler

                End If

        End Select
=======
                                'Update
                                DataOutputPos = DataOutputPos + UBound(data) + 1
                                WrittenFiles = WrittenFiles + 1
                                If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
                            Else
                                Err.Description = JsonLanguage.Item("ERROR_VERSIONES_RECURSOS").Item("TEXTO")
                                GoTo ErrHandler
                            End If
                    End Select
>>>>>>> origin/master
                    
        DoEvents
    Wend
                
    'Read everything?
    While ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)

        'GetData and update InfoHeader
        Call Get_File_RawData(ResourcePath, InfoHead, data, Modo)
        InfoHead.lngFileStart = DataOutputPos
                    
        'Save file!
        Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
        Put OutputFile, DataOutputPos, data
                    
<<<<<<< HEAD
        'Update
        DataOutputPos = DataOutputPos + UBound(data) + 1
        WrittenFiles = WrittenFiles + 1

        If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
        DoEvents
    Wend
=======
                    'Update
                    DataOutputPos = DataOutputPos + UBound(data) + 1
                    WrittenFiles = WrittenFiles + 1
                    If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
                    DoEvents
                Wend
>>>>>>> origin/master
            
    'Close the patch file
    Close OutputFile
        
    'Close the new binary file
    Close PatchFile
    
    'Close the old binary file
    Close ResourceFile
    
    'Check integrity
    If (PatchFileHead.lngNumFiles = WrittenFiles) Then

        'Replace File
        Call Kill(ResourceFilePath)
        Name OutputFilePath As ResourceFilePath

    Else
        Err.Description = JsonLanguage.Item("ERROR_LEER_PARCHE").Item("TEXTO")
        GoTo ErrHandler

    End If
    
    Apply_Patch = True
    Exit Function

ErrHandler:
    Close OutputFile
    Close PatchFile
    Close ResourceFile

    'Destroy file if created
    If FileExist(OutputFilePath, vbNormal) Then Call Kill(OutputFilePath)
    
<<<<<<< HEAD
    Call MsgBox("No se pudo parchear. Raz�n: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")

    
    Exit Function

Apply_Patch_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "Apply_Patch"
    End If
Resume Next
    
=======
    Call MsgBox(Replace$(JsonLanguage.Item("ERROR_CREAR_PARCHE").Item("TEXTO"), "VAR_ERROR", Err.number & " : " & Err.Description), vbOKOnly, JsonLanguage.Item("Error").Item("TEXTO"))
>>>>>>> origin/master
End Function

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
    '*****************************************************************
    'Author: Unknown
    'Last Modify Date: Unknown
    '*****************************************************************
    
    On Error GoTo AlignScan_Err
    
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8

    
    Exit Function

AlignScan_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "AlignScan"
    End If
Resume Next
    
End Function

''
' Retrieves the version number of a given resource file.
'
' @param    ResourceFilePath The resource file complete path.
'
' @return   The version number of the given file.

Public Function GetVersion(ByVal ResourceFilePath As String) As Long
    '*****************************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/23/2008
    '
    '*****************************************************************
    
    On Error GoTo GetVersion_Err
    
    Dim ResourceFile As Integer
    Dim FileHead     As FILEHEADER
    
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
    'Extract the FILEHEADER
    Get ResourceFile, 1, FileHead
        
    Close ResourceFile
    
    GetVersion = FileHead.lngFileVersion

    
    Exit Function

GetVersion_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "modCompression" & "->" & "GetVersion"
    End If
Resume Next
    
End Function

