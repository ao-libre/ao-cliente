VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManDyn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsSurfaceManDyn.cls - Inherits from clsSurfaceManager. Is designed to load
'surfaces dynamically without using more than an arbitrary amount of Mb.
'For removale it uses LRU, attempting to just keep in memory those surfaces
'that are actually usefull.
'
' Developed by Maraxus (Juan Martín Sotuyo Dodero - juansotuyo@hotmail.com)
' Last Modify Date: 3/06/2006
'**************************************************************

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

'Inherit from the surface manager
Implements clsSurfaceManager

Private Const BYTES_PER_MB As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE As Long = 4 * BYTES_PER_MB          '4 Mb
Private Const DEFAULT_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB     '16 Mb

Private Type SURFACE_ENTRY_DYN
    fileIndex As Long
    lastAccess As Long
    Surface As DirectDrawSurface7
End Type

Private surfaceList() As SURFACE_ENTRY_DYN
Private surfaceCount As Long

Private DirectDraw As DirectDraw7

Private maxBytesToUse As Long
Private usedBytes As Long

Private useVideoMemory As Boolean

Private GrhPath As String

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef source As Any, ByVal byteCount As Long)

Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'
'**************************************************************
    usedBytes = 0
    surfaceCount = 0
    ReDim surfaceList(0) As SURFACE_ENTRY_DYN
    maxBytesToUse = MIN_MEMORY_TO_USE
End Sub

Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Clean up
'**************************************************************
    Dim i  As Long
    
    'Destroy every surface in memory
    For i = 0 To surfaceCount - 1
        Set surfaceList(i).Surface = Nothing
    Next i
    
    'Destroy the array
    Erase surfaceList
End Sub

Private Sub clsSurfaceManager_Initialize(ByRef DD As DirectDraw7, ByVal videoMemory As Boolean, ByVal graphicPath As String, Optional ByVal maxMemoryUsageInMb As Long = -1)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Initializes the manager
'**************************************************************
    Set DirectDraw = DD
    
    useVideoMemory = videoMemory
    
    GrhPath = graphicPath
    
    If maxMemoryUsageInMb = -1 Then
        maxBytesToUse = DEFAULT_MEMORY_TO_USE   ' 16 Mb by default
    ElseIf maxMemoryUsageInMb * BYTES_PER_MB < MIN_MEMORY_TO_USE Then
        maxBytesToUse = MIN_MEMORY_TO_USE       ' 4 Mb is the minimum allowed
    Else
        maxBytesToUse = maxMemoryUsageInMb * BYTES_PER_MB
    End If
End Sub

Private Property Get clsSurfaceManager_Surface(ByVal fileIndex As Long) As DirectDrawSurface7
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Retrieves the requested texture
'**************************************************************
    Dim index As Long
    
    ' Search the index on the list
    index = BinarySearch(fileIndex)
    
    If index < 0 Then
        'Not found, we have to load the file and add it in the position given by the negation of the index
        index = Not index
        
        Call LoadSurface(fileIndex, index)
    End If
    
    'Return it
    With surfaceList(index)
        .lastAccess = GetTickCount
        Set clsSurfaceManager_Surface = .Surface
    End With
End Property

Private Function BinarySearch(ByVal fileIndex As Long) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Returns the index of the surface in the list, or the negation
'of the position were it should be if not found (for binary insertion)
'**************************************************************
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = surfaceCount - 1
    
    Do While min <= max
        mid = (min + max) \ 2
    
        If surfaceList(mid).fileIndex < fileIndex Then
            min = mid + 1
        ElseIf surfaceList(mid).fileIndex > fileIndex Then
            max = mid - 1
        Else
            'We found it
            BinarySearch = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    BinarySearch = Not mid
End Function

Private Sub LoadSurface(ByVal fileIndex As Long, ByVal listIndex As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Loads the surface named fileIndex + ".bmp" and inserts it to the
'surface list in the listIndex position
'**************************************************************
On Error GoTo ErrHandler

    Dim newSurface As SURFACE_ENTRY_DYN
    Dim ddsd As DDSURFACEDESC2
    Dim ddck As DDCOLORKEY
    Dim filePath As String
    
    'Store complete file path
    filePath = GrhPath & CStr(fileIndex) & ".bmp"
    
    'Set up the surface desc
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    
    If useVideoMemory Then
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    
    Call surfaceDimensions(filePath, ddsd.lHeight, ddsd.lWidth)
    
    With newSurface
        .fileIndex = fileIndex
        
        'Load surface
        Set .Surface = DirectDraw.CreateSurfaceFromFile(filePath, ddsd)
        
        'Set colorkey
        ddck.high = 0
        ddck.low = 0
        Call .Surface.SetColorKey(DDCKEY_SRCBLT, ddck)
        
        'Retrieve the updated surface desc
        Call .Surface.GetSurfaceDesc(ddsd)
    End With
    
    'Insert surface to the list at the given pos
    Call InsertSurface(newSurface, listIndex)
    
    'Update used bytes
    usedBytes = usedBytes + ddsd.lHeight * ddsd.lPitch
    
    'Check if we have exceeded our allowed share of memory usage
    While usedBytes > maxBytesToUse And surfaceCount > 0
        Call RemoveLRU
    Wend
Exit Sub

ErrHandler:
    If Err.Number = DDERR_OUTOFMEMORY Or Err.Number = DDERR_OUTOFVIDEOMEMORY Then
        If surfaceCount Then
            'Remove a surface and try again
            Call RemoveLRU
            Resume Next
        Else
            MsgBox "No hay memoria disponible! El programa abortará. Cierra algunos programas e intenta de nuevo"
            End
        End If
    Else
        MsgBox "Un error inesperado ocurrió al intentar cargar el gráfico " & filePath & ". " & vbCrLf & _
                "El código de error es " & CStr(Err.Number) & vbCrLf & "Copia este mensaje y notifica a los administradores.", _
                vbOKOnly Or vbCritical Or vbExclamation, "Error"
        End
    End If
End Sub

Private Sub surfaceDimensions(ByVal Archivo As String, ByRef Height As Long, ByRef Width As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Loads the headers of a bmp file to retrieve it's dimensions at rt
'**************************************************************
    Dim handle As Integer
    Dim bmpFileHead As BITMAPFILEHEADER
    Dim bmpInfoHead As BITMAPINFOHEADER
    
    handle = FreeFile()
    Open Archivo For Binary Access Read Lock Write As handle
        Get handle, , bmpFileHead
        Get handle, , bmpInfoHead
    Close handle
    
    Height = bmpInfoHead.biHeight
    Width = bmpInfoHead.biWidth
End Sub

Private Sub InsertSurface(ByRef Surface As SURFACE_ENTRY_DYN, ByVal listIndex As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Inserts the given surface in the requested position of the surface list
'**************************************************************
On Error GoTo ErrHandler
    Dim i As Long
    
    'Enlarge the list
    ReDim Preserve surfaceList(surfaceCount) As SURFACE_ENTRY_DYN
    
    'Move back the list
    For i = surfaceCount - 1 To listIndex Step -1
        surfaceList(i + 1) = surfaceList(i)
    Next i
    
    'Insert the new surface
    surfaceList(listIndex) = Surface
    
    'Increase surface count
    surfaceCount = surfaceCount + 1
Exit Sub

ErrHandler:
    MsgBox "Un error irreparable ocurrió al insertar un nuevo gráfico en la lista." & vbCrLf _
            & "El cliente se cerrará" & vbCrLf _
            & "Intente usar el cliente no dinámico"
    End
End Sub

Private Sub RemoveLRU()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Removes the Least Recently Used surface to make some room for new ones
'**************************************************************
    Dim LRU As Long
    Dim i As Long
    Dim ddsd As DDSURFACEDESC2
    
    'Should never happen, but just in case....
    If surfaceCount = 0 Then Exit Sub
    
    'Initialize with the first element of the list
    LRU = 0
    
    'Check out through the whole list for the least recently used
    For i = 0 To surfaceCount - 1
        If surfaceList(LRU).lastAccess > surfaceList(i).lastAccess Then
            LRU = i
        End If
    Next i
    
    'Retrieve the surface desc
    Call surfaceList(LRU).Surface.GetSurfaceDesc(ddsd)
    
    'Remove it
    Set surfaceList(LRU).Surface = Nothing
    
    'Move back the list (if necessary)
    If LRU < surfaceCount - 1 Then
        CopyMemory surfaceList(LRU), surfaceList(LRU + 1), (surfaceCount - LRU - 1) * LenB(surfaceList(LRU))
    End If
    
    'Decrease surface count
    surfaceCount = surfaceCount - 1
    
    'Update the used bytes
    usedBytes = usedBytes - ddsd.lHeight * ddsd.lPitch
End Sub
