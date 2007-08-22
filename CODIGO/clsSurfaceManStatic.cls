VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsSurfaceManStatic.cls - Inherits from clsSurfaceManager. Is designed to load
' surfaces at startup, and never deallocating them.
' This grants high performance can use a lot of RAM.
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
'Affero General Public License for more details.
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

'Size to which we resize the list when we start loading textures to prevent ReDim Preserve on each add
'Once done the list is trimmed to the proper size if it's larger than needed.
'A ReDim Preserve is executed for each surface after DEFAULT_LIST_SIZE + 1
Private Const DEFAULT_LIST_SIZE As Integer = 1500

Private Type SURFACE_ENTRY_STATIC
    fileIndex As Long
    Surface As DirectDrawSurface7
End Type

Private surfaceList() As SURFACE_ENTRY_STATIC
Private surfaceCount As Long

Private DirectDraw As DirectDraw7

Private useVideoMemory As Boolean

Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'
'**************************************************************
    surfaceCount = 0
    ReDim surfaceList(0) As SURFACE_ENTRY_STATIC
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

Private Sub clsSurfaceManager_Initialize(ByRef DD As DirectDraw7, ByVal videoMemory As Boolean, ByVal graphicPath As String, Optional ByVal maxMemoryUsageInMb As Long = -1&)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'
'**************************************************************
    Set DirectDraw = DD
    
    useVideoMemory = videoMemory
    
    Call LoadSurfaces(graphicPath)
End Sub

Private Property Get clsSurfaceManager_Surface(ByVal fileIndex As Long) As DirectDrawSurface7
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Retrieves the requested texture
'**************************************************************
On Error GoTo ErrHandler:
    Dim index As Long
    
    ' Search the index on the list
    index = BinarySearch(fileIndex)
    
    'Return it
    Set clsSurfaceManager_Surface = surfaceList(index).Surface
Exit Property

ErrHandler:
    If index < 0 Then
        MsgBox "No se encuentra el archivo " & CStr(fileIndex) & ".bmp. Reinstale el juego, " _
                & "y si el problema persiste contactese con los adminsitradores", vbOKOnly Or vbCritical Or vbExclamation, "Error"
    Else
        MsgBox "Un error inesperado ocurrió a intentar cargar el archivo " & CStr(fileIndex) & ".bmp. & vbcrlf" _
                & "Error: " & CStr(Err.Number), vbOKOnly Or vbCritical Or vbExclamation, "Error"
    End If
    
    End
End Property

Private Sub LoadSurfaces(ByVal GrhPath As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Loads all surfaces in random order and then sorts them
'**************************************************************
On Error Resume Next
    Dim FileName As String
    Dim surfaceIndex As Long
    
    'Set up the list at a standard value big enough to prevent using ReDim Preserve constantly (which is slow)
    ReDim surfaceList(DEFAULT_LIST_SIZE) As SURFACE_ENTRY_STATIC
    
    FileName = Dir$(GrhPath & "*.bmp", vbArchive)
    
    While FileName <> ""
        'Get the surface index (numeric part of the number) - NEVER use Val() here or the error won't be raised!!! - Don't use IsNumeric or hexa strings will be accepted
        surfaceIndex = CLng(Left$(FileName, Len(FileName) - 4))
        
        If Err.Number = 13 Then
            'Type mysmatch - the name of the file isn't numneric, therefore it isn't a surface
            Err.Clear
        Else
            'Increase surface count and resize list if needed
            surfaceCount = surfaceCount + 1
            If surfaceCount > DEFAULT_LIST_SIZE + 1 Then
                ReDim Preserve surfaceList(surfaceCount - 1) As SURFACE_ENTRY_STATIC
            End If
            
            Call LoadSurface(GrhPath, surfaceIndex, surfaceCount - 1)
        End If
        
        'Get next .bmp file
        FileName = Dir$()
    Wend
    
    'Trim the list if needed
    If surfaceCount <> UBound(surfaceList) + 1 Then
        ReDim Preserve surfaceList(surfaceCount - 1) As SURFACE_ENTRY_STATIC
    End If
    
    'Sort the list
    Call SortSurfaces(0, surfaceCount - 1)
End Sub

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
    BinarySearch = Not min
End Function

Private Sub LoadSurface(ByVal GrhPath As String, ByVal fileIndex As Long, ByVal listIndex As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Loads the surface named fileIndex + ".bmp" and inserts it to the
'surface list in the listIndex position
'**************************************************************
On Error GoTo ErrHandler

    Dim newSurface As SURFACE_ENTRY_STATIC
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
    End With
    
    'Store the surface in the given index (it MUST be empty or data will be lost)
    surfaceList(listIndex) = newSurface
Exit Sub

ErrHandler:
    MsgBox "Un error inesperado ocurrió al intentar cargar el gráfico " & filePath & ". " & vbCrLf & _
            "El código de error es " & CStr(Err.Number) & " - " & Err.Description & vbCrLf & "Copia este mensaje y notifica a los administradores.", _
            vbOKOnly Or vbCritical Or vbExclamation, "Error"
        End
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

Private Sub SortSurfaces(ByVal first As Integer, ByVal last As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Sorts the list using quicksort, this allows the use of BinarySearch for faster searches
'**************************************************************
    Dim min As Integer      'First item in the list
    Dim max As Integer      'Last item in the list
    Dim comp As Long        'Item used to compare
    Dim temp As SURFACE_ENTRY_STATIC
    
    min = first
    max = last
    
    comp = surfaceList((min + max) \ 2).fileIndex
    
    Do While min <= max
        Do While surfaceList(min).fileIndex < comp And min < last
            min = min + 1
        Loop
        Do While surfaceList(max).fileIndex > comp And max > first
            max = max - 1
        Loop
        If min <= max Then
            temp = surfaceList(min)
            surfaceList(min) = surfaceList(max)
            surfaceList(max) = temp
            min = min + 1
            max = max - 1
        End If
    Loop
    
    If first < max Then SortSurfaces first, max
    If min < last Then SortSurfaces min, last
End Sub
