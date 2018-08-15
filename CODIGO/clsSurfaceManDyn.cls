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

Private Const BYTES_PER_MB As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE As Long = 4 * BYTES_PER_MB          '4 Mb
Private Const DEFAULT_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB     '16 Mb

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 337

Private Type SURFACE_ENTRY_DYN
    fileIndex As Long
    lastAccess As Long
    Surface As Direct3DTexture8
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashNode

Private DirectD3D As D3DX8

Private maxBytesToUse As Long
Private usedBytes As Long

Private ResourcePath As String

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const DIB_PAL_COLORS As Long = 1
Private Const DIB_RGB_COLORS As Long = 0
Private Declare Function SetDIBitsToDevice Lib "GDI32.dll" _
    (ByVal hdc As Long, ByVal XDest As Long, ByVal YDest As Long, _
    ByVal dwWidth As Long, ByVal dwHeight As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal uStartScan As Long, ByVal CScanLine As Long, _
    lpBits As Any, lpbmi As BITMAPINFO, ByVal fuColorUse As Long) As Long

Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'
'**************************************************************
    usedBytes = 0
    maxBytesToUse = MIN_MEMORY_TO_USE
End Sub

Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Clean up
'**************************************************************
    Dim i As Long
    Dim J As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For J = 1 To .surfaceCount
                Set .SurfaceEntry(J).Surface = Nothing
            Next J
            
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i
End Sub

Private Sub clsSurfaceManager_Initialize(ByRef DD As D3DX8, ByVal graphicPath As String, Optional ByVal maxMemoryUsageInMb As Long = -1)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 29/07/2012 - ^[GS]^
'Initializes the manager
'**************************************************************
    Set DirectD3D = DD
    
    ResourcePath = graphicPath
    
    If maxMemoryUsageInMb = -1 Then
        maxBytesToUse = DEFAULT_MEMORY_TO_USE   ' 16 Mb by default
    ElseIf maxMemoryUsageInMb * BYTES_PER_MB < MIN_MEMORY_TO_USE Then
        maxBytesToUse = MIN_MEMORY_TO_USE       ' 4 Mb is the minimum allowed
    Else
        maxBytesToUse = maxMemoryUsageInMb * BYTES_PER_MB
    End If
End Sub

Private Property Get clsSurfaceManager_Surface(ByVal fileIndex As Long) As Direct3DTexture8
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Retrieves the requested texture
'**************************************************************
    Dim i As Long
    
    ' Search the index on the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).fileIndex = fileIndex Then
                .SurfaceEntry(i).lastAccess = GetTickCount
                Set clsSurfaceManager_Surface = .SurfaceEntry(i).Surface
                Exit Property
            End If
        Next i
    End With
    
    'Not in memory, load it!
    Set clsSurfaceManager_Surface = LoadSurface(fileIndex)
End Property

Private Function LoadSurface(ByVal fileIndex As Long) As Direct3DTexture8
'**************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 09/10/2012 - ^[GS]^
'Loads the surface named fileIndex + ".bmp" and inserts it to the
'surface list in the listIndex position
'**************************************************************
On Error GoTo ErrHandler

    Dim newSurface As SURFACE_ENTRY_DYN
    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
    Dim data() As Byte
    
    'get Bitmap
    Call Get_Image(ResourcePath, CStr(fileIndex), data)
    With newSurface
        .fileIndex = fileIndex
        
        'Set last access time (if we didn't we would reckon this texture as the one lru)
        .lastAccess = GetTickCount
        
        Set .Surface = DirectD3D.CreateTextureFromFileInMemoryEx( _
                DirectDevice, data(0), UBound(data) + 1, _
                D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)

        newSurface.Surface.GetLevelDesc 0, surface_desc
        
    End With
    
    'Insert surface to the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        .surfaceCount = .surfaceCount + 1
        
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        
        .SurfaceEntry(.surfaceCount) = newSurface
        
        Set LoadSurface = newSurface.Surface
    End With
    
    'Update used bytes
    usedBytes = usedBytes + surface_desc.Size
    
    'Check if we have exceeded our allowed share of memory usage
    Do While usedBytes > maxBytesToUse
        'Remove a file. If no file could be removed we continue, if the file was previous to our surface we update the index
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function

ErrHandler:

End Function

Private Function RemoveLRU() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Removes the Least Recently Used surface to make some room for new ones
'**************************************************************
    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim J As Long
    Dim ddsd As D3DSURFACE_DESC
    
    LRUtime = GetTickCount
    
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For J = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(J).lastAccess Then
                    LRUi = i
                    LRUj = J
                    LRUtime = .SurfaceEntry(J).lastAccess
                End If
            Next J
        End With
    Next i
    
    If LRUj Then
        RemoveLRU = True
        'Retrieve the surface desc
        Call surfaceList(LRUi).SurfaceEntry(LRUj).Surface.GetLevelDesc(0, ddsd)
        
        'Remove it
        Set surfaceList(LRUi).SurfaceEntry(LRUj).Surface = Nothing
        surfaceList(LRUi).SurfaceEntry(LRUj).fileIndex = 0
        
        'Move back the list (if necessary)
        With surfaceList(LRUi)
            For J = LRUj To .surfaceCount - 1
                .SurfaceEntry(J) = .SurfaceEntry(J + 1)
            Next J
            
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
        
        'Update the used bytes
        usedBytes = usedBytes - ddsd.Size
    End If
End Function



