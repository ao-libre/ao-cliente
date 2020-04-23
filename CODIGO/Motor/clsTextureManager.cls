VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextureManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private Const BYTES_PER_MB          As Long = 1048576
Private Const DEFAULT_MEMORY_TO_USE As Long = 64 ' In MB
Private Const HASH_TABLE_SIZE       As Long = 337

Private Type SURFACE_ENTRY_DYN
    FileName As Integer
    UltimoAcceso As Long
    Texture As Direct3DTexture8
    Size As Long
    texture_width As Integer
    texture_height As Integer
End Type
 
Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode
Private DirectD3D                    As D3DX8
 
' Parra was here (;
Private mGraphicsNumber              As Long
Private mCurrentMemoryBytes          As Long
Private mMaxMemoryBytes              As Long
 
Private Sub Class_Terminate()

    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
    Dim i As Long
    Dim j As Long
   
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1

        With TexList(i)

            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Texture = Nothing
            Next j
           
            'Destroy the arrays
            Erase .SurfaceEntry

        End With

    Next i

End Sub
 
Public Property Get Surface(ByVal fileIndex As Long) As Direct3DTexture8

    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Used to return a texture
    '**************************************************************
    
End Property
 
Public Function GetTexture(ByVal FileName As Integer, _
                           ByRef TextureWidth As Long, _
                           ByRef TextureHeight As Long) As Direct3DTexture8

    If FileName = 0 Then
        Debug.Print "0 GRH ATTEMPT TO BE LOADED"
        Exit Function

    End If
 
    Dim i As Long

    ' Search the index on the list
    With TexList(FileName Mod HASH_TABLE_SIZE)

        For i = 1 To .surfaceCount

            If .SurfaceEntry(i).FileName = FileName Then
                TextureWidth = .SurfaceEntry(i).texture_width
                TextureHeight = .SurfaceEntry(i).texture_height
                Set GetTexture = .SurfaceEntry(i).Texture
                Exit Function

            End If

        Next i

    End With
 
    'Not in memory, load it!
    Set GetTexture = CrearGrafico(FileName, TextureWidth, TextureHeight)

End Function
 
Public Sub Initialize(ByRef DD As D3DX8, Optional ByVal maxMemoryUsageInMb As Long = DEFAULT_MEMORY_TO_USE)
    '**************************************************************
    'Author: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
   
    'Seteamos el objeto
    Set DirectD3D = DD
   
    ' Parra is back (;
    mGraphicsNumber = 0
    mCurrentMemoryBytes = 0
    mMaxMemoryBytes = maxMemoryUsageInMb * BYTES_PER_MB

End Sub

Private Function CrearGrafico(ByVal Archivo As Integer, _
                              ByRef TextureWidth As Long, _
                              ByRef TextureHeight As Long) As Direct3DTexture8

    On Error GoTo errhandler

    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
    Dim Index        As Integer
    Dim Buffer()     As Byte
    
    Index = Archivo Mod HASH_TABLE_SIZE

    With TexList(Index)
    
        .surfaceCount = .surfaceCount + 1
        
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN

        With .SurfaceEntry(.surfaceCount)
        
            'Nombre
            .FileName = Archivo

            'TODO: No usar variable de compilacion y acceder a esto desde el config.ini
            #If Compression = 1 Then
            
                'Obtenemos el .bmp del archivo en Graphics.AO
                Call modCompression.Get_Image(Game.path(Graficos), CStr(Archivo), Buffer())
                
                'Seteamos la textura.
                Set .Texture = DirectD3D.CreateTextureFromFileInMemoryEx(DirectDevice, Buffer(0), UBound(Buffer()) + 1, D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_A8R8G8B8, _
                                                                         D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
                Erase Buffer
                
            #Else
                
                Dim Graphic As String
                    Graphic = Game.path(Graficos) & .FileName
                
                If FileExist(Graphic & ".bmp", vbNormal) Then
                    Set .Texture = DirectD3D.CreateTextureFromFileEx(DirectDevice, Graphic & ".bmp", D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_A8R8G8B8, _
                                                                     D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
                
                ElseIf FileExist(Graphic & ".png", vbNormal) Then
                    Set .Texture = DirectD3D.CreateTextureFromFileEx(DirectDevice, Graphic & ".png", D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_A8R8G8B8, _
                                                                     D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
                
                Else
                    Call MsgBox("No se ha podido cargar el grafico con indice: " & Graphic)
                    Call CloseClient
                
                End If
                
            #End If
 
            Call .Texture.GetLevelDesc(0, surface_desc)
            
            .texture_width = texture_info.Width
            .texture_height = texture_info.Height
            .Size = surface_desc.Size
            
            TextureWidth = .texture_width
            TextureHeight = .texture_height
            
            Set CrearGrafico = .Texture

            ' Parra is back (;
            mGraphicsNumber = mGraphicsNumber + 1
            mCurrentMemoryBytes = mCurrentMemoryBytes + surface_desc.Size

        End With

    End With
    
    Exit Function
    
errhandler:
    Debug.Print "ERROR AL CARGAR EL GRAFICO> " & Archivo & ".bmp o .png"

End Function
