VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBmpMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=============================================================================
'
'   NOMBRE: CBmpMan.cls
'   AUTOR: Federico Torres
'   FECHA: 11/01/03
'   CONTACTO: fede_jt@hotmail.com
'
'   DESCRIPCION:
'       Clase para manejar los graficos, evita tener que cargarlos todos en memoria
'
'
'
'
'       Primero hay que setear los miembros de la clase. MaxEntries indica la cantidad maxima de graficos a cargar.
'       lpDirectDraw7 es un puntero a un objeto de tipo DirectDraw7, se usa para cargar las surface.
'       Con los valores asignados hay que ejecutar el metodo Init().
'       Getbmp devuelve un objeto tipo directdrawsurface7, si el grafico que se pide no esta cargado se carga.
'       Al llegar al limite de graficos permitidos, al pedir el proximo grafico que no este cargado se borra el que menos
'       se uso (de los que ya estaban cargados).
'
'
'   REVISION por Federico Torres:
'       12/01/03: Agregado color key.
'
'   REVISION por Gulfas Morgolock:
'       04/04/03: Agregué el array mKeys(). Este vector
'       mapea un filename a su respectiva posición en el
'       vector de surfaces mSurfaceDB(). Si mKey(Index) es
'       cero el filename de valor index no está en memoria.
'       BMPMan originalmente hacia una busqueda lineal para
'       obtener un surface, con esta técnica el tiempo de
'       acceso es constante!
'
'   REVISION por Gulfas Morgolock:
'       01/04/03: Arreglé un bug en BorrarMenosUsado().
'       No se reseteaba la referencia al bmp.
'
'
'
'   FALTA:
'       -Comprobar errores.
'       -Elegir si se quiere colorkey y que color.
'
'
'=============================================================================

Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer 'borrar de aca para abajo?
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type BMP_DB_ENTRY
    FileName As Integer
    Height As Long
    Width As Long
    Size As Long
    Accesos As Long
    Cont As Long
    Efecto As Long
End Type

Private Type BMP_INFO
    Height As Long
    Width As Long
    Size As Long
End Type


Private mBMPDB() As BMP_DB_ENTRY
Private mDDraw As DirectDraw7
Private mSurfaceDB() As DirectDrawSurface7
Private mKeys() As Integer

Private mMaxEntries As Long
Private mContador As Long
Private mCantidadGraficos As Long

Private mUsarMemoriaVideo  As Boolean

Private mEfectoPred As Integer

Private GrhPath As String

Public Property Let Path(ByVal inString As String)
GrhPath = inString
End Property

Public Function GetBMP(ByVal FileName As Integer, Optional ByVal Efecto As Integer = -1) As DirectDrawSurface7

Dim RetVal As Integer

If Efecto < 0 Then
    Efecto = mEfectoPred
End If

If mKeys(Efecto, FileName) <> 0 Then 'esta el surface en memoria?
    mBMPDB(mKeys(Efecto, FileName)).Accesos = mBMPDB(mKeys(Efecto, FileName)).Accesos + 1 'Sumo un acceso
    Set GetBMP = mSurfaceDB(mKeys(Efecto, FileName)) 'y devuelvo una surface con el grafico cargado
Else 'Grafico no cargado
    RetVal = BuscaVacio() 'busco espacio para cargarlo
    If RetVal = -1 Then 'todas las surface ocupadas
        RetVal = 1
        RetVal = BorraMenosUsado() 'borro el grafico menos usado
    End If
    Call CrearGrafico(FileName, RetVal) 'y creo la textura
    mBMPDB(RetVal).Accesos = mBMPDB(RetVal).Accesos + 1 'Sumo un acceso
    mBMPDB(RetVal).Efecto = Efecto
    mKeys(Efecto, FileName) = RetVal
    
#If ConAlfaB = 1 Then
    If Efecto = 1 Then
        EfectoNoche mSurfaceDB(RetVal)
    End If
#End If

    Set GetBMP = mSurfaceDB(RetVal) 'y devuelvo una surface con el grafico cargado
End If


End Function


Public Function Init(ByVal UsarMemoriaVideo As Boolean) As Long   'Inicializa todo, ejecutar antes de usar el objeto
On Error GoTo ErrHandler

If mMaxEntries < 1 Then 'por lo menos 1 grafico
    Init = -1
    Exit Function
End If

ReDim mBMPDB(0 To mMaxEntries)
ReDim mSurfaceDB(0 To mMaxEntries)
ReDim mKeys(0 To 1, 0 To 15000) As Integer

mEfectoPred = 0
mUsarMemoriaVideo = UsarMemoriaVideo
mCantidadGraficos = 0
mContador = 0
Init = 1

Exit Function

ErrHandler:
    Init = -1
End Function

Public Function BorrarTodo() As Long
'Vacia la lista de graficos
Dim i As Integer
For i = 1 To mMaxEntries
    If mBMPDB(i).FileName > 0 Then
        Call Borrar(mBMPDB(i).FileName, mBMPDB(i).Efecto)
    End If
Next
BorrarTodo = 1
mCantidadGraficos = 0
End Function

Public Function Borrar(ByVal FileName As Integer, ByVal Efecto As Integer) As Integer
'Borra un grafico
Borrar = -1
Call ZeroMemory(mBMPDB(mKeys(Efecto, FileName)), Len(mBMPDB(mKeys(Efecto, FileName))))
Set mSurfaceDB(mKeys(Efecto, FileName)) = Nothing
mKeys(Efecto, FileName) = 0
Borrar = 1
mCantidadGraficos = mCantidadGraficos - 1
End Function

Private Function CrearGrafico(ByVal Archivo As Integer, ByVal bIndex As Integer) As Integer
'12/1/3, colorkey
On Error Resume Next
Dim ddsd As DDSURFACEDESC2, BMPInfo As BMP_INFO, RetVal As Long, ddck As DDCOLORKEY
ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH

Call GetBMPHeader(GrhPath & Archivo & ".bmp", BMPInfo)  'para alto y ancho de la surface

If mUsarMemoriaVideo Then
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Else
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
End If
ddsd.lHeight = BMPInfo.Height
ddsd.lWidth = BMPInfo.Width

ddck.high = 0: ddck.low = 0

With mBMPDB(bIndex)
    .FileName = Archivo
    .Accesos = 0
    .Height = BMPInfo.Height
    .Width = BMPInfo.Width
    .Size = BMPInfo.Size
    .Cont = mContador
End With
mContador = mContador + 1

Set mSurfaceDB(bIndex) = mDDraw.CreateSurfaceFromFile(GrhPath & Archivo & ".bmp", ddsd) 'creo la surface
mSurfaceDB(bIndex).SetColorKey DDCKEY_SRCBLT, ddck

CrearGrafico = 1
mCantidadGraficos = mCantidadGraficos + 1

End Function

Private Sub Class_Initialize()
mMaxEntries = 0
mUsarMemoriaVideo = True
mEfectoPred = 0
End Sub



Private Function BuscaVacio() As Integer
'Si esta todo lleno devuelve -1 (llamar BorraMenosUsado)
Dim i As Integer
For i = 1 To mMaxEntries
    If (mSurfaceDB(i) Is Nothing) Then
        BuscaVacio = i
        Exit Function
    End If
Next

BuscaVacio = -1
End Function

Private Function BorraMenosUsado() As Integer
'Borra el grafico menos usado
'Devuelve el index al borrado

BorraMenosUsado = 1

Dim lValor As Long
Dim lCont As Long

lValor = mBMPDB(1).Accesos
lCont = mBMPDB(1).Cont

Dim i As Integer

For i = 1 To mMaxEntries
    If mBMPDB(i).Accesos < lValor And mBMPDB(i).Cont < lCont Then
        lValor = mBMPDB(i).Accesos
        lCont = mBMPDB(i).Cont
        BorraMenosUsado = i
    End If
Next
mCantidadGraficos = mCantidadGraficos - 1

mKeys(mBMPDB(BorraMenosUsado).Efecto, mBMPDB(BorraMenosUsado).FileName) = 0
Set mSurfaceDB(BorraMenosUsado) = Nothing
Call ZeroMemory(mBMPDB(BorraMenosUsado), Len(mBMPDB(BorraMenosUsado)))

End Function

Private Sub Class_Terminate()
On Error Resume Next 'no me importa, estoy cerrando

'Erase mBMPDB()
Dim i As Integer

For i = 1 To mMaxEntries
    Set mSurfaceDB(i) = Nothing
Next

End Sub

Public Property Let lpDirectDraw7(ByRef objDirectDraw7 As DirectDraw7)
Set mDDraw = objDirectDraw7
End Property

Private Function GetBMPHeader(Archivo As String, ByRef lpBMPInfo As BMP_INFO) As Long
'if fileexist
Dim nfile As Integer, H_BMPH As BITMAPFILEHEADER, H_BMPI As BITMAPINFOHEADER
nfile = FreeFile()
Open Archivo For Binary Access Read As #nfile
Get #nfile, , H_BMPH
Get #nfile, , H_BMPI
Close #nfile
lpBMPInfo.Height = H_BMPI.biHeight
lpBMPInfo.Width = H_BMPI.biWidth
lpBMPInfo.Size = H_BMPH.bfSize
GetBMPHeader = 1
End Function

Public Property Get MaxEntries() As Long
MaxEntries = mMaxEntries
End Property

Public Property Let MaxEntries(ByVal vNewValue As Long)
mMaxEntries = vNewValue
End Property

Public Property Get CantidadGraficos() As Variant
CantidadGraficos = mCantidadGraficos
End Property

'*******************************************************
'*******************************************************
'*******************************************************
'*******************************************************
'*******************************************************

'Para compatibilidad con el BmpMan No Dyn
Public Property Get TotalGraficos() As Variant

End Property

Public Property Let TotalGraficos(ByVal vNewValue As Variant)

End Property

Public Property Get EsDinamico() As Variant
EsDinamico = True
End Property

Public Function CargarGrafico(Archivo As Integer) As Integer

End Function

Public Property Get EfectoPred() As Integer
EfectoPred = mEfectoPred
End Property

Public Property Let EfectoPred(ByVal iNewValue As Integer)
mEfectoPred = iNewValue
End Property
