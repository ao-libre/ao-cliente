VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private m_VBuffer  As Direct3DVertexBuffer8
Private m_IBuffer  As Direct3DIndexBuffer8
Private m_Texture  As Direct3DTexture8

Private Type TYPE_VERTEX
    X       As Single
    Y       As Single
    Z       As Single
    Colour  As Long
    tX      As Single
    tY      As Single
End Type


'
' Represent the dynamic data of the batch

Private m_DataAlpha() As Boolean
Private m_Data()   As TYPE_VERTEX
Private m_Position As Long
Private m_PrimitiveType As DxVBLibA.CONST_D3DPRIMITIVETYPE
 
Public Sub Begin()
    '
    ' Re-initialise the states of the batch
    '
    Call DirectDevice.SetTexture(0, m_Texture)
    Call DirectDevice.SetStreamSource(0, m_VBuffer, 24)
    Call DirectDevice.SetIndices(m_IBuffer, 0)
    
End Sub

Public Sub SetTexture(ByRef Texture As Direct3DTexture8)
    
    If ObjPtr(m_Texture) <> ObjPtr(Texture) Then 'te cabe vb6
        '
        ' Flush the batch since we're using a diferent texture
        '
        Call Flush
       
        Set m_Texture = Texture
       
        '
        ' Update the new texture of the batch
    
        Call DirectDevice.SetTexture(0, Texture)
    End If
End Sub
 
Public Sub Draw(ByVal X As Single, _
                ByVal Y As Single, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByRef Color() As Long, _
                Optional ByVal Tx1 As Single = 0, _
                Optional ByVal Ty1 As Single = 0, _
                Optional ByVal Tx2 As Single = 1, _
                Optional ByVal Ty2 As Single = 1, _
                Optional ByVal Alpha As Boolean = False, _
                Optional ByVal angle As Single = 0, _
                Optional ByRef PrimitiveType As DxVBLibA.CONST_D3DPRIMITIVETYPE = D3DPT_TRIANGLELIST)
                
    Dim dest As RECT
    dest.Left = X
    dest.Right = X + Width
    dest.Top = Y + Height
    dest.Bottom = Y
    
    If angle > 0 Then
        Dim x_center As Single
        Dim y_center As Single
        Dim radius As Single
        Dim x_Cor As Single
        Dim y_Cor As Single
        Dim left_point As Single
        Dim right_point As Single
        Dim temp As Single
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Bottom + (dest.Top - dest.Bottom) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.Top - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point

'***********************************************************************************************'
'***********************************************************************************************'
'***********************************************************************************************'
'***********************************************************************************************'

        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
        m_Data(m_Position).X = x_Cor
        m_Data(m_Position).Y = y_Cor
        m_Data(m_Position).Z = 0
        m_Data(m_Position).Colour = Color(0)
        m_Data(m_Position).tX = Tx1
        m_Data(m_Position).tY = Ty1
        m_DataAlpha(m_Position) = Alpha
        
        m_Position = m_Position + 1

        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
        m_Data(m_Position).X = x_Cor
        m_Data(m_Position).Y = y_Cor
        m_Data(m_Position).Z = 0
        m_Data(m_Position).Colour = Color(1)
        m_Data(m_Position).tX = Tx1
        m_Data(m_Position).tY = Ty2
        m_DataAlpha(m_Position) = Alpha
        
        m_Position = m_Position + 1

        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
        m_Data(m_Position).X = x_Cor
        m_Data(m_Position).Y = y_Cor
        m_Data(m_Position).Z = 0
        m_Data(m_Position).Colour = Color(2)
        m_Data(m_Position).tX = Tx2
        m_Data(m_Position).tY = Ty2
        m_DataAlpha(m_Position) = Alpha
        
        m_Position = m_Position + 1
        
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
        m_Data(m_Position).X = x_Cor
        m_Data(m_Position).Y = y_Cor
        m_Data(m_Position).Z = 0
        m_Data(m_Position).Colour = Color(3)
        m_Data(m_Position).tX = Tx2
        m_Data(m_Position).tY = Ty1
        m_DataAlpha(m_Position) = Alpha
        
        m_Position = m_Position + 1
'***********************************************************************************************'
'***********************************************************************************************'
'***********************************************************************************************'
'***********************************************************************************************'
'***********************************************************************************************'
    Else
        m_Data(m_Position).X = dest.Left
        m_Data(m_Position).Y = dest.Bottom
        m_Data(m_Position).Z = 0
        m_Data(m_Position).Colour = Color(0)
        m_Data(m_Position).tX = Tx1
        m_Data(m_Position).tY = Ty1
        m_DataAlpha(m_Position) = Alpha
        
        m_Position = m_Position + 1

        m_Data(m_Position).X = dest.Left
        m_Data(m_Position).Y = dest.Top
        m_Data(m_Position).Z = 0
        m_Data(m_Position).Colour = Color(1)
        m_Data(m_Position).tX = Tx1
        m_Data(m_Position).tY = Ty2
        m_DataAlpha(m_Position) = Alpha
        
        m_Position = m_Position + 1

        m_Data(m_Position).X = dest.Right
        m_Data(m_Position).Y = dest.Top
        m_Data(m_Position).Z = 0
        m_Data(m_Position).Colour = Color(2)
        m_Data(m_Position).tX = Tx2
        m_Data(m_Position).tY = Ty2
        m_DataAlpha(m_Position) = Alpha
        
        m_Position = m_Position + 1

        m_Data(m_Position).X = dest.Right
        m_Data(m_Position).Y = dest.Bottom
        m_Data(m_Position).Z = 0
        m_Data(m_Position).Colour = Color(3)
        m_Data(m_Position).tX = Tx2
        m_Data(m_Position).tY = Ty1
        m_DataAlpha(m_Position) = Alpha
        
        m_Position = m_Position + 1
        
        m_PrimitiveType = PrimitiveType
    End If
    
    If (m_Position = UBound(m_Data) + 1) Then
        Call Flush
    End If
    
End Sub

Public Sub Finish()
    '
    ' Flush the batch since we're finished
    '
    Call Flush
End Sub
 
Public Sub Flush()

    '
    ' Only draw if there is any data in the buffer
    '
    If m_Position > 0 Then
        
        '
        ' Set the new data of the vertex buffer
        '
        If m_DataAlpha((m_Position / 4) * 2) Then
            Call DirectDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_ONE)
            Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_ONE)
        End If
        
        Call D3DVertexBuffer8SetData(m_VBuffer, 0, m_Position * LenB(m_Data(0)), D3DLOCK_DISCARD, m_Data(0))
    
        Call DirectDevice.DrawIndexedPrimitive(m_PrimitiveType, 0, (m_Position / 4) * 6, 0, (m_Position / 4) * 2)
        
        If m_DataAlpha((m_Position / 4) * 2) Then
            Call DirectDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
            Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        End If
        
        '
        ' We start all over again
        '
        m_Position = 0
        
    End If
End Sub
 
Public Sub Initialise(ByVal Capacity As Long)
    ReDim m_Data(4 * Capacity) As TYPE_VERTEX
    ReDim m_DataAlpha(4 * Capacity) As Boolean
   
    '
    '  Create the vertice buffer
    '
    Set m_VBuffer = DirectDevice.CreateVertexBuffer(24 * Capacity, D3DUSAGE_DYNAMIC, D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1, D3DPOOL_DEFAULT)
  
    '
    '  Create the indice buffer, fill it with pre-baked indices
    '
    Set m_IBuffer = DirectDevice.CreateIndexBuffer(12 * Capacity, D3DUSAGE_WRITEONLY, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
   
    Dim lpIndices() As Integer
    
    ReDim lpIndices(0 To Capacity * 6 - 1) As Integer
    
    Dim i As Integer, J As Integer
   
    For i = 0 To UBound(lpIndices) Step 6
        lpIndices(i) = J
        lpIndices(i + 1) = J + 1
        lpIndices(i + 2) = J + 2
        lpIndices(i + 3) = J + 2
        lpIndices(i + 4) = J + 3
        lpIndices(i + 5) = J
       
        J = J + 4
    Next
    
    Call D3DIndexBuffer8SetData(m_IBuffer, 0, UBound(lpIndices), 0, lpIndices(0))
    
End Sub
