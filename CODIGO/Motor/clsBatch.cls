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
 
Private m_VBuffer As Direct3DVertexBuffer8
Private m_IBuffer As Direct3DIndexBuffer8
Private m_Texture As Direct3DTexture8
Private m_Alpha   As Boolean
 
Private Type TYPE_VERTEX
           
    X       As Single
    Y       As Single
    Z       As Single
 
    Colour  As Long
   
    TX      As Single
    TY      As Single
 
End Type
 
'
' Represent the dynamic data of the batch
'
Private m_Data()   As TYPE_VERTEX
Private m_Position As Long
 
Public Sub Begin()
    '
    ' Re-initialise the states of the batch
    '
    Call DirectDevice.SetTexture(0, m_Texture)
    Call DirectDevice.SetStreamSource(0, m_VBuffer, 24)
    Call DirectDevice.SetIndices(m_IBuffer, 0)
   
    If m_Alpha Then
        Call DirectDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_ONE)
        Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_ONE)
    Else
        Call DirectDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    End If
 
End Sub
 
Public Sub SetAlpha(ByVal IsAlpha As Boolean)
 
    If IsAlpha <> m_Alpha Then
        
        '
        ' Flush the batch since we're using a diferent state
        '
        Call Flush
 
        m_Alpha = IsAlpha
       
        '
        ' Update the new state of the batch
        '
        If IsAlpha Then
            Call DirectDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_ONE)
            Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_ONE)
        Else
            Call DirectDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
            Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        End If
 
    End If
 
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
        '
        Call DirectDevice.SetTexture(0, Texture)
 
    End If
 
End Sub
 
Public Sub Draw(ByVal X As Long, _
                ByVal Y As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByRef Color() As Long, _
                Optional ByVal Tx1 As Single = 0, _
                Optional ByVal Ty1 As Single = 0, _
                Optional ByVal Tx2 As Single = 1, _
                Optional ByVal Ty2 As Single = 1, _
                Optional ByVal angle As Single = 0)
               
    If angle > 0 And angle <= 360 Then
        Dim NewX     As Single
        Dim NewY     As Single
        Dim SinRad   As Single
        Dim CosRad   As Single
        Dim RadAngle As Single
        Dim CenterX  As Single
        Dim CenterY  As Single
       
        RadAngle = angle * DegreeToRadian
   
        CenterX = X + (Width * 0.5)
        CenterY = Y + (Height * 0.5)
       
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)
       
    End If
 
    m_Data(m_Position).X = X
    m_Data(m_Position).Y = Y
    m_Data(m_Position).Z = 0
    m_Data(m_Position).Colour = Color(0)
    m_Data(m_Position).TX = Tx1
    m_Data(m_Position).TY = Ty1
   
    If angle > 0 And angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).X - CenterX) * -CosRad - (m_Data(m_Position).Y - CenterY) * -SinRad
        NewY = CenterY + (m_Data(m_Position).Y - CenterY) * -CosRad + (m_Data(m_Position).X - CenterX) * -SinRad
       
        m_Data(m_Position).X = NewX
        m_Data(m_Position).Y = NewY
       
    End If
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).X = X
    m_Data(m_Position).Y = Y + Height
    m_Data(m_Position).Z = 0
    m_Data(m_Position).Colour = Color(1)
    m_Data(m_Position).TX = Tx1
    m_Data(m_Position).TY = Ty2
 
    If angle > 0 And angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).X - CenterX) * -CosRad - (m_Data(m_Position).Y - CenterY) * -SinRad
        NewY = CenterY + (m_Data(m_Position).Y - CenterY) * -CosRad + (m_Data(m_Position).X - CenterX) * -SinRad
       
        m_Data(m_Position).X = NewX
        m_Data(m_Position).Y = NewY
       
    End If
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).X = X + Width
    m_Data(m_Position).Y = Y + Height
    m_Data(m_Position).Z = 0
    m_Data(m_Position).Colour = Color(2)
    m_Data(m_Position).TX = Tx2
    m_Data(m_Position).TY = Ty2
 
    If angle > 0 And angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).X - CenterX) * -CosRad - (m_Data(m_Position).Y - CenterY) * -SinRad
        NewY = CenterY + (m_Data(m_Position).Y - CenterY) * -CosRad + (m_Data(m_Position).X - CenterX) * -SinRad
       
        m_Data(m_Position).X = NewX
        m_Data(m_Position).Y = NewY
       
    End If
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).X = X + Width
    m_Data(m_Position).Y = Y
    m_Data(m_Position).Z = 0
    m_Data(m_Position).Colour = Color(3)
    m_Data(m_Position).TX = Tx2
    m_Data(m_Position).TY = Ty1
 
    If angle > 0 And angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).X - CenterX) * -CosRad - (m_Data(m_Position).Y - CenterY) * -SinRad
        NewY = CenterY + (m_Data(m_Position).Y - CenterY) * -CosRad + (m_Data(m_Position).X - CenterX) * -SinRad
       
        m_Data(m_Position).X = NewX
        m_Data(m_Position).Y = NewY
       
    End If
   
    m_Position = m_Position + 1
   
    If (m_Position >= ((UBound(m_Data) / 4))) Then
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
        Call D3DVertexBuffer8SetData(m_VBuffer, 0, m_Position * LenB(m_Data(0)), D3DLOCK_DISCARD, m_Data(0))
   
        Call DirectDevice.DrawIndexedPrimitive(D3DPT_TRIANGLELIST, 0, (m_Position / 4) * 6, 0, (m_Position / 4) * 2)
 
        '
        ' We start all over again
        '
        m_Position = 0
 
    End If
 
End Sub
 
Public Sub Initialise(ByVal Capacity As Long)

    ReDim m_Data(4 * Capacity) As TYPE_VERTEX
   
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
   
    Dim i As Long, J As Integer
   
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

