VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If AntiExternos Then
 
Public Redundance As Byte
 
Public Sub NAC_E_Byte(ByRef ByteArray() As Byte, ByVal code As Byte)

    Dim i As Integer

    For i = 0 To UBound(ByteArray)
        ByteArray(i) = code Xor ByteArray(i)
    Next

End Sub
 
Public Sub NAC_D_Byte(ByRef ByteArray() As Byte, ByVal code As Byte)

    Dim i As Integer

    For i = 0 To UBound(ByteArray)
        ByteArray(i) = ByteArray(i) Xor code
    Next

End Sub
 
Public Function NAC_E_String(ByVal t As String, ByVal code As Byte) As String

    Dim Bytes() As Byte

    Bytes = StrConv(t, vbFromUnicode)
    Call NAC_E_Byte(Bytes, code)
    NAC_E_String = StrConv(Bytes, vbUnicode)
End Function
 
Public Function NAC_D_String(ByVal t As String, ByVal code As Byte) As String

    Dim Bytes() As Byte

    Bytes = StrConv(t, vbFromUnicode)
    Call NAC_D_Byte(Bytes, code)
    NAC_D_String = StrConv(Bytes, vbUnicode)
End Function
 
#End If



