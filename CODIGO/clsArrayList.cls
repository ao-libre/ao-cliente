VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************
'Author: FrankoH298
'Count ahora se llama GetLastItem (Recox)
' Funciones:
'           - Clear() - Limpia la lista
'           - itemExist(item) - Te indica si el item existe
'           - itemPos(item) - Te devuelve la posicion del item
'           - Item(index) - Te devuelve el item en la posicion
'           - Add() - Agrega un item a la lista
'           - Remove() - Remueve el item de la lista
'           - GetLastItem() - Te muestra el ultimo item de la lista
'***************************************************

Option Explicit

Private ArrayList() As Byte

Public Sub Clear()

'***************************************************
'Author: FrankoH298
'Last Modify Date: 07/09/2019
' Limpia la lista
'***************************************************

Dim i As Long
    For i = LBound(ArrayList) To UBound(ArrayList)
        ArrayList(i) = 0
    Next i
End Sub

Public Sub Initialize(ByVal min As Integer, ByVal max As Integer)

'***************************************************
'Author: FrankoH298
'Last Modify Date: 07/09/2019
' Inicializa la clase
'***************************************************

ReDim ArrayList(min To max) As Byte
Me.Clear
End Sub

Public Function itemExist(ByVal item As Integer) As Boolean

'***************************************************
'Author: FrankoH298
'Last Modify Date: 08/09/2019
' Devuelve si existe el item
'***************************************************

Dim i As Long
    For i = LBound(ArrayList) To UBound(ArrayList)
        If ArrayList(i) = item Then
            itemExist = True
            Exit Function
        End If
    Next i
    itemExist = False
End Function

Public Function itemPos(ByVal item As Integer) As Byte

'***************************************************
'Author: FrankoH298
'Last Modify Date: 08/09/2019
' Devuelve la posicion del item
'***************************************************

Dim i As Long
    For i = LBound(ArrayList) To UBound(ArrayList)
        If ArrayList(i) = item Then
            itemPos = i
            Exit Function
        End If
    Next i
    itemPos = 0
End Function

Public Function item(ByVal Index As Integer) As Byte

'***************************************************
'Author: FrankoH298
'Last Modify Date: 08/09/2019
' Devuelve que item esta en esa posicion
'***************************************************

item = ArrayList(Index)
End Function

Public Sub Add(ByVal item As Integer)

'***************************************************
'Author: FrankoH298
'Last Modify Date: 07/09/2019
' Agrega el item a la lista
'***************************************************

Dim i As Long
    For i = LBound(ArrayList) To UBound(ArrayList)
        If ArrayList(i) = 0 Then
            ArrayList(i) = item
            Exit Sub
        End If
    Next i
End Sub

Public Sub Remove(ByVal item As Integer)

'***************************************************
'Author: FrankoH298
'Last Modify Date: 07/09/2019
' Remueve el item de la lista
'***************************************************

Dim i As Long
Dim a As Long
    For i = LBound(ArrayList) To (UBound(ArrayList) - 1)
        If ArrayList(i) = item Then
            For a = i To (UBound(ArrayList) - 1)
                ArrayList(a) = ArrayList(a + 1)
            Next a
            ArrayList(UBound(ArrayList)) = 0
            Exit Sub
        End If
    Next i
    If ArrayList(UBound(ArrayList)) = item Then ArrayList(UBound(ArrayList)) = 0
End Sub

Public Function GetLastItem() As Byte

'***************************************************
'Author: FrankoH298
'Last Modify Date: 08/09/2019
' Devuelve el ultimo objeto de la lista que sea diferente a 0
'***************************************************

Dim i As Long
    For i = UBound(ArrayList) To LBound(ArrayList) Step -1
        If ArrayList(i) <> 0 Then
            GetLastItem = ArrayList(i)
            Exit Function
        End If
    Next i
    GetLastItem = 0
End Function

