VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martin Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Private InventoryOffset As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem As Long             'Currently selected item
Private MouseDownSelectedItem As Long
Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer
Private prevSelItem As Long
Private ShowValue As Boolean
Private ValueFormated As Boolean

Private ImgContainer As Boolean
Private InventoryTickCount As Long

' Variables drag drop
Private TempItem As Integer
Private TempSlot As Integer

Private MouseX As Integer
Private MouseY As Integer
Private DragItem As Byte
Private Drag As Byte

Private DrawBuffer As cDIBSection ' Para guardar temporalmente una "foto" del inventario (WyroX)

Private isInventory As Boolean ' Para separar al inventario del main, de los otros

Private Inventory() As Inventory    'User's inventory

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************

    Set DrawBuffer = Nothing

End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Public Sub ScrollInventory(ByVal Up As Boolean)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 14/01/2010
'Scrolls the graphical inventory up / down
'14/01/2010: ZaMa - Now redraws inventory only if needed
'***************************************************

    'Only move if there is something to show up /down
    If CanScroll(Up) Then
    
        InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
        
        'Redraw the inventory if needed
        Call DrawInventory
        
    End If

End Sub

Public Function CanScroll(ByVal Up As Boolean)

    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 14/01/2010
    'Determines wether inventory can be scrolled up / down
    '***************************************************
    
    If Up Then
        CanScroll = (InventoryOffset + (InventoryWindow.ScaleHeight \ ownTilePixelHeight)) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) < UBound(Inventory)
    
    Else
        CanScroll = InventoryOffset > 0

    End If
    
End Function

Public Sub SelectGold()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
    'Redraw old item to deselect it
    If prevSelectedItem <> FLAGORO Then
        'Call DrawInventorySlot(prevSelectedItem)
    End If
    
    Call DrawInventory
End Sub

Public Sub DeselectItem()
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Deselect the currently selected item
'***************************************************
    If InvSelectedItem = 0 Then Exit Sub
    
    Dim ClearSlot As Byte
        ClearSlot = InvSelectedItem
    
    'Select nothing
    InvSelectedItem = 0
    
    'Redraw old item to deselect it
    Call DrawInventory

End Sub

Public Sub ChangeSlotItemAmount(ByVal slot As Byte, ByVal Amount As Long)
'***************************************************
'Author: ZaMa
'Created Date: 30/11/2009
'Last Modify Date: 09/12/2009 (Budi)
'Change the amount of the required item, and update the slot.
'***************************************************
    
    Inventory(slot).Amount = Amount
    
    'Update Amount
    Call DrawInventory
    
End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = Inventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = Inventory(slot).MinHit
End Property

Public Property Get MaxDef(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    MaxDef = Inventory(slot).MaxDef
End Property

Public Property Get MinDef(ByVal slot As Byte) As Integer
'***************************************************
'Author: Budi
'Last Modify Date: 02/1209
'Retrieves the defense of the selected item
'***************************************************
    MinDef = Inventory(slot).MinDef
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = Inventory(slot).GrhIndex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = Inventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(Inventory) And slot <= UBound(Inventory) Then
        Amount = Inventory(slot).Amount
    End If
End Property

Public Property Get ObjIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    ObjIndex = Inventory(slot).ObjIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = Inventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = Inventory(slot).name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = Inventory(slot).Equipped
End Property

Public Property Get MaxObjs() As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Retrieves the capacity of the Inventory
'***************************************************
    MaxObjs = UBound(Inventory)
End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Set the capacity of the Inventary
'***************************************************
    If UBound(Inventory) = MaxObjs Then Exit Sub
    
    ReDim Preserve Inventory(1 To MaxObjs) As Inventory
End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 11/12
'Indicates if the amount of the items in the inventory is shown
'***************************************************
    ShowValue = bShow
    
    Dim slot As Integer
    ' Update Inventory
        Call DrawInventory
End Sub

Public Sub ClearAllSlots()
'***************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Clear All Slots.
'***************************************************

    Dim slot As Long
    For slot = 1 To Me.MaxObjs
        Call SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString)
    Next slot

End Sub

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Long, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Long, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal MaxDef As Integer, ByVal MinDef As Integer, ByVal eValor As Single, ByVal eName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(slot)
        .Amount = eAmount
        .MaxDef = MaxDef
        .MinDef = MinDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .name = eName
        .ObjIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
    End With
    
    'If InventoryWindow.Visible = False Then InventoryWindow.Visible = True
    
    'Render inventory slot (we don't need to render the whole inventory)
    Call DrawInventory
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Selects the item clicked if it's valid and return's it's index
    '***************************************************
    Dim TempItem As Long
    Dim temp_x   As Long
    Dim temp_y   As Long
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    TempSlot = TempItem

    'Make sure it's within limits
    If TempItem <= UBound(Inventory) Then

        'Make sure slot isn't empty
        If Inventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If

    End If

End Function

Public Sub DrawInventory()

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 18/01/2020
    'Renders the inventory to the given PictureBox.
    '***************************************************
    Dim LoopC        As Long
    Dim X            As Integer
    Dim Y            As Integer
    Dim DestRect     As RECT
    Dim Inv_Color(3) As Long
    Dim Sel_Color(3) As Long

    With DestRect
        .Bottom = InventoryWindow.Height
        .Right = InventoryWindow.Width
    End With
    
    'Color al seleccionar un item.
    Call Engine_Long_To_RGB_List(Sel_Color(), D3DColorARGB(100, 255, 128, 0))
    
    'Clear the inventory window
    Call Engine_BeginScene

    For LoopC = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(Inventory)
            
        startdX = ((LoopC - 1) Mod (frmMain.PicInv.Width / 32)) * 32
        startdY = ((LoopC - 1) \ (frmMain.PicInv.Width / 32)) * 32

        Call Draw_GrhIndex(26095, startdX, startdY, 0, Normal_RGBList(), 0, False) 'Dibuja el fondo
            
        If Inventory(LoopC).GrhIndex Then

            ' Draw Item
            If LoopC <> DragItem Then
                    
                If LoopC = SelectedItem Then
                    Call Draw_GrhIndex(26096, startdX, startdY, 0, Normal_RGBList(), 0, False) 'Dibuja el fondo cuando lo seleccionamos

                End If

                Call Draw_GrhIndex(Inventory(LoopC).GrhIndex, startdX, startdY, 0, Normal_RGBList())

                'If equipped we render "E"
                If Inventory(LoopC).Equipped Then
                    Call DrawText(startdX + 20, startdY + 20, "E", -1)

                End If
                    
                Call DrawText(startdX, startdY, Inventory(LoopC).Amount, -1)

            End If

        End If

    Next LoopC

    If DragItem <> 0 And Drag = 1 And isInventory Then
        Call Draw_GrhIndex(Inventory(DragItem).GrhIndex, MouseX - 16, MouseY - 16, 0, Normal_RGBList())
        
        ' Si estamos arrastrando un item, dibujamos normal
        InventoryWindow.AutoRedraw = False
        Call Engine_EndScene(DestRect, InventoryWindow.hWnd)
        
    Else
        ' WyroX:
        ' Brujeria necesaria para que no quede negro el inventario.
        ' En realidad: con el AutoRedraw desactivado, copiamos una "foto" del inventario a un buffer en memoria.
        ' Activamos el AutoDraw, lo copiamos de vuelta al inventario y lo dejamos como imagen estatica.

        InventoryWindow.AutoRedraw = False

        Call Engine_EndScene(DestRect, InventoryWindow.hWnd)

        Call DrawBuffer.LoadPictureBlt(InventoryWindow.hdc)

        InventoryWindow.AutoRedraw = True

        Call DrawBuffer.PaintPicture(InventoryWindow.hdc, 0, 0, InventoryWindow.Width, InventoryWindow.Height, 0, 0, vbSrcCopy)

        InventoryWindow.Picture = InventoryWindow.Image

    End If

End Sub

Public Sub DrawDragAndDrop()

    If ((DragItem <> 0 And Drag = 1) And (timeGetTime - InventoryTickCount > 40)) Then

        Call Inventario.DrawInventory

    End If

End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 13/09/19
    'Implements the mouse move event of the inventory picture box
    'Displays a ToolTip of the item under the mouse
    'Now we show more information about certain items (Recox)
    '***************************************************
    Dim temp_x   As Integer
    Dim temp_y   As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's areanv
    If X < 16 Then X = 16
    If Y < 16 Then Y = 16
    If X > InventoryWindow.Width - 16 Then X = InventoryWindow.Width - 16
    If Y > InventoryWindow.Height - 16 Then Y = InventoryWindow.Height - 16
        
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    TempSlot = TempItem
        
    'Incorporado por ReyarB, arreglado por Recox
    If TempItem <= UBound(Inventory) Then

        Select Case Inventory(TempItem).OBJType
            Case eObjType.otWeapon, eObjType.otFlechas
                InventoryWindow.ToolTipText = Inventory(TempItem).name & ", Ataque: " & Inventory(TempItem).MaxHit & "/" & Inventory(TempItem).MinHit
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                InventoryWindow.ToolTipText = Inventory(TempItem).name & ", Defensa: " & Inventory(TempItem).MaxDef & "/" & Inventory(TempItem).MinDef
            Case Else
                InventoryWindow.ToolTipText = Inventory(TempItem).name
        End Select

    End If

    MouseX = CInt(X)
    MouseY = CInt(Y)

End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton And isInventory Then
        If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub

        Dim temp_x As Integer
        Dim temp_y As Integer

        temp_x = X \ 32
        temp_y = Y \ 32

        'Get the currently clickced item
        DragItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1

        If Not (DragItem <= 25 And DragItem >= 1) Then
            DragItem = 0
            Exit Sub
        Else

            If Inventory(DragItem).GrhIndex = 0 Then
                DragItem = 0
                Exit Sub

            End If

        End If

        If isInventory Then
            Drag = 1
            
        End If

    End If

End Sub

Public Sub Initialize(ByRef DirectD3D As D3DX8, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 32, Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, Optional ByVal startY As Integer = 0, Optional ByVal bImgContainer As Boolean = True, Optional ByVal bShowText As Boolean = True, Optional thisInventory As Boolean = False)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 03/12/09
    'Sets the reference to the picture box on which to render the inventory
    '03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
    '***************************************************
    
    Set InventoryWindow = InvPic
    isInventory = thisInventory
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    ' Inicializar buffer
    Set DrawBuffer = New cDIBSection
    Call DrawBuffer.Create(InvPic.Width, InvPic.Height)
    
    ReDim Inventory(1 To MaxObjs) As Inventory
    
    'Make sure DirectD3D was correctly initialized
    If DirectD3D Is Nothing Then Exit Sub
    
    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    
    ImgContainer = bImgContainer
    ShowValue = bShowText
    
    If ImgContainer Then InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item

End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Implements the mouse up event of the inventory picture box
    'Check outs which item was clicked
    '***************************************************
    
    'Store previously selected item
    Dim prevSelItem As Long

    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(MouseX, MouseY)

    If Drag Then

        If Not DragItem = TempSlot And Not (TempSlot < 1 Or DragItem < 1) Then
            Call Protocol.WriteMoveItem(DragItem, TempSlot, eMoveType.Inventory)
            InvSelectedItem = TempSlot
        End If

        DragItem = 0
        Drag = 0

    End If

    Call DrawInventory

End Sub

Public Function GetPos(ByVal X As Integer, ByVal Y As Integer) As Byte
    
    Dim temp_x As Integer
    Dim temp_y As Integer
    
    temp_x = X \ 32
    temp_y = Y \ 32
    
    GetPos = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1

End Function

