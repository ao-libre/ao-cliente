VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
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

Private InventoryOffset            As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem            As Long             'Currently selected item

Private ownTilePixelWidth          As Integer
Private ownTilePixelHeight         As Integer
Private startdX                    As Integer
Private startdY                    As Integer

Private ShowValue                  As Boolean
Private ValueFormated              As Boolean

Private ImgContainer               As Boolean
Private InventoryTickCount         As Long
Private slotDragged                As Integer

Private dragAndDropOn              As Boolean

Private Inventory()                As Inventory    'User's inventory

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

'Public Event dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If

Private Sub Class_Initialize()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    '
    '***************************************************
End Sub

Private Sub Class_Terminate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    '
    '***************************************************
End Sub

Public Sub ScrollInventory(ByVal Up As Boolean)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 14/01/2010
    'Scrolls the graphical inventory up / down
    '14/01/2010: ZaMa - Now redraws inventory only if needed
    '***************************************************
    
    On Error GoTo ScrollInventory_Err
    

    'Only move if there is something to show up /down
    If CanScroll(Up) Then
        InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
        
        'Redraw the inventory if needed
        'DrawInventory
    End If

    
    Exit Sub

ScrollInventory_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "ScrollInventory"
    End If
Resume Next
    
End Sub

Public Function CanScroll(ByVal Up As Boolean)
    
    On Error GoTo CanScroll_Err
    

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
    
    
    Exit Function

CanScroll_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "CanScroll"
    End If
Resume Next
    
End Function

Public Sub SelectGold()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Sets the gold as the currently selected item
    '***************************************************
    
    On Error GoTo SelectGold_Err
    
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
    'Redraw old item to deselect it
    If prevSelectedItem <> FLAGORO Then

        'Call DrawInventorySlot(prevSelectedItem)
    End If

    
    Exit Sub

SelectGold_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "SelectGold"
    End If
Resume Next
    
End Sub

Public Sub DeselectItem()
    
    On Error GoTo DeselectItem_Err
    

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
    'Call DrawInventorySlot(ClearSlot)

    
    Exit Sub

DeselectItem_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "DeselectItem"
    End If
Resume Next
    
End Sub

Public Sub ChangeSlotItemAmount(ByVal slot As Byte, ByVal Amount As Long)
    '***************************************************
    'Author: ZaMa
    'Created Date: 30/11/2009
    'Last Modify Date: 09/12/2009 (Budi)
    'Change the amount of the required item, and update the slot.
    '***************************************************
    
    On Error GoTo ChangeSlotItemAmount_Err
    
    
    Inventory(slot).Amount = Amount
    
    'Update Amount
    'Call DrawInventorySlot(slot)

    
    Exit Sub

ChangeSlotItemAmount_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "ChangeSlotItemAmount"
    End If
Resume Next
    
End Sub

Public Property Get SelectedItem() As Long
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item index
    '***************************************************
    
    On Error GoTo SelectedItem_Err
    
    SelectedItem = InvSelectedItem

    
    Exit Property

SelectedItem_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "SelectedItem"
    End If
Resume Next
    
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the max hit of the selected item
    '***************************************************
    
    On Error GoTo MaxHit_Err
    
    MaxHit = Inventory(slot).MaxHit

    
    Exit Property

MaxHit_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "MaxHit"
    End If
Resume Next
    
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the min hit of the selected item
    '***************************************************
    
    On Error GoTo MinHit_Err
    
    MinHit = Inventory(slot).MinHit

    
    Exit Property

MinHit_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "MinHit"
    End If
Resume Next
    
End Property

Public Property Get MaxDef(ByVal slot As Byte) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the defense of the selected item
    '***************************************************
    
    On Error GoTo MaxDef_Err
    
    MaxDef = Inventory(slot).MaxDef

    
    Exit Property

MaxDef_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "MaxDef"
    End If
Resume Next
    
End Property

Public Property Get MinDef(ByVal slot As Byte) As Integer
    '***************************************************
    'Author: Budi
    'Last Modify Date: 02/1209
    'Retrieves the defense of the selected item
    '***************************************************
    
    On Error GoTo MinDef_Err
    
    MinDef = Inventory(slot).MinDef

    
    Exit Property

MinDef_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "MinDef"
    End If
Resume Next
    
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the grh index of the selected item
    '***************************************************
    
    On Error GoTo GrhIndex_Err
    
    GrhIndex = Inventory(slot).GrhIndex

    
    Exit Property

GrhIndex_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "GrhIndex"
    End If
Resume Next
    
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the value of the selected item
    '***************************************************
    
    On Error GoTo Valor_Err
    
    Valor = Inventory(slot).Valor

    
    Exit Property

Valor_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "Valor"
    End If
Resume Next
    
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
    
    On Error GoTo Amount_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's amount
    '***************************************************
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(Inventory) And slot <= UBound(Inventory) Then
        Amount = Inventory(slot).Amount

    End If

    
    Exit Property

Amount_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "Amount"
    End If
Resume Next
    
End Property

Public Property Get ObjIndex(ByVal slot As Byte) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's object index
    '***************************************************
    
    On Error GoTo ObjIndex_Err
    
    ObjIndex = Inventory(slot).ObjIndex

    
    Exit Property

ObjIndex_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "ObjIndex"
    End If
Resume Next
    
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's object type
    '***************************************************
    
    On Error GoTo OBJType_Err
    
    OBJType = Inventory(slot).OBJType

    
    Exit Property

OBJType_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "OBJType"
    End If
Resume Next
    
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's name
    '***************************************************
    
    On Error GoTo ItemName_Err
    
    ItemName = Inventory(slot).Name

    
    Exit Property

ItemName_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "ItemName"
    End If
Resume Next
    
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves True if the item at the given pos is eqiupped
    '***************************************************
    
    On Error GoTo Equipped_Err
    
    Equipped = Inventory(slot).Equipped

    
    Exit Property

Equipped_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "Equipped"
    End If
Resume Next
    
End Property

Public Property Get MaxObjs() As Byte
    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modify Date: 09/16/09
    'Retrieves the capacity of the Inventory
    '***************************************************
    
    On Error GoTo MaxObjs_Err
    
    MaxObjs = UBound(Inventory)

    
    Exit Property

MaxObjs_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "MaxObjs"
    End If
Resume Next
    
End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Byte)
    
    On Error GoTo SetMaxObjs_Err
    

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modify Date: 09/16/09
    'Set the capacity of the Inventary
    '***************************************************
    If UBound(Inventory) = MaxObjs Then Exit Sub
    
    ReDim Preserve Inventory(1 To MaxObjs) As Inventory

    
    Exit Sub

SetMaxObjs_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "SetMaxObjs"
    End If
Resume Next
    
End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)
    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 11/12
    'Indicates if the amount of the items in the inventory is shown
    '***************************************************
    
    On Error GoTo ToggleShowValue_Err
    
    ShowValue = bShow
    
    Dim slot As Integer

    ' Update Inventory
    'For slot = 1 To Me.MaxObjs
    'DrawInventorySlot (slot)
    'Next slot
    
    Exit Sub

ToggleShowValue_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "ToggleShowValue"
    End If
Resume Next
    
End Sub

Public Sub ClearAllSlots()
    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Clear All Slots.
    '***************************************************
    
    On Error GoTo ClearAllSlots_Err
    

    Dim slot As Long

    For slot = 1 To Me.MaxObjs
        Call SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString)
    Next slot

    
    Exit Sub

ClearAllSlots_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "ClearAllSlots"
    End If
Resume Next
    
End Sub

Public Sub SetItem(ByVal slot As Byte, _
                   ByVal eOBJIndex As Integer, _
                   ByVal eAmount As Long, _
                   ByVal eEquipped As Byte, _
                   ByVal eGrhIndex As Integer, _
                   ByVal eObjType As Integer, _
                   ByVal eMaxHit As Integer, _
                   ByVal eMinHit As Integer, _
                   ByVal MaxDef As Integer, _
                   ByVal MinDef As Integer, _
                   ByVal eValor As Single, _
                   ByVal eName As String)
    
    On Error GoTo SetItem_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
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
        .Name = eName
        .ObjIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor

    End With
    
    'If InventoryWindow.Visible = False Then InventoryWindow.Visible = True
    
    'Render inventory slot (we don't need to render the whole inventory)
    'Call DrawInventorySlot(slot)
    
    Exit Sub

SetItem_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "SetItem"
    End If
Resume Next
    
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Selects the item clicked if it's valid and return's it's index
    '***************************************************
    
    On Error GoTo ClickItem_Err
    
    Dim TempItem As Long
    Dim temp_x   As Long
    Dim temp_y   As Long
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    'Make sure it's within limits
    If TempItem <= UBound(Inventory) Then

        'Make sure slot isn't empty
        If Inventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0

        End If

    End If

    
    Exit Function

ClickItem_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "ClickItem"
    End If
Resume Next
    
End Function

Private Sub DrawInventory()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Renders the inventory to the given PictureBox
    '***************************************************
    
    On Error GoTo DrawInventory_Err
    
    Dim LoopC             As Long

    Dim destRect          As RECT
    Dim Inv_Color(0 To 3) As Long
    Dim Sel_Color(0 To 3) As Long
    
    With destRect
        .bottom = InventoryWindow.Height
        .Right = InventoryWindow.Width

    End With
    
    'Clear the inventory window
    
    Call Engine_BeginScene
    
    For LoopC = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(Inventory)
            
        startdX = ((LoopC - 1) Mod (frmMain.PicInv.Width / 32)) * 32
        startdY = ((LoopC - 1) \ (frmMain.PicInv.Width / 32)) * 32
            
        Call Engine_Long_To_RGB_List(Inv_Color(), -1)
            
        Call Engine_Long_To_RGB_List(Sel_Color(), D3DColorARGB(100, 255, 128, 0))
            
        'Call DDrawTransGrhIndextoSurface(24664, startdX, startdY, 0, Inv_Color(), 0, False)
            
        If Inventory(LoopC).GrhIndex Then
                
            Call DDrawTransGrhIndextoSurface(Inventory(LoopC).GrhIndex, startdX, startdY, 0, Inv_Color())
                
            If LoopC = SelectedItem Then
                Call DDrawTransGrhIndextoSurface(2, startdX, startdY, 0, Sel_Color())

            End If
            
            'If equipped we render "E"
            If Inventory(LoopC).Equipped Then
                'Fonts_Render_String "E", startdX + 20, startdY + 20, -1, 2
                DrawText startdX + 20, startdY + 20, "E", -1

            End If
                
            'Render the item grh and the amount
            'Fonts_Render_String Inventory(LoopC).Amount, startdX, startdY, -1, 2
            DrawText startdX, startdY, Inventory(LoopC).Amount, -1

        End If

    Next LoopC
        
    Call Engine_EndScene(destRect, InventoryWindow.hwnd)
    
    
    Exit Sub

DrawInventory_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "DrawInventory"
    End If
Resume Next
    
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Implements the mouse move event of the inventory picture box
    'Displays a ToolTip of the item under the mouse
    '***************************************************
    
    On Error GoTo InventoryWindow_MouseMove_Err
    
    Dim temp_x   As Integer
    Dim temp_y   As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's areanv
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
        
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    If TempItem <= UBound(Inventory) Then
        InventoryWindow.ToolTipText = Inventory(TempItem).Name

    End If

    
    Exit Sub

InventoryWindow_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "InventoryWindow_MouseMove"
    End If
Resume Next
    
End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    
    On Error GoTo InventoryWindow_MouseDown_Err
    
    Dim temp_x   As Integer
    Dim temp_y   As Integer
    Dim TempItem As Integer
    
    If dragAndDropOn Then

        'Exit if it got outside the control's area
        If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
        
        If Button = vbRightButton Then
            temp_x = X \ ownTilePixelWidth
            temp_y = Y \ ownTilePixelHeight
        
            TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
            
            slotDragged = TempItem

        End If

    End If

    
    Exit Sub

InventoryWindow_MouseDown_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "InventoryWindow_MouseDown"
    End If
Resume Next
    
End Sub

Public Sub DrawInv(Optional ByVal r As Boolean = False)
    
    On Error GoTo DrawInv_Err
    

    If GetTickCount - InventoryTickCount > 100 Then
        Call DrawInventory
        InventoryTickCount = GetTickCount

    End If

    
    Exit Sub

DrawInv_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "DrawInv"
    End If
Resume Next
    
End Sub

Public Sub Initialize(ByRef DirectD3D As D3DX8, _
                      ByRef InvPic As PictureBox, _
                      ByVal MaxObjs As Byte, _
                      Optional ByVal FontSize As Integer = 7, _
                      Optional ByVal TileWidth As Integer = 32, _
                      Optional ByVal TileHeight As Integer = 32, _
                      Optional ByVal startX As Integer = 0, _
                      Optional ByVal startY As Integer = 0, _
                      Optional ByVal bImgContainer As Boolean = True, _
                      Optional ByVal bShowText As Boolean = True)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 03/12/09
    'Sets the reference to the picture box on which to render the inventory
    '03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
    '***************************************************
    
    On Error GoTo Initialize_Err
    
    Set InventoryWindow = InvPic
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
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

    
    Exit Sub

Initialize_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "Initialize"
    End If
Resume Next
    
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Implements the mouse up event of the inventory picture box
    'Check outs which item was clicked
    '***************************************************
    'Store previously selected item
    
    On Error GoTo InventoryWindow_MouseUp_Err
    
    Dim prevSelItem As Long
    Dim TempItem    As Long
    Dim temp_x      As Long
    Dim temp_y      As Long
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
    
    prevSelItem = InvSelectedItem
    
    If Button = vbRightButton Then
        If dragAndDropOn Then
            temp_x = X \ ownTilePixelWidth
            temp_y = Y \ ownTilePixelHeight
        
            InvSelectedItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
            
            Call DrawInventory
            
            If InvSelectedItem <> slotDragged Then

                'RaiseEvent dragDone(slotDragged, InvSelectedItem)
            End If

        End If

    Else
        'Get the currently clickced item
        InvSelectedItem = ClickItem(CInt(X), CInt(Y))
    
        'Update needed inventory slots
        If prevSelItem <> InvSelectedItem Then

            'If prevSelItem <> 0 And prevSelItem <> FLAGORO Then Call DrawInventorySlot(prevSelItem)
            'If InvSelectedItem Then Call DrawInventorySlot(InvSelectedItem)
        End If

    End If

    
    Exit Sub

InventoryWindow_MouseUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGrapchicalInventory" & "->" & "InventoryWindow_MouseUp"
    End If
Resume Next
    
End Sub
