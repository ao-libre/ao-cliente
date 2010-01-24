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

Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost
Dim InvSelectedItem As Long             'Currently selected item

Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything

Dim UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory    'User's inventory

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

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

Public Sub ScrollInventory(ByVal up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Scrolls the graphical inventory up / down
'***************************************************
    'Only move if there is something to show up /down
    If up Then
        If (InventoryOffset + (InventoryWindow.ScaleHeight \ TilePixelHeight)) * (InventoryWindow.ScaleWidth \ TilePixelWidth) < MAX_INVENTORY_SLOTS Then _
            InventoryOffset = InventoryOffset + 1
    Else
        If InventoryOffset > 0 Then _
            InventoryOffset = InventoryOffset - 1
    End If
    
    'Redraw the inventory
    DrawInventory
End Sub

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
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
        Call DrawInventorySlot(prevSelectedItem)
    End If
End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = UserInventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = UserInventory(slot).MinHit
End Property

Public Property Get Def(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    Def = UserInventory(slot).Def
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = UserInventory(slot).GrhIndex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = UserInventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        Amount = UserInventory(slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = UserInventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = UserInventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = UserInventory(slot).Name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = UserInventory(slot).Equipped
End Property

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Single, ByVal eName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserInventory(slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
    End With
    
    'Render inventory slot (we don't need to render the whole inventory)
    Call DrawInventorySlot(slot)
End Sub

Private Function ClickItem(ByVal x As Long, ByVal y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = x \ TilePixelWidth
    temp_y = y \ TilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ TilePixelWidth) + 1
    
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        'Make sure slot isn't empty
        If UserInventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
End Function

Private Sub DrawInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Renders the inventory to the given PictureBox
'***************************************************
    Dim loopc As Long
    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim destRect As RECT
    
    With TempRect
        .Bottom = TilePixelHeight
        .Right = TilePixelWidth
    End With
    
    destRect = TempRect
    
    'Clear the inventory window
    InventoryWindow.Cls
    
    For loopc = InventoryOffset * (InventoryWindow.ScaleWidth \ TilePixelWidth) + 1 To UBound(UserInventory)
        'Clear the slot area
        Call InvSurface.BltColorFill(TempRect, InventoryWindow.BackColor)
        
        If UserInventory(loopc).GrhIndex Then
            'Get source rect
            With SrcRect
                .Left = GrhData(UserInventory(loopc).GrhIndex).sX
                .Top = GrhData(UserInventory(loopc).GrhIndex).sY
                .Right = .Left + TilePixelWidth
                .Bottom = .Top + TilePixelHeight
            End With
            
            'Render the item grh and the amount
            InvSurface.SetForeColor vbWhite
            InvSurface.BltFast 0, 0, SurfaceDB.Surface(GrhData(UserInventory(loopc).GrhIndex).FileNum), SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
            InvSurface.DrawText 0, 0, UserInventory(loopc).Amount, False
            
            'If equipped we render "E"
            If UserInventory(loopc).Equipped Then
                InvSurface.SetForeColor vbYellow
                InvSurface.DrawText 20, 20, "E", False
            End If
            
            'Render a box around the selected item
            If InvSelectedItem = loopc Then
                InvSurface.SetForeColor vbGreen
                InvSurface.setDrawStyle DrawStyleConstants.vbDot
                InvSurface.DrawBox 0, 0, TilePixelWidth, TilePixelHeight
            End If
        End If
        
        'Render the item to the Inventory Window
        Call InvSurface.BltToDC(InventoryWindow.hdc, TempRect, destRect)
        
        Call InventoryWindow.Refresh
        
        With destRect
            'Check if we have reached the right end of the inventory picturebox
            If .Left = (InventoryWindow.ScaleWidth - (InventoryWindow.ScaleWidth Mod TilePixelWidth) - TilePixelWidth) Then
                .Left = 0
                .Right = TilePixelWidth
                .Bottom = .Bottom + TilePixelHeight
                .Top = .Top + TilePixelHeight
            Else
                .Left = .Left + TilePixelWidth
                .Right = .Right + TilePixelWidth
            End If
        End With
    Next loopc
End Sub

Private Sub DrawInventorySlot(ByVal slot As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Renders a inventory slot to the given PictureBox
'***************************************************
    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim destRect As RECT
    
    With TempRect
        .Bottom = TilePixelHeight
        .Right = TilePixelWidth
    End With
    
    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth \ TilePixelWidth) + 1
    
    'If not in renderable area we exit
    If slot < upperLeftSlot Then Exit Sub
    
    With destRect
        .Top = ((slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth \ TilePixelWidth)) * TilePixelHeight
        .Left = ((slot - 1) Mod (InventoryWindow.ScaleWidth \ TilePixelWidth)) * TilePixelWidth
        .Bottom = .Top + TilePixelHeight
        .Right = .Left + TilePixelWidth
    End With
    
    'Clear the slot area
    Call InvSurface.BltColorFill(TempRect, InventoryWindow.BackColor)
    
    If UserInventory(slot).GrhIndex Then
        'Get source rect
        With SrcRect
            .Left = GrhData(UserInventory(slot).GrhIndex).sX
            .Top = GrhData(UserInventory(slot).GrhIndex).sY
            .Right = .Left + TilePixelWidth
            .Bottom = .Top + TilePixelHeight
        End With
        
        'Render the item grh and the amount
        InvSurface.SetForeColor vbWhite
        InvSurface.BltFast 0, 0, SurfaceDB.Surface(CLng(GrhData(UserInventory(slot).GrhIndex).FileNum)), SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
        InvSurface.DrawText 0, 0, UserInventory(slot).Amount, False
        
        'If equipped we render "E"
        If UserInventory(slot).Equipped Then
            InvSurface.SetForeColor vbYellow
            InvSurface.DrawText 20, 20, "E", False
        End If
        
        'Render a box around the selected item
        If InvSelectedItem = slot Then
            InvSurface.SetForeColor vbGreen
            InvSurface.setDrawStyle DrawStyleConstants.vbDot
            InvSurface.DrawBox 0, 0, TilePixelWidth, TilePixelHeight
        End If
    End If
    
    'Render the item to the Inventory Window
    InvSurface.BltToDC InventoryWindow.hdc, TempRect, destRect
    
    Call InventoryWindow.Refresh
End Sub

#If ConMenuesConextuales = 1 Then   'Argument list varies!!!
    Public Sub Initialize(ByRef DirectDraw As DirectDraw7, ByRef InvPic As PictureBox, ByRef PopUpMenu As Menu)
#Else   'Argument list varies!!!
    Public Sub Initialize(ByRef DirectDraw As DirectDraw7, ByRef InvPic As PictureBox)
#End If
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the reference to the picture box on which to render the inventory
'***************************************************
    Set InventoryWindow = InvPic
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    'initialize DX stuff
    Dim SurfaceDesc As DDSURFACEDESC2
    
    'Make sure DirectDraw was correctly initialized
    If DirectDraw Is Nothing Then Exit Sub
    
    'Set surface's description
    With SurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lHeight = TilePixelHeight
        .lWidth = TilePixelWidth
    End With

    'Create inventory surface
    Set InvSurface = DirectDraw.CreateSurface(SurfaceDesc)
    
    'Prepare surfaces for text rendering
    InvSurface.SetFontTransparency True
'TODO : Fonts should be in a separate class / collection
    Dim font As New StdFont
    Dim Ifnt As IFont
    
    font.Name = "Verdana"
    font.bold = True
    font.italic = False
    font.size = 7
    font.Underline = False
    font.Strikethrough = False
    
    Set Ifnt = font
    
    InvSurface.SetFont Ifnt
'\TODO
    
    InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
    
#If ConMenuesConextuales = 1 Then
    Set ItemMenu = PopUpMenu
#End If
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's area
    If x < 0 Or y < 0 Or x > InventoryWindow.Width Or y > InventoryWindow.Height Then _
        Exit Sub
    
    temp_x = x \ TilePixelWidth
    temp_y = y \ TilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ TilePixelWidth) + 1
    
    If TempItem <= MAX_INVENTORY_SLOTS Then
        InventoryWindow.ToolTipText = UserInventory(TempItem).Name
    End If
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    
    'Exit if it got outside the control's area
    If x < 0 Or y < 0 Or x > InventoryWindow.Width Or y > InventoryWindow.Height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(x), CInt(y))
    
    'Update needed inventory slots
    If prevSelItem <> InvSelectedItem Then
        If prevSelItem <> 0 And prevSelItem <> FLAGORO Then Call DrawInventorySlot(prevSelItem)
        If InvSelectedItem Then Call DrawInventorySlot(InvSelectedItem)
    End If
    
#If ConMenuesConextuales = 1 Then
    'If it was a right click on the same item we had selected before, show popup menu
    If Button = vbRightButton And prevSelItem = InvSelectedItem Then
        'Show the provided menu as a popup
        Call InventoryWindow.Parent.PopUpMenu(ItemMenu, , x, y)
    End If
#End If
End Sub
