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

Private InventoryOffset As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem As Long             'Currently selected item

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer

Private ShowValue As Boolean
Private ImgContainer As Boolean

Private InvSurface As DirectDrawSurface7            'DD Surface used to render everything

Private Inventory() As Inventory    'User's inventory

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
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

Public Sub ScrollInventory(ByVal Up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 14/01/2010
'Scrolls the graphical inventory up / down
'14/01/2010: ZaMa - Now redraws inventory only if needed
'***************************************************

    'Only move if there is something to show up /down
    If CanScroll(Up) Then
        InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
        
        'Redraw the inventory if needed
        DrawInventory
    End If

End Sub

Public Function CanScroll(ByVal Up As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 14/01/2010
'Determines wether inventory can be scrolled up / down
'***************************************************
    If Up Then
        CanScroll = (InventoryOffset + (InventoryWindow.ScaleHeight \ ownTilePixelHeight)) * _
                    (InventoryWindow.ScaleWidth \ ownTilePixelWidth) < UBound(Inventory)
    Else
        CanScroll = InventoryOffset > 0
    End If
    
End Function

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
    Call DrawInventorySlot(ClearSlot)

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
    Call DrawInventorySlot(slot)

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
    MaxHit = Inventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = Inventory(slot).MinHit
End Property

Public Property Get MaxDef(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
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

Public Property Get GrhIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = Inventory(slot).GrhIndex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = Inventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
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
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = Inventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = Inventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = Inventory(slot).Name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
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
    For slot = 1 To Me.MaxObjs
        DrawInventorySlot (slot)
    Next slot
End Sub

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Long, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal MaxDef As Integer, ByVal MinDef As Integer, ByVal eValor As Single, ByVal eName As String)
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
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
    End With
    
    'If InventoryWindow.Visible = False Then InventoryWindow.Visible = True
    
    'Render inventory slot (we don't need to render the whole inventory)
    Call DrawInventorySlot(slot)
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    
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
        .Bottom = ownTilePixelHeight
        .Right = ownTilePixelWidth
    End With
    
    destRect = TempRect
    
    'Clear the inventory window
    InventoryWindow.Cls
    
    For loopc = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(Inventory)
        'Clear the slot area
        Call InvSurface.BltColorFill(TempRect, InventoryWindow.BackColor)
        
        If Inventory(loopc).GrhIndex Then
            'Get source rect
            With SrcRect
                .Left = GrhData(Inventory(loopc).GrhIndex).sX
                .Top = GrhData(Inventory(loopc).GrhIndex).sY
                .Right = .Left + ownTilePixelWidth
                .Bottom = .Top + ownTilePixelHeight
            End With
            
            'Render the item grh and the amount
            InvSurface.SetForeColor vbWhite
            InvSurface.BltFast startdX, startdY, SurfaceDB.Surface(GrhData(Inventory(loopc).GrhIndex).FileNum), SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
            InvSurface.DrawText startdX / 2, startdY / 2, Inventory(loopc).Amount, False
            
            'If equipped we render "E"
            If Inventory(loopc).Equipped Then
                InvSurface.SetForeColor vbYellow
                InvSurface.DrawText 20, 20, "E", False
            End If
            
            'Render a box around the selected item
            If InvSelectedItem = loopc Then
                If ImgContainer Then
                    InvSurface.SetForeColor vbGreen
                    InvSurface.setDrawStyle DrawStyleConstants.vbDot
                    InvSurface.DrawBox 0, 0, ownTilePixelWidth, ownTilePixelHeight
                End If
            End If
        End If
        
        'Render the item to the Inventory Window
        Call InvSurface.BltToDC(InventoryWindow.hdc, TempRect, destRect)
        
        Call InventoryWindow.Refresh
        
        With destRect
            'Check if we have reached the right end of the inventory picturebox
            If .Left = (InventoryWindow.ScaleWidth - (InventoryWindow.ScaleWidth Mod ownTilePixelWidth) - ownTilePixelWidth) Then
                .Left = 0
                .Right = ownTilePixelWidth
                .Bottom = .Bottom + ownTilePixelHeight
                .Top = .Top + ownTilePixelHeight
            Else
                .Left = .Left + ownTilePixelWidth
                .Right = .Right + ownTilePixelWidth
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
    Dim SrcRectSlot As RECT
    
    With TempRect
        .Bottom = ownTilePixelHeight
        .Right = ownTilePixelWidth
    End With
    
    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    'If not in renderable area we exit
    If slot < upperLeftSlot Then Exit Sub
    
    With destRect
        .Top = ((slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth \ ownTilePixelWidth)) * ownTilePixelHeight
        .Left = ((slot - 1) Mod (InventoryWindow.ScaleWidth \ ownTilePixelWidth)) * ownTilePixelWidth
        .Bottom = .Top + ownTilePixelHeight
        .Right = .Left + ownTilePixelWidth
    End With
    
    'Clear the slot area
    Call InvSurface.BltColorFill(TempRect, InventoryWindow.BackColor)
    
    If Inventory(slot).GrhIndex Then
        'Get source rect
        With SrcRect
            .Left = GrhData(Inventory(slot).GrhIndex).sX
            .Top = GrhData(Inventory(slot).GrhIndex).sY
            .Right = .Left + ownTilePixelWidth
            .Bottom = .Top + ownTilePixelHeight
        End With
        With SrcRectSlot
            .Left = GrhData(Inventory(slot).GrhIndex).sX
            .Top = GrhData(Inventory(slot).GrhIndex).sY
            .Right = .Left + TilePixelWidth
            .Bottom = .Top + TilePixelHeight
        End With
        
        'Render the item grh and the amount
        InvSurface.SetForeColor vbWhite
        InvSurface.BltFast startdX, startdY, SurfaceDB.Surface(CLng(GrhData(Inventory(slot).GrhIndex).FileNum)), SrcRectSlot, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
        If ShowValue Then InvSurface.DrawText startdX / 2, startdY / 2, Inventory(slot).Amount, False

        
        'If equipped we render "E"
        If Inventory(slot).Equipped Then
            InvSurface.SetForeColor vbYellow
            InvSurface.DrawText 20, 20, "E", False
        End If
        
        'Render a box around the selected item
        If InvSelectedItem = slot Then
            If ImgContainer Then
                InvSurface.SetForeColor vbGreen
                InvSurface.setDrawStyle DrawStyleConstants.vbDot
                InvSurface.DrawBox 0, 0, ownTilePixelWidth, ownTilePixelHeight
            End If
        End If
    End If
    
    'Render the item to the Inventory Window
    InvSurface.BltToDC InventoryWindow.hdc, TempRect, destRect
    
    Call InventoryWindow.Refresh
End Sub

#If ConMenuesConextuales = 1 Then   'Argument list varies!!!
    Public Sub Initialize(ByRef DirectDraw As DirectDraw7, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, _
    ByRef PopUpMenu As Menu)

#Else   'Argument list varies!!!
    Public Sub Initialize(ByRef DirectDraw As DirectDraw7, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, _
                            Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 32, _
                            Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, _
                            Optional ByVal startY As Integer = 0, Optional ByVal bImgContainer As Boolean = True, _
                            Optional ByVal bShowText As Boolean = True)
#End If
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 03/12/09
'Sets the reference to the picture box on which to render the inventory
'03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
'***************************************************
    Set InventoryWindow = InvPic
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    ReDim Inventory(1 To MaxObjs) As Inventory
    
    'initialize DX stuff
    Dim SurfaceDesc As DDSURFACEDESC2
    
    'Make sure DirectDraw was correctly initialized
    If DirectDraw Is Nothing Then Exit Sub
      
    'Set surface's description
    With SurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lHeight = TileHeight
        .lWidth = TileWidth
    End With

    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    
    ImgContainer = bImgContainer
    ShowValue = bShowText
   
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
    font.size = FontSize
    font.Underline = False
    font.Strikethrough = False
    
    Set Ifnt = font
    
    InvSurface.SetFont Ifnt
'\TODO
    
    If ImgContainer Then _
        InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
    
#If ConMenuesConextuales = 1 Then
    Set ItemMenu = PopUpMenu
#End If
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    If TempItem <= UBound(Inventory) Then
        InventoryWindow.ToolTipText = Inventory(TempItem).Name
    End If
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
    
    'Update needed inventory slots
    If prevSelItem <> InvSelectedItem Then
        If prevSelItem <> 0 And prevSelItem <> FLAGORO Then Call DrawInventorySlot(prevSelItem)
        If InvSelectedItem Then Call DrawInventorySlot(InvSelectedItem)
    End If
    
#If ConMenuesConextuales = 1 Then
    'If it was a right click on the same item we had selected before, show popup menu
    If Button = vbRightButton And prevSelItem = InvSelectedItem Then
        'Show the provided menu as a popup
        Call InventoryWindow.Parent.PopUpMenu(ItemMenu, , X, Y)
    End If
#End If
End Sub
