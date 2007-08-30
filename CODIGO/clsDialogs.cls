VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsDialogs.cls - Chats over head manager.
' Creates and manipulates chats over heads, rendering and updating them.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
' Creates and manipulates chats over heads, rendering and updating them.
'
' @file     clsDialogs.cls
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20070729

Option Explicit
Option Base 0

''
' Maximum number of dialogs that can exist.
Private Const MAX_DIALOGS As Byte = 100

''
' Maximum length of any dialog line without having to split it.
Private Const MAX_LENGTH As Byte = 18

''
' Number of milliseconds to add to the lifetime per dialog character
Private Const MS_PER_CHAR As Byte = 100

''
' Number of extra milliseconds to add to the lifetime of a new dialog
Private Const MS_ADD_EXTRA As Integer = 5000

''
' The dialog structure
'
' @param    textLine    Array of lines of a formated chat.
' @param    x           X coord for rendering.
' @param    y           Y coord for rendering.
' @param    startTime   The time (in ms) at which the dialog was created.
' @param    lifeTime    Time (in ms) this dialog should last.
' @param    charIndex   The charIndex that created this dialog.
' @param    color       The color to be used when rendering the text.
' @param    renderable  Set to True if the chat should be rendered this frame, False otherwise
'                           (used to skip dialogs from people outside render area).
Private Type dialog
    textLine()  As String
    x           As Integer
    y           As Integer
    startTime   As Long
    lifeTime    As Long
    CharIndex   As Integer
    color       As Long
    renderable  As Boolean
End Type

''
' Array if dialogs, sorted by the charIndex.
Private dialogs(MAX_DIALOGS - 1) As dialog

''
' The number of dialogs being used at the moment.
Private dialogCount As Byte

''
' the size of the font being used to render
Private usedFont As StdFont

''
' A millisecond precision timer
Private Declare Function GetTickCount Lib "kernel32" () As Long

''
' Initializes the instance of the class.

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/28/07
'Initializes the dialog instance
'***************************************************
    dialogCount = 0
End Sub

''
' Clean up and release resources

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/28/07
'Clean up
'***************************************************
End Sub

Private Function BinarySearch(ByVal CharIndex As Integer) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Returns the index of the dialog in the list, or the negation
'of the position were it should be if not found (for binary insertion)
'**************************************************************
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = dialogCount - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If dialogs(mid).CharIndex < CharIndex Then
            min = mid + 1
        ElseIf dialogs(mid).CharIndex > CharIndex Then
            max = mid - 1
        Else
            'We found it
            BinarySearch = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    BinarySearch = Not min
End Function

''
' Creates a new dialog. If the user already had a dialog, then it's replaced,
' otherwise a new slot is used.
'
' @param    chat        The text of the chat being spoken.
' @param    charIndex   The character saying the chat.
' @param    color       The color being used in the chat when rendered.
'
' @remarks  If there are no more free slots, and the char didn't have a previous dialog,
'           the chat is ignored.

Public Sub CreateDialog(ByRef chat As String, ByVal CharIndex As Integer, ByVal color As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modification: 08/29/07 by Rapsodius
'Changed life time of the dialogs to a variable amount,
'depending on the number of chars
'**************************************************************
    Dim slot As Integer
    Dim i As Long
    
    slot = BinarySearch(CharIndex)
    
    If slot < 0 Then
        If dialogCount = MAX_DIALOGS Then Exit Sub  'Out of space! Should never happen....
        
        'We need to add it. Get insertion index and move list backwards.
        slot = Not slot
        
        For i = dialogCount To slot + 1 Step -1
            dialogs(i) = dialogs(i - 1)
        Next i
        
        dialogCount = dialogCount + 1
    End If
    
    With dialogs(slot)
        .CharIndex = CharIndex
        .color = color
        .renderable = False
        .startTime = GetTickCount()
        .textLine = FormatChat(chat)
        .lifeTime = MS_ADD_EXTRA + (MS_PER_CHAR * Len(chat))
    End With
End Sub

''
' Splits a text into several lines to make it comply with the MAX_LENGTH unless it's impossible (a single word longer than MAX_LENGTH).
'
' @param    chat    The text to be formated.
'
' @return   The array of lines into which the text is splitted.
'
' @see      MAX_LENGTH

Private Function FormatChat(ByRef chat As String) As String()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Formats a dialog into different text lines.
'**************************************************************
    Dim word As String
    Dim curPos As Long
    Dim length As Long
    Dim acumLength As Long
    Dim lineLength As Long
    Dim wordLength As Long
    Dim curLine As Long
    Dim chatLines() As String
    
    'Initialize variables
    curLine = 0
    curPos = 1
    length = Len(chat)
    acumLength = 0
    lineLength = -1
    ReDim chatLines(FieldCount(chat, 32)) As String
    
    'Start formating
    Do While acumLength < length
        word = ReadField(curPos, chat, 32)
        
        wordLength = Len(word)
        
        ' Is the first word of the first line? (it's the only that can start at -1)
        If lineLength = -1 Then
            chatLines(curLine) = word
            
            lineLength = wordLength
            acumLength = wordLength
        Else
            ' Is the word too long to fit in this line?
            If lineLength + wordLength + 1 > MAX_LENGTH Then
                'Put it in the next line
                curLine = curLine + 1
                chatLines(curLine) = word
                
                lineLength = wordLength
            Else
                'Add it to this line
                chatLines(curLine) = chatLines(curLine) & " " & word
                
                lineLength = lineLength + wordLength + 1
            End If
            
            acumLength = acumLength + wordLength + 1
        End If
        
        'Increase to search for next word
        curPos = curPos + 1
    Loop
    
    ' If it's only one line, center text
    If curLine = 0 And length < MAX_LENGTH Then
        chatLines(curLine) = String((MAX_LENGTH - length) \ 2 + 1, " ") & chatLines(curLine)
    End If
    
    'Resize array to fit
    ReDim Preserve chatLines(curLine) As String
    
    FormatChat = chatLines
End Function

''
' Updates the position where a given dialog will be rendered.
'
' @param    x           The X coord where to start rendering.
' @param    y           The Y coord where to start rendering.
' @param    charIndex   The character whose dialog's position is being updated.

Public Sub UpdateDialogPos(ByVal x As Integer, ByVal y As Integer, ByVal CharIndex As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Updates a dialog's position.
'**************************************************************
    Dim slot As Integer
    
    slot = BinarySearch(CharIndex)
    
    If slot >= 0 Then
        With dialogs(slot)
            .x = x - (MAX_LENGTH \ 2) * 3
            .y = y - (UBound(.textLine()) * 3)
            .renderable = True
        End With
    End If
End Sub

''
' Displays all renderable dialogs on screen.

Public Sub Render()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modification: 08/29/07 by Rapsodius
'Changed the way it gets rid of older dialogs.
'**************************************************************
    Dim i As Long
    Dim j As Long
    Dim offset As Integer
    Dim removedDialogs As Long
    
    For i = 0 To dialogCount - 1
        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)
            If (GetTickCount() - .startTime) >= .lifeTime Then
                Call RemoveDialog(.CharIndex)
                
                removedDialogs = removedDialogs + 1
            ElseIf .renderable Then
                offset = -(usedFont.Size + 2) * UBound(.textLine())
                
                'Render it
                For j = 0 To UBound(.textLine())
                    Call RenderText(.x, .y + offset, .textLine(j), .color, usedFont)
                    offset = offset + usedFont.Size + 5
                Next j
                
                'Won't render it unless it's updated in next frmae
                .renderable = False
            End If
        End With
    Next i
End Sub

''
' Removes the dialog from the requested char.
'
' @param    charIndex   The character whose dialog will be removed.

Public Sub RemoveDialog(ByVal CharIndex As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Removes the dialog from the requested char.
'**************************************************************
    Dim slot As Integer
    Dim i As Long
    
    slot = BinarySearch(CharIndex)
    
    If slot < 0 Then Exit Sub
    
    For i = slot To MAX_DIALOGS - 2
        dialogs(i) = dialogs(i + 1)
    Next i
    
    dialogCount = dialogCount - 1
End Sub

''
' Removes all dialogs from all characters.

Public Sub RemoveAllDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Removes all dialogs from all chars.
'**************************************************************
    dialogCount = 0
End Sub

''
' Retrieves the font being used to render dialogs.
'
' @return   The font being used to render dialogs.

Public Property Get font() As StdFont
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Retrieves the font being used to render dialogs.
'**************************************************************
    font = usedFont
End Property

''
' Sets the font being used to render dialogs.
'
' @param    nValue  The new font to be used to render dialogs.

Public Property Let font(ByRef nValue As StdFont)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Sets the font being used to render dialogs.
'**************************************************************
    Set usedFont = nValue
End Property
