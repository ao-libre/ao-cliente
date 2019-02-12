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
Private Const MAX_DIALOGS  As Byte = 100

''
' Maximum length of any dialog line without having to split it.
Private Const MAX_LENGTH   As Byte = 18

''
' Number of milliseconds to add to the lifetime per dialog character
Private Const MS_PER_CHAR  As Byte = 100

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
    X           As Integer
    Y           As Integer
    startTime   As Long
    lifeTime    As Long
    CharIndex   As Integer
    Color       As Long
    renderable  As Boolean

End Type

''
' Array if dialogs, sorted by the charIndex.
Private dialogs(MAX_DIALOGS - 1) As dialog

''
' The number of dialogs being used at the moment.
Private dialogCount              As Byte

''
' the size of the font being used to render
Private usedFont                 As StdFont

''
' A millisecond precision timer

''
' Initializes the instance of the class.

Private Sub Class_Initialize()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 07/28/07
    'Initializes the dialog instance
    '***************************************************
    
    On Error GoTo Class_Initialize_Err
    
    dialogCount = 0

    
    Exit Sub

Class_Initialize_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsDialogs" & "->" & "Class_Initialize"
    End If
Resume Next
    
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
    
    On Error GoTo BinarySearch_Err
    
    Dim Min As Long
    Dim Max As Long
    Dim mid As Long
    
    Min = 0
    Max = dialogCount - 1
    
    Do While Min <= Max
        mid = (Min + Max) \ 2
        
        If dialogs(mid).CharIndex < CharIndex Then
            Min = mid + 1
        ElseIf dialogs(mid).CharIndex > CharIndex Then
            Max = mid - 1
        Else
            'We found it
            BinarySearch = mid
            Exit Function

        End If

    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    BinarySearch = Not Min

    
    Exit Function

BinarySearch_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsDialogs" & "->" & "BinarySearch"
    End If
Resume Next
    
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

Public Sub CreateDialog(ByRef chat As String, _
                        ByVal CharIndex As Integer, _
                        ByVal Color As Long)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modification: 08/29/07 by Rapsodius
    'Changed life time of the dialogs to a variable amount,
    'depending on the number of chars
    '**************************************************************
    
    On Error GoTo CreateDialog_Err
    
    Dim slot As Integer
    Dim i    As Long
    
    Dim slotAdd As Long
    
    slot = BinarySearch(CharIndex)
    slotAdd = slot + 1
    
    If slot < 0 Then
        If dialogCount = MAX_DIALOGS Then Exit Sub  'Out of space! Should never happen....
        
        'We need to add it. Get insertion index and move list backwards.
        slot = Not slot
        
        For i = dialogCount To slotAdd Step -1
            dialogs(i) = dialogs(i - 1)
        Next i
        
        dialogCount = dialogCount + 1

    End If
    
    With dialogs(slot)
        .CharIndex = CharIndex
        .Color = ColorToDX8(Color)
        .renderable = False
        .startTime = GetTickCount()
        .textLine = FormatChat(chat)
        .lifeTime = MS_ADD_EXTRA + (MS_PER_CHAR * Len(chat))

    End With

    
    Exit Sub

CreateDialog_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsDialogs" & "->" & "CreateDialog"
    End If
Resume Next
    
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
    
    On Error GoTo FormatChat_Err
    
    Dim word        As String
    Dim curPos      As Long
    Dim length      As Long
    Dim acumLength  As Long
    Dim lineLength  As Long
    Dim wordLength  As Long
    Dim curLine     As Long
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
        chatLines(curLine) = String$((MAX_LENGTH - length) \ 2 + 1, " ") & chatLines(curLine)

    End If
    
    'Resize array to fit
    ReDim Preserve chatLines(curLine) As String
    
    FormatChat = chatLines

    
    Exit Function

FormatChat_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsDialogs" & "->" & "FormatChat"
    End If
Resume Next
    
End Function

''
' Updates the position where a given dialog will be rendered.
'
' @param    x           The X coord where to start rendering.
' @param    y           The Y coord where to start rendering.
' @param    charIndex   The character whose dialog's position is being updated.

Public Sub UpdateDialogPos(ByVal X As Integer, _
                           ByVal Y As Integer, _
                           ByVal CharIndex As Integer)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 07/28/07
    'Updates a dialog's position.
    '**************************************************************
    
    On Error GoTo UpdateDialogPos_Err
    
    Dim slot As Integer
    
    slot = BinarySearch(CharIndex)
    
    If slot >= 0 Then

        With dialogs(slot)
            .X = X - (MAX_LENGTH \ 2) * 3
            .Y = Y - (UBound(.textLine()) * 3)
            .renderable = True

        End With

    End If

    
    Exit Sub

UpdateDialogPos_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsDialogs" & "->" & "UpdateDialogPos"
    End If
Resume Next
    
End Sub

''
' Displays all renderable dialogs on screen.

Public Sub Render()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modification: 08/29/07 by Rapsodius
    'Changed the way it gets rid of older dialogs.
    '**************************************************************
    
    On Error GoTo Render_Err
    
    Dim i              As Long
    Dim j              As Long
    Dim offset         As Integer
    Dim removedDialogs As Long
    Dim To_dialogCount As Long
    Dim Upper_textLine As Long
    
    To_dialogCount = dialogCount - 1
    
    For i = 0 To To_dialogCount

        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)

            If (GetTickCount() - .startTime) >= .lifeTime Then
                Call RemoveDialog(.CharIndex)
                
                removedDialogs = removedDialogs + 1
            ElseIf .renderable Then
                
                offset = -(usedFont.Size + 2) * UBound(.textLine())
                Upper_textLine = UBound(.textLine())
                
                'Render it
                For j = 0 To Upper_textLine
                    Call DrawText(.X, .Y + offset, .textLine(j), .Color)
                    offset = offset + usedFont.Size + 5
                Next j
                
                'Won't render it unless it's updated in next frmae
                .renderable = False

            End If

        End With

    Next i

    
    Exit Sub

Render_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsDialogs" & "->" & "Render"
    End If
Resume Next
    
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
    
    On Error GoTo RemoveDialog_Err
    
    Dim slot As Integer
    Dim i    As Long
    Dim To_dialogs As Long
    
    slot = BinarySearch(CharIndex)
    To_dialogs = MAX_DIALOGS - 2
    
    If slot < 0 Then Exit Sub
    
    For i = slot To To_dialogs
        dialogs(i) = dialogs(i + 1)
    Next i
    
    dialogCount = dialogCount - 1

    
    Exit Sub

RemoveDialog_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsDialogs" & "->" & "RemoveDialog"
    End If
Resume Next
    
End Sub

''
' Removes all dialogs from all characters.

Public Sub RemoveAllDialogs()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 07/28/07
    'Removes all dialogs from all chars.
    '**************************************************************
    
    On Error GoTo RemoveAllDialogs_Err
    
    dialogCount = 0

    
    Exit Sub

RemoveAllDialogs_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsDialogs" & "->" & "RemoveAllDialogs"
    End If
Resume Next
    
End Sub

''
' Retrieves the font being used to render dialogs.
'
' @return   The font being used to render dialogs.

Public Property Get Font() As StdFont
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 07/28/07
    'Retrieves the font being used to render dialogs.
    '**************************************************************
    
    On Error GoTo Font_Err
    
    Font = usedFont

    
    Exit Property

Font_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsDialogs" & "->" & "Font"
    End If
Resume Next
    
End Property

''
' Sets the font being used to render dialogs.
'
' @param    nValue  The new font to be used to render dialogs.

Public Property Let Font(ByRef nValue As StdFont)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 07/28/07
    'Sets the font being used to render dialogs.
    '**************************************************************
    
    On Error GoTo Font_Err
    
    Set usedFont = nValue

    
    Exit Property

Font_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsDialogs" & "->" & "Font"
    End If
Resume Next
    
End Property

