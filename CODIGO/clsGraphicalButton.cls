VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical Button
'Designed and implemented by ZaMa
'zaxtor@gmail.com
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

Private WithEvents ButtonHolder As Image
Attribute ButtonHolder.VB_VarHelpID = -1

Private RollOverButtonImage     As Picture  'img holder
Private NormalButtonImage       As Picture 'img holder
Private ClickButtonImage        As Picture  'img holder
Private DisabledButtonImage     As Picture  'img holder

Private SoundClick              As Boolean

Private FormActual              As Form ' Form owner

Private OverButton              As Boolean 'Indicates wether the mouse is over the button or not
Private Disabled                As Boolean 'Indicates wether it's enabled or not

Public Sub Initialize(ByRef imgButton As Image, _
                      ByVal imgNormalPath As String, _
                      ByVal imgRollOverPath As String, _
                      ByVal imgClickPath As String, _
                      ByRef Formulario As Form, _
                      Optional ByVal imgDisabledPath As String, _
                      Optional ByVal StartDisabled As Boolean = False, _
                      Optional ByVal MakeClickSound As Boolean = True, _
                      Optional ByVal ShowHandPointer As Boolean = True)
    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 22/12/09
    'Sets the reference to the image container
    '***************************************************
    
    On Error GoTo Initialize_Err
    
    Set ButtonHolder = imgButton
    Set FormActual = Formulario
    
    Set RollOverButtonImage = LoadPicture(imgRollOverPath)
    Set NormalButtonImage = LoadPicture(imgNormalPath)
    Set ClickButtonImage = LoadPicture(imgClickPath)
    Set DisabledButtonImage = LoadPicture(imgDisabledPath)
    
    SoundClick = MakeClickSound
    
    With ButtonHolder

        If ShowHandPointer Then
            .MousePointer = vbCustom
            .MouseIcon = picMouseIcon

        End If
        
        If StartDisabled Then
            Disabled = True
            .Picture = DisabledButtonImage
        Else
            .Picture = NormalButtonImage

        End If

    End With
    
    
    Exit Sub

Initialize_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGraphicalButton" & "->" & "Initialize"
    End If
Resume Next
    
End Sub

Private Sub ButtonHolder_Click()
    
    On Error GoTo ButtonHolder_Click_Err
    

    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 02/01/2010
    'Implements the mouse click event of the image button.
    'Plays the Click Sound.
    '***************************************************
    If Disabled Then Exit Sub
    If Not SoundClick Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)

    
    Exit Sub

ButtonHolder_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGraphicalButton" & "->" & "ButtonHolder_Click"
    End If
Resume Next
    
End Sub

Private Sub ButtonHolder_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    
    On Error GoTo ButtonHolder_MouseDown_Err
    

    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 29/12/09
    'Implements the mouse down event of the image button.
    'Toggle to Click image button.
    '***************************************************
    If Disabled Then Exit Sub
    
    ButtonHolder.Picture = ClickButtonImage

    
    Exit Sub

ButtonHolder_MouseDown_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGraphicalButton" & "->" & "ButtonHolder_MouseDown"
    End If
Resume Next
    
End Sub

Private Sub ButtonHolder_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    
    On Error GoTo ButtonHolder_MouseMove_Err
    

    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 22/12/09
    'Implements the mouse move event of the image button.
    'Toggle to pressed image button.
    '***************************************************
    If Disabled Then Exit Sub
    
    If Not OverButton Then
    
        'This helps when two buttons are close
        FormActual.LastButtonPressed.ToggleToNormal
        
        ButtonHolder.Picture = RollOverButtonImage
        OverButton = True
        
        Set FormActual.LastButtonPressed = Me

    End If

    
    Exit Sub

ButtonHolder_MouseMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGraphicalButton" & "->" & "ButtonHolder_MouseMove"
    End If
Resume Next
    
End Sub

Public Sub ToggleToNormal()
    
    On Error GoTo ToggleToNormal_Err
    

    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 22/12/09
    'Toggle to normal image button.
    '***************************************************
    If Disabled Then Exit Sub
    
    If Not OverButton Then Exit Sub
    
    ButtonHolder.Picture = NormalButtonImage
    OverButton = False

    
    Exit Sub

ToggleToNormal_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGraphicalButton" & "->" & "ToggleToNormal"
    End If
Resume Next
    
End Sub

Private Sub ButtonHolder_MouseUp(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo ButtonHolder_MouseUp_Err
    

    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 29/12/09
    'Implements the mouse up event of the image button.
    'Toggle to RollOver image button.
    '***************************************************
    If Disabled Then Exit Sub
    
    ButtonHolder.Picture = RollOverButtonImage

    
    Exit Sub

ButtonHolder_MouseUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGraphicalButton" & "->" & "ButtonHolder_MouseUp"
    End If
Resume Next
    
End Sub

Public Sub EnableButton(Optional ByVal ToogleEnable As Boolean = True)
    
    On Error GoTo EnableButton_Err
    

    If Disabled = Not ToogleEnable Then Exit Sub
    
    OverButton = False
    Disabled = Not ToogleEnable

    If Disabled Then
        ButtonHolder.Picture = DisabledButtonImage
    Else
        ButtonHolder.Picture = NormalButtonImage

    End If
    
    
    Exit Sub

EnableButton_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGraphicalButton" & "->" & "EnableButton"
    End If
Resume Next
    
End Sub

Public Function IsEnabled() As Boolean
    
    On Error GoTo IsEnabled_Err
    
    IsEnabled = Not Disabled

    
    Exit Function

IsEnabled_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGraphicalButton" & "->" & "IsEnabled"
    End If
Resume Next
    
End Function
