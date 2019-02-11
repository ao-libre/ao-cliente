VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuildDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const p_SEGUNDOS = 7

Private p_CANTIDADDIALOGOS As Byte

Private Type t_GuildDlg

    Texto As String
    Segundos As Byte

End Type

Private p_Dialogos() As t_GuildDlg
Private p_Activo     As Boolean

Private usedFont     As StdFont

Public Property Let CantidadDialogos(ByVal v As Byte)
    
    On Error GoTo CantidadDialogos_Err
    

    If v > 0 Then
        ReDim Preserve p_Dialogos(1 To v) As t_GuildDlg
        p_CANTIDADDIALOGOS = v

    End If

    
    Exit Property

CantidadDialogos_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGuildDlg" & "->" & "CantidadDialogos"
    End If
Resume Next
    
End Property

''
' Removes all dialogs from screen.

Public Sub RemoveDialogs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/02/07
    'Removes all dialogs
    '***************************************************
    
    On Error GoTo RemoveDialogs_Err
    
    Dim i As Long
    
    For i = 1 To p_CANTIDADDIALOGOS
        p_Dialogos(i).Texto = vbNullString
    Next i

    
    Exit Sub

RemoveDialogs_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGuildDlg" & "->" & "RemoveDialogs"
    End If
Resume Next
    
End Sub

''
' Retrieves the number of dialogs to be displayed on screen.
'
' @return   The number of dialogs to be displayed on screen.

Public Property Get CantidadDialogos() As Byte
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/01/07
    'Retrieves the number of dialogs to be displayed on screen
    '***************************************************
    
    On Error GoTo CantidadDialogos_Err
    
    CantidadDialogos = p_CANTIDADDIALOGOS

    
    Exit Property

CantidadDialogos_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGuildDlg" & "->" & "CantidadDialogos"
    End If
Resume Next
    
End Property

Public Property Let Activo(ByVal v As Boolean)
    
    On Error GoTo Activo_Err
    
    p_Activo = v

    
    Exit Property

Activo_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGuildDlg" & "->" & "Activo"
    End If
Resume Next
    
End Property

Public Property Get Activo() As Boolean
    
    On Error GoTo Activo_Err
    
    Activo = p_Activo

    
    Exit Property

Activo_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGuildDlg" & "->" & "Activo"
    End If
Resume Next
    
End Property

Private Sub Class_Initialize()
    
    On Error GoTo Class_Initialize_Err
    
    p_CANTIDADDIALOGOS = 5
    ReDim p_Dialogos(1 To p_CANTIDADDIALOGOS) As t_GuildDlg
    p_Activo = True

    
    Exit Sub

Class_Initialize_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGuildDlg" & "->" & "Class_Initialize"
    End If
Resume Next
    
End Sub

Public Sub Draw()
    
    On Error GoTo Draw_Err
    
    Dim i As Byte
    Dim Y As Long

    If Not p_Activo Then Exit Sub

    Y = 250

    '500 600 700 800 900
    For i = 1 To p_CANTIDADDIALOGOS

        If Trim$(p_Dialogos(i).Texto) <> vbNullString Then
            If Left$(p_Dialogos(i).Texto, 1) = Chr$(3) Then
                'Call Mod_TileEngine.RenderText(260, Y + (i * 10), Mid$(p_Dialogos(i).Texto, 2), vbGreen, usedFont)
                'DrawText1 TexthDC, 260, Y + (i * 10), Mid$(p_Dialogos(i).Texto, 2), vbGreen
                
            Else

                'Call Mod_TileEngine.RenderText(260, Y + (i * 10), p_Dialogos(i).Texto, vbYellow, usedFont)
                'DrawText1 TexthDC, 260, Y + (i * 10), p_Dialogos(i).Texto, vbYellow
            End If

        End If

    Next i

    
    Exit Sub

Draw_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGuildDlg" & "->" & "Draw"
    End If
Resume Next
    
End Sub

Public Sub PassTimer()
    
    On Error GoTo PassTimer_Err
    
    Dim i As Byte

    For i = 1 To p_CANTIDADDIALOGOS

        If p_Dialogos(i).Segundos > 0 Then
            p_Dialogos(i).Segundos = p_Dialogos(i).Segundos - 1
        Else
            p_Dialogos(i).Texto = vbNullString

        End If

    Next i
    
    
    Exit Sub

PassTimer_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGuildDlg" & "->" & "PassTimer"
    End If
Resume Next
    
End Sub

''
' Splits dialogs into lines fitting properly the render area and inserts them

Public Sub PushBackText(ByVal s As String)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 07/04/2009
    'Splits dialogs into lines fitting properly the render area and inserts them
    '07/04/2009: Now cuts the string properly in spite of not fitting in the screen. This avoids an infite loop.
    '***************************************************
    
    On Error GoTo PushBackText_Err
    
    Dim str As String
    Dim tmp As Integer
    
    str = s
    
    ' If it's too long to fit, split it
    Do While frmMain.TextWidth(str) > 500
        tmp = InStrRev(str, " ")
        
        If tmp = 0 Then Exit Do
        
        str = Left$(str, tmp - 1)
    Loop
    
    'Show message and continue with the rest
    Call RealPushBackText(str)
    
    If LenB(str) <> LenB(s) Then
        Call PushBackText(Right$(s, Len(s) - Len(str) - 1))

    End If

    
    Exit Sub

PushBackText_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGuildDlg" & "->" & "PushBackText"
    End If
Resume Next
    
End Sub

Private Sub RealPushBackText(ByVal s As String)
    
    On Error GoTo RealPushBackText_Err
    
    Dim i     As Byte
    Dim Vacio As Boolean

    If p_CANTIDADDIALOGOS > 0 Then
        i = p_CANTIDADDIALOGOS
        Vacio = True

        While i > 0 And Vacio

            Vacio = p_Dialogos(i).Texto = vbNullString

            If Vacio Then i = i - 1
        Wend

        If i = p_CANTIDADDIALOGOS Then
            'hay q scrollear, estamos llenos
            i = 1

            While i < p_CANTIDADDIALOGOS

                p_Dialogos(i) = p_Dialogos(i + 1)
                i = i + 1
            Wend
            p_Dialogos(i).Texto = s
            p_Dialogos(i).Segundos = p_SEGUNDOS
        Else
            p_Dialogos(i + 1).Texto = s
            p_Dialogos(i + 1).Segundos = p_SEGUNDOS

        End If

    End If

    
    Exit Sub

RealPushBackText_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGuildDlg" & "->" & "RealPushBackText"
    End If
Resume Next
    
End Sub

''
' Sets the font being used to render dialogs.
'
' @param    nValue  The new font to be used to render dialogs.

Public Property Let Font(ByRef nValue As StdFont)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 04/27/06
    'Sets the font being used to render dialogs.
    '**************************************************************
    
    On Error GoTo Font_Err
    
    Set usedFont = nValue

    
    Exit Property

Font_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsGuildDlg" & "->" & "Font"
    End If
Resume Next
    
End Property

