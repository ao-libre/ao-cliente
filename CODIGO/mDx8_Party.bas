Attribute VB_Name = "mDx8_Party"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 26/05/10
'Blisse-AO | Party Engine! Live the Pepa! _
 Show in screen the Party Gays!
'***************************************************

Public Type c_PartyMember

    Name As String
    Head As Integer
    Lvl As Byte
    ExpParty As Long

End Type

Public Mostrar              As Byte
Public PartyMembers(1 To 5) As c_PartyMember

Public Sub Reset_Party()
    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 27/07/10
    'Reset all of Party Members
    '***************************************************
    
    On Error GoTo Reset_Party_Err
    
    Dim i As Byte

    For i = 1 To 5
        PartyMembers(i).ExpParty = 0
        PartyMembers(i).Head = 0
        PartyMembers(i).Lvl = 0
        PartyMembers(i).Name = vbNullString
    Next i

    
    Exit Sub

Reset_Party_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Party" & "->" & "Reset_Party"
    End If
Resume Next
    
End Sub

Public Sub Draw_Party_Members()
    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 26/05/10
    'Render Party Members
    '***************************************************
    
    On Error GoTo Draw_Party_Members_Err
    
    Dim i As Byte, Count As Byte
    Count = 0

    For i = 1 To 5

        If PartyMembers(i).Name <> "" Then
            Count = Count + 1
            Engine_Draw_Box 410, 20 + (Count - 1) * 50 + 5, 120, 40, D3DColorARGB(100, 0, 0, 0)
            DDrawTransGrhIndextoSurface HeadData(PartyMembers(i).Head).Head(3).GrhIndex, 410, 20 + (Count - 1) * 50 + 35, 1, Normal_RGBList(), 0, True

            'Fonts_Render_String PartyMembers(i).Name, 440, 20 + (Count - 1) * 50 + 10, D3DColorARGB(150, 255, 255, 255), 2
            'Fonts_Render_String "Nivel: " & PartyMembers(i).Lvl, 440, 20 + (Count - 1) * 50 + 20, D3DColorARGB(150, 255, 255, 255), 2
            'Fonts_Render_String "Exp: " & PartyMembers(i).ExpParty, 440, 20 + (Count - 1) * 50 + 30, D3DColorARGB(150, 255, 255, 255), 2
        End If

    Next i
            
    If Count <> 0 Then

        'Fonts_Render_String "Miembros de Party", 405, 5, D3DColorARGB(100, 255, 128, 0), 3
    End If

    
    Exit Sub

Draw_Party_Members_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Party" & "->" & "Draw_Party_Members"
    End If
Resume Next
    
End Sub

Public Sub Set_PartyMember(ByVal Member As Byte, _
                           Name As String, _
                           ExpParty As Long, _
                           Lvl As Byte, _
                           Head As Integer)
    
    On Error GoTo Set_PartyMember_Err
    

    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 28/05/10
    'Add User to Party
    '***************************************************
    If Member < 1 Or Member > 5 Then Exit Sub

    With PartyMembers(Member)
        .Name = Name
        .ExpParty = ExpParty
        .Head = Head
        .Lvl = Lvl

    End With

    
    Exit Sub

Set_PartyMember_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Party" & "->" & "Set_PartyMember"
    End If
Resume Next
    
End Sub

Public Sub Kick_PartyMember(ByVal Member As Byte)
    
    On Error GoTo Kick_PartyMember_Err
    

    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 27/05/10
    'Kick User From Party
    '***************************************************
    If Member < 1 Or Member > 5 Then Exit Sub

    With PartyMembers(Member)
        .Name = vbNullString
        .ExpParty = 0
        .Head = 0
        .Lvl = 0

    End With

    
    Exit Sub

Kick_PartyMember_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "mDx8_Party" & "->" & "Kick_PartyMember"
    End If
Resume Next
    
End Sub

