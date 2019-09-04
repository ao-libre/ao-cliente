Attribute VB_Name = "mDx8_Party"
Option Explicit

'***************************************************
'Author: Ezequiel Juarez (Standelf)
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

Public PartyMembers(1 To 5) As c_PartyMember

Public Sub Reset_Party()
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 27/07/10
'Reset all of Party Members
'***************************************************
    Dim i As Byte
        For i = 1 To 5
            PartyMembers(i).ExpParty = 0
            PartyMembers(i).Head = 0
            PartyMembers(i).Lvl = 0
            PartyMembers(i).Name = vbNullString
        Next i
End Sub

Public Sub Draw_Party_Members()
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 26/05/10
'Render Party Members
'***************************************************
        Dim i As Byte, Count As Byte
        Count = 0
            For i = 1 To 5
                If Len(PartyMembers(i).Name) > 0 Then
                    Count = Count + 1
                    Call Engine_Draw_Box(410, 20 + (Count - 1) * 50 + 5, 120, 40, D3DColorARGB(100, 0, 0, 0))
                    Call Draw_GrhIndex(HeadData(PartyMembers(i).Head).Head(3).GrhIndex, 410, 20 + (Count - 1) * 50 + 35, 1, Normal_RGBList(), 0, True)
                    'Fonts_Render_String PartyMembers(i).Name, 440, 20 + (Count - 1) * 50 + 10, D3DColorARGB(150, 255, 255, 255), 2
                    'Fonts_Render_String "Nivel: " & PartyMembers(i).Lvl, 440, 20 + (Count - 1) * 50 + 20, D3DColorARGB(150, 255, 255, 255), 2
                    'Fonts_Render_String "Exp: " & PartyMembers(i).ExpParty, 440, 20 + (Count - 1) * 50 + 30, D3DColorARGB(150, 255, 255, 255), 2
                End If
            Next i
            
            If Count <> 0 Then
                'Fonts_Render_String "Miembros de Party", 405, 5, D3DColorARGB(100, 255, 128, 0), 3
            End If
End Sub

Public Sub Set_PartyMember(ByVal Member As Byte, Name As String, ExpParty As Long, Lvl As Byte, Head As Integer)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
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
End Sub

Public Sub Kick_PartyMember(ByVal Member As Byte)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
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
End Sub




