VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Forum
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

Private Const MAX_MENSAJES_FORO As Byte = 30
Private Const MAX_ANUNCIOS_FORO As Byte = 5

' Post counter
Private CantPosts(0 To 2) As Byte
Private CantStickys(0 To 2) As Byte

Public Sub AddPost(ByVal ForumType As eForumType, ByRef Title As String, ByRef Author As String, ByRef MsgPost As String, ByVal Sticky As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: -
'
'***************************************************

    With Foros(ForumType)
        If Sticky Then
            If CantStickys(ForumType) < MAX_ANUNCIOS_FORO Then _
                CantStickys(ForumType) = CantStickys(ForumType) + 1
        
            Call MoveArray(ForumType, Sticky)
            
            .StickyTitle(1) = Title
            .StickyAuthor(1) = Author
            .StickyPost(1) = MsgPost
        Else
            If CantPosts(ForumType) < MAX_MENSAJES_FORO Then _
                CantPosts(ForumType) = CantPosts(ForumType) + 1
            
            Call MoveArray(ForumType, Sticky)
            
            .GeneralTitle(1) = Title
            .GeneralAuthor(1) = Author
            .GeneralPost(1) = MsgPost
        End If
    End With

End Sub

Public Sub ClearForums()
'***************************************************
'Author: ZaMa
'Last Modify Date: -
'
'***************************************************

    Dim Index As Long
    
    If CantPosts(eForumType.ieGeneral) > 0 Then
        For Index = 1 To CantPosts(eForumType.ieGeneral)
            Foros(eForumType.ieGeneral).GeneralAuthor(Index) = vbNullString
            Foros(eForumType.ieGeneral).GeneralTitle(Index) = vbNullString
            Foros(eForumType.ieGeneral).GeneralPost(Index) = vbNullString
        Next Index
        
        CantPosts(eForumType.ieGeneral) = 0
    End If
    
    If CantPosts(eForumType.ieREAL) > 0 Then
        For Index = 1 To CantPosts(eForumType.ieREAL)
            Foros(eForumType.ieREAL).GeneralAuthor(Index) = vbNullString
            Foros(eForumType.ieREAL).GeneralTitle(Index) = vbNullString
            Foros(eForumType.ieREAL).GeneralPost(Index) = vbNullString
        Next Index
        
        CantPosts(eForumType.ieREAL) = 0
    End If
    
    If CantPosts(eForumType.ieCAOS) > 0 Then
        For Index = 1 To CantPosts(eForumType.ieCAOS)
            Foros(eForumType.ieCAOS).GeneralAuthor(Index) = vbNullString
            Foros(eForumType.ieCAOS).GeneralTitle(Index) = vbNullString
            Foros(eForumType.ieCAOS).GeneralPost(Index) = vbNullString
        Next Index
        
        CantPosts(eForumType.ieCAOS) = 0
    End If
    
    If CantStickys(eForumType.ieGeneral) > 0 Then
        For Index = 1 To CantStickys(eForumType.ieGeneral)
            Foros(eForumType.ieGeneral).StickyAuthor(Index) = vbNullString
            Foros(eForumType.ieGeneral).StickyTitle(Index) = vbNullString
            Foros(eForumType.ieGeneral).StickyPost(Index) = vbNullString
        Next Index
        
        CantStickys(eForumType.ieGeneral) = 0
    End If
    
    If CantStickys(eForumType.ieREAL) > 0 Then
        For Index = 1 To CantStickys(eForumType.ieREAL)
            Foros(eForumType.ieREAL).StickyAuthor(Index) = vbNullString
            Foros(eForumType.ieREAL).StickyTitle(Index) = vbNullString
            Foros(eForumType.ieREAL).StickyPost(Index) = vbNullString
        Next Index
        
        CantStickys(eForumType.ieREAL) = 0
    End If
    
    If CantStickys(eForumType.ieCAOS) > 0 Then
        For Index = 1 To CantStickys(eForumType.ieCAOS)
            Foros(eForumType.ieCAOS).StickyAuthor(Index) = vbNullString
            Foros(eForumType.ieCAOS).StickyTitle(Index) = vbNullString
            Foros(eForumType.ieCAOS).StickyPost(Index) = vbNullString
        Next Index
        
        CantStickys(eForumType.ieCAOS) = 0
    End If

End Sub

Public Function GetNroPost(ByVal ForumType As eForumType) As Byte
'***************************************************
'Author: ZaMa
'Last Modify Date: -
'
'***************************************************

    GetNroPost = CantPosts(ForumType)
End Function

Public Function GetNroSticky(ByVal ForumType As eForumType) As Byte
'***************************************************
'Author: ZaMa
'Last Modify Date: -
'
'***************************************************

    GetNroSticky = CantStickys(ForumType)
End Function

Private Sub MoveArray(ByVal ForumType As eForumType, ByVal Sticky As Boolean)
Dim i As Long

With Foros(ForumType)
    If Sticky Then
        For i = CantStickys(ForumType) To 2 Step -1
            .StickyTitle(i) = .StickyTitle(i - 1)
            .StickyPost(i) = .StickyPost(i - 1)
            .StickyAuthor(i) = .StickyAuthor(i - 1)
        Next i
    Else
        For i = CantPosts(ForumType) To 2 Step -1
            .GeneralTitle(i) = .GeneralTitle(i - 1)
            .GeneralPost(i) = .GeneralPost(i - 1)
            .GeneralAuthor(i) = .GeneralAuthor(i - 1)
        Next i
    End If
End With
End Sub
