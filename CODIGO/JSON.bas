Attribute VB_Name = "JSON"

' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed

Option Explicit

<<<<<<< HEAD
<<<<<<< HEAD
Private psErrors       As String

Public Function GetParserErrors() As String
    
    On Error GoTo GetParserErrors_Err
    
    GetParserErrors = psErrors

    
    Exit Function

GetParserErrors_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "GetParserErrors"
    End If
Resume Next
    
End Function

Public Function ClearParserErrors() As String
    
    On Error GoTo ClearParserErrors_Err
    
    psErrors = vbNullString

    
    Exit Function

ClearParserErrors_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "ClearParserErrors"
    End If
Resume Next
    
End Function

'
'   parse string and create JSON object
'
Public Function parse(ByRef str As String) As Object

    Dim Index As Long
    Index = 1
    psErrors = vbNullString

    On Error Resume Next

    Call skipChar(str, Index)

    Select Case mid$(str, Index, 1)

        Case "{"
            Set parse = parseObject(str, Index)

        Case "["
            Set parse = parseArray(str, Index)

        Case Else
            psErrors = "Invalid JSON"

    End Select

End Function

'
'   parse collection of key/value
'
Private Function parseObject(ByRef str As String, ByRef Index As Long) As Dictionary

    Set parseObject = New Dictionary
    Dim sKey As String
   
    ' "{"
    Call skipChar(str, Index)

    If mid$(str, Index, 1) <> "{" Then
        psErrors = psErrors & "Invalid Object at position " & Index & " : " & mid$(str, Index) & vbNewLine
        Exit Function

    End If
   
    Index = Index + 1

    Do
        Call skipChar(str, Index)

        If "}" = mid$(str, Index, 1) Then
            Index = Index + 1
            Exit Do
        ElseIf "," = mid$(str, Index, 1) Then
            Index = Index + 1
            Call skipChar(str, Index)
        ElseIf Index > Len(str) Then
            psErrors = psErrors & "Missing '}': " & Right$(str, 20) & vbNewLine
            Exit Do

        End If
      
        ' add key/value pair
        sKey = parseKey(str, Index)

        On Error Resume Next
      
        parseObject.Add sKey, parseValue(str, Index)

        If Err.number <> 0 Then
            psErrors = psErrors & Err.Description & ": " & sKey & vbNewLine
            Exit Do

        End If

    Loop
eh:
=======
' DECLARACIONES API
Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()

' CONSTANTES LOCALE API
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_SGROUPING = &H10

' CONSTANTES JSON
Private Const A_CURLY_BRACKET_OPEN   As Integer = 123       ' AscW("{")
Private Const A_CURLY_BRACKET_CLOSE  As Integer = 125       ' AscW("}")
Private Const A_SQUARE_BRACKET_OPEN  As Integer = 91        ' AscW("[")
Private Const A_SQUARE_BRACKET_CLOSE As Integer = 93        ' AscW("]")
Private Const A_BRACKET_OPEN         As Integer = 40        ' AscW("(")
Private Const A_BRACKET_CLOSE        As Integer = 41        ' AscW(")")
Private Const A_COMMA                As Integer = 44        ' AscW(",")
Private Const A_DOUBLE_QUOTE         As Integer = 34        ' AscW("""")
Private Const A_SINGLE_QUOTE         As Integer = 39        ' AscW("'")
Private Const A_BACKSLASH            As Integer = 92        ' AscW("\")
Private Const A_FORWARDSLASH         As Integer = 47        ' AscW("/")
Private Const A_COLON                As Integer = 58        ' AscW(":")
Private Const A_SPACE                As Integer = 32        ' AscW(" ")
Private Const A_ASTERIX              As Integer = 42        ' AscW("*")
Private Const A_VBCR                 As Integer = 13        ' AscW("vbcr")
Private Const A_VBLF                 As Integer = 10        ' AscW("vblf")
Private Const A_VBTAB                As Integer = 9         ' AscW("vbTab")
Private Const A_VBCRLF               As Integer = 13        ' AscW("vbcrlf")
Private Const A_b                    As Integer = 98        ' AscW("b")
Private Const A_f                    As Integer = 102       ' AscW("f")
Private Const A_n                    As Integer = 110       ' AscW("n")
Private Const A_r                    As Integer = 114       ' AscW("r"
Private Const A_t                    As Integer = 116       ' AscW("t"))
Private Const A_u                    As Integer = 117       ' AscW("u")

Private m_decSep                     As String
Private m_groupSep                   As String

Private m_parserrors                 As String

Private m_str()                      As Integer
Private m_length                     As Long
=======
Const INVALID_JSON      As Long = 1
Const INVALID_OBJECT    As Long = 2
Const INVALID_ARRAY     As Long = 3
Const INVALID_BOOLEAN   As Long = 4
Const INVALID_NULL      As Long = 5
Const INVALID_KEY       As Long = 6
Const INVALID_RPC_CALL  As Long = 7

Private psErrors As String
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'

Public Function GetParserErrors() As String
   GetParserErrors = psErrors
End Function

Public Function ClearParserErrors() As String
   psErrors = vbNullString
End Function


'
'   parse string and create JSON object
'
Public Function parse(ByRef str As String) As Object

   Dim Index As Long
   Index = 1
   psErrors = vbNullString
   On Error Resume Next
   Call skipChar(str, Index)
   Select Case mid$(str, Index, 1)
      Case "{"
         Set parse = parseObject(str, Index)
      Case "["
         Set parse = parseArray(str, Index)
      Case Else
         psErrors = "Invalid JSON"
   End Select

<<<<<<< HEAD
    Loop
>>>>>>> origin/master
=======
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'

End Function

 '
 '   parse collection of key/value
 '
Private Function parseObject(ByRef str As String, ByRef Index As Long) As Dictionary

<<<<<<< HEAD
    Set parseArray = New Collection
<<<<<<< HEAD

    ' "["
    Call skipChar(str, Index)

    If mid$(str, Index, 1) <> "[" Then
        psErrors = psErrors & "Invalid Array at position " & Index & " : " + mid$(str, Index, 20) & vbNewLine
        Exit Function

    End If
   
    Index = Index + 1

    Do

        Call skipChar(str, Index)

        If "]" = mid$(str, Index, 1) Then
            Index = Index + 1
            Exit Do
        ElseIf "," = mid$(str, Index, 1) Then
            Index = Index + 1
            Call skipChar(str, Index)
        ElseIf Index > Len(str) Then
            psErrors = psErrors & "Missing ']': " & Right$(str, 20) & vbNewLine
            Exit Do

        End If

        ' add value
        On Error Resume Next

        parseArray.Add parseValue(str, Index)

        If Err.number <> 0 Then
            psErrors = psErrors & Err.Description & ": " & mid$(str, Index, 20) & vbNewLine
=======

    Call skipChar(index)

    If mid$(str, index, 1) <> "[" Then
        m_parserrors = m_parserrors & "Array invalido en la posicion " & index & " : " + mid$(str, index, 20) & vbCrLf
        Exit Function
    End If
   
    index = index + 1

    Do
        Call skipChar(index)
    
        charint = m_str(index)
    
        If charint = A_SQUARE_BRACKET_CLOSE Then
            index = index + 1
            Exit Do
        ElseIf charint = A_COMMA Then
            index = index + 1
            Call skipChar(index)
        ElseIf index > m_length Then
            m_parserrors = m_parserrors & "Falta ']': " & Right(str, 20) & vbCrLf
            Exit Do
        End If
    
        'add value
        On Error Resume Next

        parseArray.Add parseValue(str, index)

        If Err.number <> 0 Then
            m_parserrors = m_parserrors & Err.Description & ": " & mid$(str, index, 20) & vbCrLf
>>>>>>> origin/master
            Exit Do

        End If
=======
   Set parseObject = New Dictionary
   Dim sKey As String
   
   ' "{"
   Call skipChar(str, Index)
   If mid$(str, Index, 1) <> "{" Then
      psErrors = psErrors & "Invalid Object at position " & Index & " : " & mid$(str, Index) & vbNewLine
      Exit Function
   End If
   
   Index = Index + 1

   Do
      Call skipChar(str, Index)
      If "}" = mid$(str, Index, 1) Then
         Index = Index + 1
         Exit Do
      ElseIf "," = mid$(str, Index, 1) Then
         Index = Index + 1
         Call skipChar(str, Index)
      ElseIf Index > Len(str) Then
         psErrors = psErrors & "Missing '}': " & Right$(str, 20) & vbNewLine
         Exit Do
      End If
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'

      
      ' add key/value pair
      sKey = parseKey(str, Index)
      On Error Resume Next
      
      parseObject.Add sKey, parseValue(str, Index)
      If Err.number <> 0 Then
         psErrors = psErrors & Err.Description & ": " & sKey & vbNewLine
         Exit Do
      End If
   Loop
eh:

End Function

<<<<<<< HEAD
<<<<<<< HEAD
'
'   parse string / number / object / array / true / false / null
'
Private Function parseValue(ByRef str As String, ByRef Index As Long)
    
    On Error GoTo parseValue_Err
    

    Call skipChar(str, Index)

    Select Case mid$(str, Index, 1)

        Case "{"
            Set parseValue = parseObject(str, Index)

        Case "["
            Set parseValue = parseArray(str, Index)
=======
Private Function parseValue(ByRef str As String, ByRef index As Long)

    Call skipChar(index)

    Select Case m_str(index)

        Case A_DOUBLE_QUOTE, A_SINGLE_QUOTE
            parseValue = parseString(str, index)
            Exit Function

        Case A_SQUARE_BRACKET_OPEN
            Set parseValue = parseArray(str, index)
            Exit Function
=======
'
'   parse list
'
Private Function parseArray(ByRef str As String, ByRef Index As Long) As Collection
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'

   Set parseArray = New Collection

   ' "["
   Call skipChar(str, Index)
   If mid$(str, Index, 1) <> "[" Then
      psErrors = psErrors & "Invalid Array at position " & Index & " : " + mid$(str, Index, 20) & vbNewLine
      Exit Function
   End If
   
   Index = Index + 1

   Do

      Call skipChar(str, Index)
      If "]" = mid$(str, Index, 1) Then
         Index = Index + 1
         Exit Do
      ElseIf "," = mid$(str, Index, 1) Then
         Index = Index + 1
         Call skipChar(str, Index)
      ElseIf Index > Len(str) Then
         psErrors = psErrors & "Missing ']': " & Right$(str, 20) & vbNewLine
         Exit Do
      End If

      ' add value
      On Error Resume Next
      parseArray.Add parseValue(str, Index)
      If Err.number <> 0 Then
         psErrors = psErrors & Err.Description & ": " & mid$(str, Index, 20) & vbNewLine
         Exit Do
      End If
   Loop

End Function

<<<<<<< HEAD
    End Select
>>>>>>> origin/master

        Case """", "'"
            parseValue = parseString(str, Index)
=======
'
'   parse string / number / object / array / true / false / null
'
Private Function parseValue(ByRef str As String, ByRef Index As Long)

   Call skipChar(str, Index)

   Select Case mid$(str, Index, 1)
      Case "{"
         Set parseValue = parseObject(str, Index)
      Case "["
         Set parseValue = parseArray(str, Index)
      Case """", "'"
         parseValue = parseString(str, Index)
      Case "t", "f"
         parseValue = parseBoolean(str, Index)
      Case "n"
         parseValue = parseNull(str, Index)
      Case Else
         parseValue = parseNumber(str, Index)
   End Select
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'

        Case "t", "f"
            parseValue = parseBoolean(str, Index)

        Case "n"
            parseValue = parseNull(str, Index)

        Case Else
            parseValue = parseNumber(str, Index)

    End Select

    
    Exit Function

parseValue_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "parseValue"
    End If
Resume Next
    
End Function

<<<<<<< HEAD
'
'   parse string
'
Private Function parseString(ByRef str As String, ByRef Index As Long) As String
    
    On Error GoTo parseString_Err
    

    Dim quote As String
    Dim Char  As String
    Dim code  As String

    Dim SB    As New cStringBuilder

    Call skipChar(str, Index)
    quote = mid$(str, Index, 1)
    Index = Index + 1
   
    Do While Index > 0 And Index <= Len(str)
        Char = mid$(str, Index, 1)

        Select Case (Char)

            Case "\"
                Index = Index + 1
                Char = mid$(str, Index, 1)

                Select Case (Char)

                    Case """", "\", "/", "'"
                        SB.Append Char
                        Index = Index + 1

                    Case "b"
                        SB.Append vbBack
                        Index = Index + 1

                    Case "f"
                        SB.Append vbFormFeed
                        Index = Index + 1

                    Case "n"
                        SB.Append vbLf
                        Index = Index + 1

                    Case "r"
                        SB.Append vbCr
                        Index = Index + 1

                    Case "t"
                        SB.Append vbTab
                        Index = Index + 1

                    Case "u"
                        Index = Index + 1
                        code = mid$(str, Index, 4)
                        SB.Append ChrW$(Val("&h" + code))
                        Index = Index + 4

                End Select

            Case quote
                Index = Index + 1
            
                parseString = SB.toString
                Set SB = Nothing
            
                Exit Function
            
            Case Else
                SB.Append Char
                Index = Index + 1

        End Select

    Loop
   
    parseString = SB.toString
    Set SB = Nothing
   
    
    Exit Function

parseString_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "parseString"
    End If
Resume Next
    
End Function

'
'   parse number
'
Private Function parseNumber(ByRef str As String, ByRef Index As Long)
    
    On Error GoTo parseNumber_Err
    

    Dim value As String
    Dim Char  As String

    Call skipChar(str, Index)

    Do While Index > 0 And Index <= Len(str)
        Char = mid$(str, Index, 1)

        If InStr("+-0123456789.eE", Char) Then
            value = value & Char
            Index = Index + 1
        Else
            parseNumber = CDec(value)
            Exit Function

        End If

    Loop

    
    Exit Function

parseNumber_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "parseNumber"
    End If
Resume Next
    
End Function

'
'   parse true / false
'
Private Function parseBoolean(ByRef str As String, ByRef Index As Long) As Boolean
    
    On Error GoTo parseBoolean_Err
    

    Call skipChar(str, Index)

    If mid$(str, Index, 4) = "true" Then
        parseBoolean = True
        Index = Index + 4
    ElseIf mid$(str, Index, 5) = "false" Then
        parseBoolean = False
        Index = Index + 5
    Else
        psErrors = psErrors & "Invalid Boolean at position " & Index & " : " & mid$(str, Index) & vbNewLine

    End If

    
    Exit Function

parseBoolean_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "parseBoolean"
    End If
Resume Next
    
End Function

'
<<<<<<< HEAD
'   parse null
'
Private Function parseNull(ByRef str As String, ByRef Index As Long)
    
    On Error GoTo parseNull_Err
    

    Call skipChar(str, Index)

    If mid$(str, Index, 4) = "null" Then
        parseNull = Null
        Index = Index + 4
    Else
        psErrors = psErrors & "Invalid null value at position " & Index & " : " & mid$(str, Index) & vbNewLine

    End If

    
    Exit Function

parseNull_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "parseNull"
    End If
Resume Next
    
End Function

Private Function parseKey(ByRef str As String, ByRef Index As Long) As String
    
    On Error GoTo parseKey_Err
    

    Dim dquote As Boolean
    Dim squote As Boolean
    Dim Char   As String

    Call skipChar(str, Index)

    Do While Index > 0 And Index <= Len(str)
        Char = mid$(str, Index, 1)

        Select Case (Char)

            Case """"
                dquote = Not dquote
                Index = Index + 1

                If Not dquote Then
                    Call skipChar(str, Index)

                    If mid$(str, Index, 1) <> ":" Then
                        psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbNewLine
                        Exit Do

                    End If

                End If

            Case "'"
                squote = Not squote
                Index = Index + 1

                If Not squote Then
                    Call skipChar(str, Index)

                    If mid$(str, Index, 1) <> ":" Then
                        psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbNewLine
                        Exit Do

                    End If

                End If

            Case ":"
                Index = Index + 1

                If Not dquote And Not squote Then
                    Exit Do
                Else
                    parseKey = parseKey & Char

                End If

            Case Else

                If InStr(vbNewLine & vbCr & vbLf & vbTab & " ", Char) Then
                Else
                    parseKey = parseKey & Char

                End If

                Index = Index + 1
=======
Private Function parseString(ByRef str As String, ByRef index As Long) As String
=======
'   parse string
'
Private Function parseString(ByRef str As String, ByRef Index As Long) As String
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'

   Dim quote   As String
   Dim Char    As String
   Dim code    As String

   Dim SB As New cStringBuilder

   Call skipChar(str, Index)
   quote = mid$(str, Index, 1)
   Index = Index + 1
   
   Do While Index > 0 And Index <= Len(str)
      Char = mid$(str, Index, 1)
      Select Case (Char)
         Case "\"
            Index = Index + 1
            Char = mid$(str, Index, 1)
            Select Case (Char)
               Case """", "\", "/", "'"
                  SB.Append Char
                  Index = Index + 1
               Case "b"
                  SB.Append vbBack
                  Index = Index + 1
               Case "f"
                  SB.Append vbFormFeed
                  Index = Index + 1
               Case "n"
                  SB.Append vbLf
                  Index = Index + 1
               Case "r"
                  SB.Append vbCr
                  Index = Index + 1
               Case "t"
                  SB.Append vbTab
                  Index = Index + 1
               Case "u"
                  Index = Index + 1
                  code = mid$(str, Index, 4)
                  SB.Append ChrW$(Val("&h" + code))
                  Index = Index + 4
            End Select
         Case quote
            Index = Index + 1
            
            parseString = SB.toString
            Set SB = Nothing
            
            Exit Function
            
         Case Else
            SB.Append Char
            Index = Index + 1
      End Select
   Loop
   
<<<<<<< HEAD
        charint = m_str(index)
      
        Select Case charint

            Case A_BACKSLASH

                index = index + 1
                charint = m_str(index)

                Select Case charint

                    Case A_DOUBLE_QUOTE, A_BACKSLASH, A_FORWARDSLASH, A_SINGLE_QUOTE
                        parseString = parseString & ChrW$(charint)
                        index = index + 1

                    Case A_b
                        parseString = parseString & vbBack
                        index = index + 1

                    Case A_f
                        parseString = parseString & vbFormFeed
                        index = index + 1

                    Case A_n
                        parseString = parseString & vbLf
                        index = index + 1

                    Case A_r
                        parseString = parseString & vbCr
                        index = index + 1

                    Case A_t
                        parseString = parseString & vbTab
                        index = index + 1

                    Case A_u
                        index = index + 1
                        Code = mid$(str, index, 4)

                        parseString = parseString & ChrW$(Val("&h" + Code))
                        index = index + 4

                End Select

            Case quoteint
        
                index = index + 1
                Exit Function

            Case Else
                parseString = parseString & ChrW$(charint)
                index = index + 1
>>>>>>> origin/master

        End Select

    Loop
<<<<<<< HEAD

    
    Exit Function

parseKey_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "parseKey"
    End If
Resume Next
    
End Function

'
'   skip special character
'
Private Sub skipChar(ByRef str As String, ByRef Index As Long)
    
    On Error GoTo skipChar_Err
    
    Dim bComment      As Boolean
    Dim bStartComment As Boolean
    Dim bLongComment  As Boolean

    Do While Index > 0 And Index <= Len(str)

        Select Case mid$(str, Index, 1)

            Case vbCr, vbLf

                If Not bLongComment Then
                    bStartComment = False
                    bComment = False

                End If
         
            Case vbTab, " ", "(", ")"
         
            Case "/"

                If Not bLongComment Then
                    If bStartComment Then
                        bStartComment = False
                        bComment = True
                    Else
                        bStartComment = True
                        bComment = False
                        bLongComment = False

                    End If

                Else

                    If bStartComment Then
                        bLongComment = False
                        bStartComment = False
                        bComment = False

                    End If

                End If
         
            Case "*"

                If bStartComment Then
                    bStartComment = False
                    bComment = True
                    bLongComment = True
                Else
                    bStartComment = True

                End If
         
            Case Else

                If Not bComment Then
                    Exit Do

                End If

        End Select
      
        Index = Index + 1
    Loop

    
    Exit Sub

skipChar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "skipChar"
    End If
Resume Next
    
End Sub

Public Function toString(ByRef obj As Variant) As String
    
    On Error GoTo toString_Err
    
    Dim SB As New cStringBuilder

    Select Case VarType(obj)

        Case vbNull
            SB.Append "null"

        Case vbDate
            SB.Append """" & CStr(obj) & """"

        Case vbString
            SB.Append """" & Encode(obj) & """"

        Case vbObject
         
            Dim bFI As Boolean
            Dim i   As Long
         
            bFI = True

            If TypeName(obj) = "Dictionary" Then

                SB.Append "{"
                Dim keys
                keys = obj.keys

                For i = 0 To obj.Count - 1

                    If bFI Then bFI = False Else SB.Append ","
                    Dim key
                    key = keys(i)
                    SB.Append """" & key & """:" & toString(obj.Item(key))
                Next i

                SB.Append "}"

            ElseIf TypeName(obj) = "Collection" Then

                SB.Append "["
                Dim value

                For Each value In obj

                    If bFI Then bFI = False Else SB.Append ","
                    SB.Append toString(value)
                Next value

                SB.Append "]"

            End If

        Case vbBoolean

            If obj Then SB.Append "true" Else SB.Append "false"

        Case vbVariant, vbArray, vbArray + vbVariant
            Dim sEB
            SB.Append multiArray(obj, 1, "", sEB)

        Case Else
            SB.Append Replace(obj, ",", ".")

    End Select

    toString = SB.toString
    Set SB = Nothing
=======
=======
   parseString = SB.toString
   Set SB = Nothing
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'
   
End Function

'
'   parse number
'
Private Function parseNumber(ByRef str As String, ByRef Index As Long)

   Dim Value   As String
   Dim Char    As String

   Call skipChar(str, Index)
   Do While Index > 0 And Index <= Len(str)
      Char = mid$(str, Index, 1)
      If InStr("+-0123456789.eE", Char) Then
         Value = Value & Char
         Index = Index + 1
      Else
         parseNumber = CDec(Value)
         Exit Function
      End If
   Loop
End Function

'
'   parse true / false
'
Private Function parseBoolean(ByRef str As String, ByRef Index As Long) As Boolean

   Call skipChar(str, Index)
   If mid$(str, Index, 4) = "true" Then
      parseBoolean = True
      Index = Index + 4
   ElseIf mid$(str, Index, 5) = "false" Then
      parseBoolean = False
      Index = Index + 5
   Else
      psErrors = psErrors & "Invalid Boolean at position " & Index & " : " & mid$(str, Index) & vbNewLine
   End If

End Function

'
'   parse null
'
Private Function parseNull(ByRef str As String, ByRef Index As Long)

   Call skipChar(str, Index)
   If mid$(str, Index, 4) = "null" Then
      parseNull = Null
      Index = Index + 4
   Else
      psErrors = psErrors & "Invalid null value at position " & Index & " : " & mid$(str, Index) & vbNewLine
   End If

End Function

Private Function parseKey(ByRef str As String, ByRef Index As Long) As String

   Dim dquote  As Boolean
   Dim squote  As Boolean
   Dim Char    As String

   Call skipChar(str, Index)
   Do While Index > 0 And Index <= Len(str)
      Char = mid$(str, Index, 1)
      Select Case (Char)
         Case """"
            dquote = Not dquote
            Index = Index + 1
            If Not dquote Then
               Call skipChar(str, Index)
               If mid$(str, Index, 1) <> ":" Then
                  psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbNewLine
                  Exit Do
               End If
            End If
         Case "'"
            squote = Not squote
            Index = Index + 1
            If Not squote Then
               Call skipChar(str, Index)
               If mid$(str, Index, 1) <> ":" Then
                  psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbNewLine
                  Exit Do
               End If
            End If
         Case ":"
            Index = Index + 1
            If Not dquote And Not squote Then
               Exit Do
            Else
               parseKey = parseKey & Char
            End If
         Case Else
            If InStr(vbNewLine & vbCr & vbLf & vbTab & " ", Char) Then
            Else
               parseKey = parseKey & Char
            End If
            Index = Index + 1
      End Select
   Loop

<<<<<<< HEAD
    Loop
>>>>>>> origin/master
   
    
    Exit Function

toString_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "toString"
    End If
Resume Next
    
End Function

<<<<<<< HEAD
Private Function Encode(str) As String
    
    On Error GoTo Encode_Err

    Dim SB      As New cStringBuilder
    Dim i   As Long
    Dim j   As Long
    Dim aL1 As Variant
    Dim aL2 As Variant
    Dim c   As String
    Dim p   As Boolean

    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)

    For i = 1 To Len(str)
        p = True
        c = mid$(str, i, 1)

        For j = 0 To 7

            If c = Chr$(aL1(j)) Then
                SB.Append "\" & Chr$(aL2(j))
                p = False
                Exit For

            End If

        Next

        If p Then
            Dim a
            a = AscW(c)

            If a > 31 And a < 127 Then
                SB.Append c
            ElseIf a > -1 Or a < 65535 Then
                SB.Append "\u" & String$(4 - LenB(Hex$(a)), "0") & Hex$(a)

            End If

        End If

    Next
   
    Encode = SB.toString
    Set SB = Nothing
    
    Exit Function

Encode_Err:

    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "Encode"

    End If

    Resume Next
    
End Function

Private Function multiArray(aBD, _
                            iBC, _
                            sPS, _
                            ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition
   
    Dim iDU As Long
    Dim iDL As Long
    Dim i   As Long
   
    On Error Resume Next

    iDL = LBound(aBD, iBC)
    iDU = UBound(aBD, iBC)

    Dim SB As New cStringBuilder

    Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2

    If Err.number = 9 Then
        sPB1 = sPT & sPS

        For i = 1 To Len(sPB1)

            If i <> 1 Then sPB2 = sPB2 & ","
            sPB2 = sPB2 & mid$(sPB1, i, 1)
        Next
        '        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
        SB.Append toString(aBD(sPB2))
    Else
        sPT = sPT & sPS
        SB.Append "["

        For i = iDL To iDU
            SB.Append multiArray(aBD, iBC + 1, i, sPT)

            If i < iDU Then SB.Append ","
        Next
        SB.Append "]"
        sPT = Left$(sPT, iBC - 2)

    End If

    Err.Clear
    multiArray = SB.toString
   
    Set SB = Nothing
=======
Private Function parseBoolean(ByRef str As String, ByRef index As Long) As Boolean

    Call skipChar(index)
   
    If mid$(str, index, 4) = "true" Then
        parseBoolean = True
        index = index + 4
    ElseIf mid$(str, index, 5) = "false" Then
        parseBoolean = False
        index = index + 5
    Else
        m_parserrors = m_parserrors & "Boolean invalido en la posicion " & index & " : " & mid$(str, index) & vbCrLf

    End If

End Function
=======
End Function

'
'   skip special character
'
Private Sub skipChar(ByRef str As String, ByRef Index As Long)
   Dim bComment As Boolean
   Dim bStartComment As Boolean
   Dim bLongComment As Boolean
   Do While Index > 0 And Index <= Len(str)
      Select Case mid$(str, Index, 1)
      Case vbCr, vbLf
         If Not bLongComment Then
            bStartComment = False
            bComment = False
         End If
         
      Case vbTab, " ", "(", ")"
         
      Case "/"
         If Not bLongComment Then
            If bStartComment Then
               bStartComment = False
               bComment = True
            Else
               bStartComment = True
               bComment = False
               bLongComment = False
            End If
         Else
            If bStartComment Then
               bLongComment = False
               bStartComment = False
               bComment = False
            End If
         End If
         
      Case "*"
         If bStartComment Then
            bStartComment = False
            bComment = True
            bLongComment = True
         Else
            bStartComment = True
         End If
         
      Case Else
         If Not bComment Then
            Exit Do
         End If
      End Select
      
      Index = Index + 1
   Loop
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'

End Sub

Public Function toString(ByRef obj As Variant) As String
   Dim SB As New cStringBuilder
   Select Case VarType(obj)
      Case vbNull
         SB.Append "null"
      Case vbDate
         SB.Append """" & CStr(obj) & """"
      Case vbString
         SB.Append """" & Encode(obj) & """"
      Case vbObject
         
         Dim bFI As Boolean
         Dim i As Long
         
         bFI = True
         If TypeName(obj) = "Dictionary" Then

            SB.Append "{"
            Dim keys
            keys = obj.keys
            For i = 0 To obj.Count - 1
               If bFI Then bFI = False Else SB.Append ","
               Dim key
               key = keys(i)
               SB.Append """" & key & """:" & toString(obj.Item(key))
            Next i
            SB.Append "}"

         ElseIf TypeName(obj) = "Collection" Then

            SB.Append "["
            Dim Value
            For Each Value In obj
               If bFI Then bFI = False Else SB.Append ","
               SB.Append toString(Value)
            Next Value
            SB.Append "]"

         End If
      Case vbBoolean
         If obj Then SB.Append "true" Else SB.Append "false"
      Case vbVariant, vbArray, vbArray + vbVariant
         Dim sEB
         SB.Append multiArray(obj, 1, "", sEB)
      Case Else
         SB.Append Replace(obj, ",", ".")
   End Select

   toString = SB.toString
   Set SB = Nothing
   
End Function

Private Function Encode(str) As String

   Dim SB As New cStringBuilder
   Dim i As Long
   Dim j As Long
   Dim aL1 As Variant
   Dim aL2 As Variant
   Dim c As String
   Dim p As Boolean

   aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
   aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
   For i = 1 To Len(str)
      p = True
      c = mid$(str, i, 1)
      For j = 0 To 7
         If c = Chr$(aL1(j)) Then
            SB.Append "\" & Chr$(aL2(j))
            p = False
            Exit For
         End If
      Next

      If p Then
         Dim a
         a = AscW(c)
         If a > 31 And a < 127 Then
            SB.Append c
         ElseIf a > -1 Or a < 65535 Then
            SB.Append "\u" & String(4 - Len(Hex$(a)), "0") & Hex$(a)
         End If
      End If
   Next
   
   Encode = SB.toString
   Set SB = Nothing
   
<<<<<<< HEAD
    Do While index > 0 And index <= m_length
    
        charint = m_str(index)
        
        Select Case charint

            Case A_DOUBLE_QUOTE
                dquote = Not dquote
                index = index + 1

                If Not dquote Then
            
                    Call skipChar(index)
                
                    If m_str(index) <> A_COLON Then
                        m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & index & " : " & parseKey & vbCrLf
                        Exit Do

                    End If

                End If

            Case A_SINGLE_QUOTE
                squote = Not squote
                index = index + 1

                If Not squote Then
                    Call skipChar(index)
                
                    If m_str(index) <> A_COLON Then
                        m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & index & " : " & parseKey & vbCrLf
                        Exit Do

                    End If
                
                End If
        
            Case A_COLON
                index = index + 1

                If Not dquote And Not squote Then
                    Exit Do
                Else
                    parseKey = parseKey & ChrW$(charint)

                End If

            Case Else
            
                If A_VBCRLF = charint Then
                ElseIf A_VBCR = charint Then
                ElseIf A_VBLF = charint Then
                ElseIf A_VBTAB = charint Then
                ElseIf A_SPACE = charint Then
                Else
                    parseKey = parseKey & ChrW$(charint)

                End If

                index = index + 1

        End Select

    Loop
>>>>>>> origin/master

End Function

Private Sub skipChar(ByRef index As Long)

<<<<<<< HEAD
Public Function StringToJSON(st As String) As String
    
    On Error GoTo StringToJSON_Err
    
   
    Const FIELD_SEP = "~"
    Const RECORD_SEP = "|"

    Dim sFlds   As String
    Dim sRecs   As New cStringBuilder
    Dim lRecCnt As Long
    Dim lFld    As Long
    Dim fld     As Variant
    Dim rows    As Variant

=======
    Dim bComment      As Boolean

    Dim bStartComment As Boolean

    Dim bLongComment  As Boolean

    Do While index > 0 And index <= m_length
    
        Select Case m_str(index)

            Case A_VBCR, A_VBLF

                If Not bLongComment Then
                    bStartComment = False
                    bComment = False

                End If
    
            Case A_VBTAB, A_SPACE, A_BRACKET_OPEN, A_BRACKET_CLOSE
                'do nothing
        
            Case A_FORWARDSLASH

                If Not bLongComment Then
                    If bStartComment Then
                        bStartComment = False
                        bComment = True
                    Else
                        bStartComment = True
                        bComment = False
                        bLongComment = False

                    End If

                Else

                    If bStartComment Then
                        bLongComment = False
                        bStartComment = False
                        bComment = False

                    End If

                End If

            Case A_ASTERIX

                If bStartComment Then
                    bStartComment = False
                    bComment = True
                    bLongComment = True
                Else
                    bStartComment = True

                End If

            Case Else
        
                If Not bComment Then
                    Exit Do

                End If

        End Select

        index = index + 1
    Loop

End Sub

Public Function GetRegionalSettings(ByVal regionalsetting As Long) As String
    ' Devuelve la configuracion regional del sistema

    On Error GoTo errorHandler

    Dim Locale      As Long

    Dim Symbol      As String

    Dim iRet1       As Long

    Dim iRet2       As Long

    Dim lpLCDataVar As String

    Dim Pos         As Integer
      
    Locale = GetUserDefaultLCID()

    iRet1 = GetLocaleInfo(Locale, regionalsetting, lpLCDataVar, 0)
    Symbol = String$(iRet1, 0)
    iRet2 = GetLocaleInfo(Locale, regionalsetting, Symbol, iRet1)
    Pos = InStr(Symbol, Chr$(0))

    If Pos > 0 Then
        Symbol = Left$(Symbol, Pos - 1)

    End If
      
errorHandler:
    GetRegionalSettings = Symbol

    Select Case Err.number

        Case 0

        Case Else
            Err.Raise 123, "GetRegionalSetting", "GetRegionalSetting: " & regionalsetting

    End Select

End Function

'********************************************************************************************************
'                   FUNCIONES MISCELANEAS DE LA ANTERIOR VERSION DEL MODULO
'********************************************************************************************************

Private Function Encode(str) As String

    Dim SB  As New cStringBuilder

    Dim i   As Long

    Dim j   As Long

    Dim aL1 As Variant

    Dim aL2 As Variant

    Dim c   As String

    Dim p   As Boolean

    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)

    For i = 1 To LenB(str)
        p = True
        c = mid$(str, i, 1)

        For j = 0 To 7

            If c = Chr$(aL1(j)) Then
                SB.Append "\" & Chr$(aL2(j))
                p = False
                Exit For

            End If

        Next

        If p Then

            Dim a

            a = AscW(c)

            If a > 31 And a < 127 Then
                SB.Append c
            ElseIf a > -1 Or a < 65535 Then
                SB.Append "\u" & String(4 - LenB(Hex$(a)), "0") & Hex$(a)

            End If

        End If

    Next
=======
End Function

Private Function multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition
   
   Dim iDU As Long
   Dim iDL As Long
   Dim i As Long
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'
   
   On Error Resume Next
   iDL = LBound(aBD, iBC)
   iDU = UBound(aBD, iBC)

   Dim SB As New cStringBuilder

   Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
   If Err.number = 9 Then
      sPB1 = sPT & sPS
      For i = 1 To Len(sPB1)
         If i <> 1 Then sPB2 = sPB2 & ","
         sPB2 = sPB2 & mid$(sPB1, i, 1)
      Next
      '        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
      SB.Append toString(aBD(sPB2))
   Else
      sPT = sPT & sPS
      SB.Append "["
      For i = iDL To iDU
         SB.Append multiArray(aBD, iBC + 1, i, sPT)
         If i < iDU Then SB.Append ","
      Next
      SB.Append "]"
      sPT = Left$(sPT, iBC - 2)
   End If
   Err.Clear
   multiArray = SB.toString
   
   Set SB = Nothing
End Function

' Miscellaneous JSON functions

Public Function StringToJSON(st As String) As String
   
<<<<<<< HEAD
    Const FIELD_SEP = "~"

    Const RECORD_SEP = "|"

    Dim sFlds   As String

    Dim sRecs   As New cStringBuilder

    Dim lRecCnt As Long

    Dim lFld    As Long

    Dim fld     As Variant

    Dim rows    As Variant

>>>>>>> origin/master
    lRecCnt = 0
=======
   Const FIELD_SEP = "~"
   Const RECORD_SEP = "|"

   Dim sFlds As String
   Dim sRecs As New cStringBuilder
   Dim lRecCnt As Long
   Dim lFld As Long
   Dim fld As Variant
   Dim rows As Variant

   lRecCnt = 0
   If LenB(st) = 0 Then
      StringToJSON = "null"
   Else
      rows = Split(st, RECORD_SEP)
      For lRecCnt = LBound(rows) To UBound(rows)
         sFlds = vbNullString
         fld = Split(rows(lRecCnt), FIELD_SEP)
         For lFld = LBound(fld) To UBound(fld) Step 2
            sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
         Next 'fld
         sRecs.Append IIf((Trim$(sRecs.toString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}"
      Next 'rec
      StringToJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")
   End If
End Function
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'


Public Function RStoJSON(rs As ADODB.Recordset) As String
   On Error GoTo ErrHandler
   Dim sFlds As String
   Dim sRecs As New cStringBuilder
   Dim lRecCnt As Long
   Dim fld As ADODB.Field

   lRecCnt = 0
   If rs.State = adStateClosed Then
      RStoJSON = "null"
   Else
      If rs.EOF Or rs.BOF Then
         RStoJSON = "null"
      Else
         Do While Not rs.EOF And Not rs.BOF
            lRecCnt = lRecCnt + 1
            sFlds = vbNullString
            For Each fld In rs.Fields
               sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.Name & """:""" & toUnicode(fld.Value & "") & """")
            Next 'fld
            sRecs.Append IIf((Trim$(sRecs.toString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}"
<<<<<<< HEAD
        Next 'rec
<<<<<<< HEAD

        StringToJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")

    End If

    
    Exit Function

StringToJSON_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "StringToJSON"
    End If
Resume Next
    
End Function

Public Function RStoJSON(rs As ADODB.Recordset) As String

    On Error GoTo ErrHandler

    Dim sFlds   As String
    Dim sRecs   As New cStringBuilder
    Dim lRecCnt As Long
    Dim fld     As ADODB.Field

    lRecCnt = 0

    If rs.State = adStateClosed Then
        RStoJSON = "null"
    Else

        If rs.EOF Or rs.BOF Then
            RStoJSON = "null"
        Else

            Do While Not rs.EOF And Not rs.BOF
                lRecCnt = lRecCnt + 1
                sFlds = vbNullString

                For Each fld In rs.Fields

                    sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.Name & """:""" & toUnicode(fld.value & "") & """")
                Next 'fld

                sRecs.Append IIf((Trim$(sRecs.toString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}"
                rs.MoveNext
            Loop
            RStoJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")

        End If

    End If

    Exit Function
ErrHandler:
=======

        StringToJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")

    End If
>>>>>>> origin/master

End Function

Public Function RStoJSON(rs As ADODB.Recordset) As String

<<<<<<< HEAD
Public Function toUnicode(str As String) As String
    
    On Error GoTo toUnicode_Err
    

    Dim X        As Long
    Dim uStr     As New cStringBuilder
    Dim uChrCode As Integer

    For X = 1 To Len(str)
        uChrCode = Asc(mid$(str, X, 1))

        Select Case uChrCode

            Case 8:   ' backspace
                uStr.Append "\b"

            Case 9: ' tab
                uStr.Append "\t"
=======
    On Error GoTo ErrHandler

    Dim sFlds   As String

    Dim sRecs   As New cStringBuilder

    Dim lRecCnt As Long

    Dim fld     As ADODB.Field

    lRecCnt = 0

    If rs.State = adStateClosed Then
        RStoJSON = "null"
    Else

        If rs.EOF Or rs.BOF Then
            RStoJSON = "null"
        Else

            Do While Not rs.EOF And Not rs.BOF
                lRecCnt = lRecCnt + 1
                sFlds = vbNullString

                For Each fld In rs.Fields

                    sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.Name & """:""" & toUnicode(fld.Value & "") & """")
                Next 'fld

                sRecs.Append IIf((Trim$(sRecs.toString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}"
                rs.MoveNext
            Loop
            RStoJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")

        End If

    End If

    Exit Function
=======
            rs.MoveNext
         Loop
         RStoJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")
      End If
   End If

   Exit Function
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'
ErrHandler:
>>>>>>> origin/master

            Case 10:  ' line feed
                uStr.Append "\n"

            Case 12:  ' formfeed
                uStr.Append "\f"

            Case 13: ' carriage return
                uStr.Append "\r"

            Case 34: ' quote
                uStr.Append "\"""

            Case 39:  ' apostrophe
                uStr.Append "\'"

            Case 92: ' backslash
                uStr.Append "\\"

            Case 123, 125:  ' "{" and "}"
                uStr.Append ("\u" & Right$("0000" & Hex$(uChrCode), 4))

            Case Is < 32, Is > 127: ' non-ascii characters
                uStr.Append ("\u" & Right$("0000" & Hex$(uChrCode), 4))

            Case Else
                uStr.Append Chr$(uChrCode)

        End Select

    Next
    toUnicode = uStr.toString
    Exit Function

    
    Exit Function

toUnicode_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "toUnicode"
    End If
Resume Next
    
End Function

<<<<<<< HEAD
<<<<<<< HEAD
Private Sub Class_Initialize()
    
    On Error GoTo Class_Initialize_Err
    
    psErrors = vbNullString

    
    Exit Sub

Class_Initialize_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "JSON" & "->" & "Class_Initialize"
    End If
Resume Next
    
End Sub

=======
Public Function toUnicode(str As String) As String
=======
'Public Function JsonRpcCall(url As String, methName As String, args(), Optional user As String, Optional pwd As String) As Object
'    Dim r As Object
'    Dim cli As Object
'    Dim pText As String
'    Static reqId As Integer
'
'    reqId = reqId + 1
'
'    Set r = CreateObject("Scripting.Dictionary")
'    r("jsonrpc") = "2.0"
'    r("method") = methName
'    r("params") = args
'    r("id") = reqId
'
'    pText = toString(r)
'
'    Set cli = CreateObject("MSXML2.XMLHTTP.6.0")
'   ' Set cli = New MSXML2.XMLHTTP60
'    If Len(user) > 0 Then   ' If Not IsMissing(user) Then
'        cli.Open "POST", url, False, user, pwd
'    Else
'        cli.Open "POST", url, False
'    End If
'    cli.setRequestHeader "Content-Type", "application/json"
'    cli.Send pText
'
'    If cli.Status <> 200 Then
'        Err.Raise vbObjectError + INVALID_RPC_CALL + cli.Status, , cli.statusText
'    End If
'
'    Set r = parse(cli.responseText)
'    Set cli = Nothing
'
'    If r("id") <> reqId Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response id"
'
'    If r.Exists("error") Or Not r.Exists("result") Then
'        Err.Raise vbObjectError + INVALID_RPC_CALL, , "Json-Rpc Response error: " & r("error")("message")
'    End If
'
'    If Not r.Exists("result") Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response, missing result"
'
'    Set JsonRpcCall = r("result")
'End Function
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'




Public Function toUnicode(str As String) As String

   Dim X As Long
   Dim uStr As New cStringBuilder
   Dim uChrCode As Integer

   For X = 1 To Len(str)
      uChrCode = Asc(mid$(str, X, 1))
      Select Case uChrCode
         Case 8:   ' backspace
            uStr.Append "\b"
         Case 9: ' tab
            uStr.Append "\t"
         Case 10:  ' line feed
            uStr.Append "\n"
         Case 12:  ' formfeed
            uStr.Append "\f"
         Case 13: ' carriage return
            uStr.Append "\r"
         Case 34: ' quote
            uStr.Append "\"""
         Case 39:  ' apostrophe
            uStr.Append "\'"
         Case 92: ' backslash
            uStr.Append "\\"
         Case 123, 125:  ' "{" and "}"
            uStr.Append ("\u" & Right$("0000" & Hex$(uChrCode), 4))
         Case Is < 32, Is > 127: ' non-ascii characters
            uStr.Append ("\u" & Right$("0000" & Hex$(uChrCode), 4))
         Case Else
            uStr.Append Chr$(uChrCode)
      End Select
   Next
   toUnicode = uStr.toString
   Exit Function

End Function

Private Sub Class_Initialize()
   psErrors = vbNullString
End Sub


<<<<<<< HEAD
End Function
>>>>>>> origin/master
=======
>>>>>>> parent of 2e97656... Merge remote-tracking branch 'origin/nuevo-modulo-JSON'
