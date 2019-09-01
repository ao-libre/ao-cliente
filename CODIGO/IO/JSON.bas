Attribute VB_Name = "JSON"

' VBJSONDeserializer is a VB6 adaptation of the VB-JSON project @
' Fuente: https://www.codeproject.com/Articles/720368/VB-JSON-Parser-Improved-Performance

' BSD Licensed

Option Explicit

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

Public Function GetParserErrors() As String
    GetParserErrors = m_parserrors
End Function

Public Function parse(ByRef str As String) As Object

    m_decSep = GetRegionalSettings(LOCALE_SDECIMAL)
    m_groupSep = GetRegionalSettings(LOCALE_SGROUPING)

    Dim Index As Long

    Index = 1

    GenerateStringArray str

    m_parserrors = vbNullString

    On Error Resume Next

    Call skipChar(Index)

    Select Case m_str(Index)

        Case A_SQUARE_BRACKET_OPEN
            Set parse = parseArray(str, Index)

        Case A_CURLY_BRACKET_OPEN
            Set parse = parseObject(str, Index)

        Case Else
            m_parserrors = "JSON Invalido"

    End Select

    'clean array
    ReDim m_str(1)

End Function

Private Sub GenerateStringArray(ByRef str As String)

    Dim i As Long

    m_length = Len(str)
    ReDim m_str(1 To m_length)

    For i = 1 To m_length
        m_str(i) = AscW(mid$(str, i, 1))
    Next i

End Sub

Private Function parseObject(ByRef str As String, ByRef Index As Long) As Dictionary

    Set parseObject = New Dictionary

    Dim sKey    As String

    Dim charint As Integer

    Call skipChar(Index)

    If m_str(Index) <> A_CURLY_BRACKET_OPEN Then
        m_parserrors = m_parserrors & "Objeto invalido en la posicion " & Index & " : " & mid$(str, Index) & vbCrLf
        Exit Function

    End If

    Index = Index + 1

    Do
        Call skipChar(Index)
    
        charint = m_str(Index)
    
        If charint = A_COMMA Then
            Index = Index + 1
            Call skipChar(Index)
        ElseIf charint = A_CURLY_BRACKET_CLOSE Then
            Index = Index + 1
            Exit Do
        ElseIf Index > m_length Then
            m_parserrors = m_parserrors & "Falta '}': " & Right$(str, 20) & vbCrLf
            Exit Do
        End If

        ' add key/value pair
        sKey = parseKey(Index)

        On Error Resume Next

        parseObject.Add sKey, parseValue(str, Index)

        If Err.number <> 0 Then
            m_parserrors = m_parserrors & Err.Description & ": " & sKey & vbCrLf
            Exit Do
        End If

    Loop

End Function

Private Function parseArray(ByRef str As String, ByRef Index As Long) As Collection

    Dim charint As Integer

    Set parseArray = New Collection

    Call skipChar(Index)

    If mid$(str, Index, 1) <> "[" Then
        m_parserrors = m_parserrors & "Array invalido en la posicion " & Index & " : " + mid$(str, Index, 20) & vbCrLf
        Exit Function
    End If
   
    Index = Index + 1

    Do
        Call skipChar(Index)
    
        charint = m_str(Index)
    
        If charint = A_SQUARE_BRACKET_CLOSE Then
            Index = Index + 1
            Exit Do
        ElseIf charint = A_COMMA Then
            Index = Index + 1
            Call skipChar(Index)
        ElseIf Index > m_length Then
            m_parserrors = m_parserrors & "Falta ']': " & Right$(str, 20) & vbCrLf
            Exit Do
        End If
    
        'add value
        On Error Resume Next

        parseArray.Add parseValue(str, Index)

        If Err.number <> 0 Then
            m_parserrors = m_parserrors & Err.Description & ": " & mid$(str, Index, 20) & vbCrLf
            Exit Do

        End If

    Loop

End Function

Private Function parseValue(ByRef str As String, ByRef Index As Long)

    Call skipChar(Index)

    Select Case m_str(Index)

        Case A_DOUBLE_QUOTE, A_SINGLE_QUOTE
            parseValue = parseString(str, Index)
            Exit Function

        Case A_SQUARE_BRACKET_OPEN
            Set parseValue = parseArray(str, Index)
            Exit Function

        Case A_t, A_f
            parseValue = parseBoolean(str, Index)
            Exit Function

        Case A_n
            parseValue = parseNull(str, Index)
            Exit Function

        Case A_CURLY_BRACKET_OPEN
            Set parseValue = parseObject(str, Index)
            Exit Function

        Case Else
            parseValue = parseNumber(str, Index)
            Exit Function

    End Select

End Function

Private Function parseString(ByRef str As String, ByRef Index As Long) As String

    Dim quoteint As Integer

    Dim charint  As Integer

    Dim Code     As String
   
    Call skipChar(Index)
   
    quoteint = m_str(Index)
   
    Index = Index + 1
   
    Do While Index > 0 And Index <= m_length
   
        charint = m_str(Index)
      
        Select Case charint

            Case A_BACKSLASH

                Index = Index + 1
                charint = m_str(Index)

                Select Case charint

                    Case A_DOUBLE_QUOTE, A_BACKSLASH, A_FORWARDSLASH, A_SINGLE_QUOTE
                        parseString = parseString & ChrW$(charint)
                        Index = Index + 1

                    Case A_b
                        parseString = parseString & vbBack
                        Index = Index + 1

                    Case A_f
                        parseString = parseString & vbFormFeed
                        Index = Index + 1

                    Case A_n
                        parseString = parseString & vbLf
                        Index = Index + 1

                    Case A_r
                        parseString = parseString & vbCr
                        Index = Index + 1

                    Case A_t
                        parseString = parseString & vbTab
                        Index = Index + 1

                    Case A_u
                        Index = Index + 1
                        Code = mid$(str, Index, 4)

                        parseString = parseString & ChrW$(Val("&h" + Code))
                        Index = Index + 4

                End Select

            Case quoteint
        
                Index = Index + 1
                Exit Function

            Case Else
                parseString = parseString & ChrW$(charint)
                Index = Index + 1

        End Select

    Loop
   
End Function

Private Function parseNumber(ByRef str As String, ByRef Index As Long)

    Dim Value As String

    Dim Char  As String

    Call skipChar(Index)

    Do While Index > 0 And Index <= m_length
        Char = mid$(str, Index, 1)

        If InStr("+-0123456789.eE", Char) Then
            Value = Value & Char
            Index = Index + 1
        Else

            'check what is the grouping seperator
            If Not m_decSep = "." Then
                Value = Replace(Value, ".", m_decSep)

            End If
     
            If m_groupSep = "." Then
                Value = Replace(Value, ".", m_decSep)

            End If
     
            parseNumber = CDec(Value)
            Exit Function

        End If

    Loop
   
End Function

Private Function parseBoolean(ByRef str As String, ByRef Index As Long) As Boolean

    Call skipChar(Index)
   
    If mid$(str, Index, 4) = "true" Then
        parseBoolean = True
        Index = Index + 4
    ElseIf mid$(str, Index, 5) = "false" Then
        parseBoolean = False
        Index = Index + 5
    Else
        m_parserrors = m_parserrors & "Boolean invalido en la posicion " & Index & " : " & mid$(str, Index) & vbCrLf

    End If

End Function

Private Function parseNull(ByRef str As String, ByRef Index As Long)

    Call skipChar(Index)
   
    If mid$(str, Index, 4) = "null" Then
        parseNull = Null
        Index = Index + 4
    Else
        m_parserrors = m_parserrors & "Valor nulo invalido en la posicion " & Index & " : " & mid$(str, Index) & vbCrLf

    End If

End Function

Private Function parseKey(ByRef Index As Long) As String

    Dim dquote  As Boolean

    Dim squote  As Boolean

    Dim charint As Integer
   
    Call skipChar(Index)
   
    Do While Index > 0 And Index <= m_length
    
        charint = m_str(Index)
        
        Select Case charint

            Case A_DOUBLE_QUOTE
                dquote = Not dquote
                Index = Index + 1

                If Not dquote Then
            
                    Call skipChar(Index)
                
                    If m_str(Index) <> A_COLON Then
                        m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & Index & " : " & parseKey & vbCrLf
                        Exit Do

                    End If

                End If

            Case A_SINGLE_QUOTE
                squote = Not squote
                Index = Index + 1

                If Not squote Then
                    Call skipChar(Index)
                
                    If m_str(Index) <> A_COLON Then
                        m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & Index & " : " & parseKey & vbCrLf
                        Exit Do

                    End If
                
                End If
        
            Case A_COLON
                Index = Index + 1

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

                Index = Index + 1

        End Select

    Loop

End Function

Private Sub skipChar(ByRef Index As Long)

    Dim bComment      As Boolean

    Dim bStartComment As Boolean

    Dim bLongComment  As Boolean

    Do While Index > 0 And Index <= m_length
    
        Select Case m_str(Index)

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

        Index = Index + 1
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

Private Function Encode(ByVal str As String) As String

    Dim SB  As New cStringBuilder
    Dim i   As Long
    Dim j   As Long
    Dim aL1 As Variant
    Dim aL2 As Variant
    Dim c   As String
    Dim p   As Boolean
    Dim Len_str As Long

    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
    
    Len_str = LenB(str)
    
    For i = 1 To Len_str
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

            Dim A As Integer: A = AscW(c)

            If A > 31 And A < 127 Then
                SB.Append c
            ElseIf A > -1 Or A < 65535 Then
                SB.Append "\u" & String$(4 - LenB(Hex$(A)), "0") & Hex$(A)

            End If

        End If

    Next
   
    Encode = SB.toString
    Set SB = Nothing
   
End Function

Public Function StringToJSON(st As String) As String
   
    Const FIELD_SEP = "~"

    Const RECORD_SEP = "|"

    Dim sFlds   As String

    Dim sRecs   As New cStringBuilder

    Dim lRecCnt As Long

    Dim lFld    As Long

    Dim fld     As Variant

    Dim rows    As Variant
    
    Dim Lower_rows As Long, Upper_rows As Long
    Dim Lower_fld As Long, Upper_fld As Long

    lRecCnt = 0

    If LenB(st) = 0 Then
        StringToJSON = "null"
    Else
        rows = Split(st, RECORD_SEP)
        
        Lower_rows = LBound(rows)
        Upper_rows = UBound(rows)
        
        For lRecCnt = Lower_rows To Upper_rows
            sFlds = vbNullString
            fld = Split(rows(lRecCnt), FIELD_SEP)
            
            Lower_fld = LBound(fld)
            Upper_fld = UBound(fld)
            
            For lFld = Lower_fld To Upper_fld Step 2
                sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
            Next 'fld

            sRecs.Append IIf((Trim$(sRecs.toString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}"
        Next 'rec

        StringToJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")

    End If
End Function

Public Function RStoJSON(rs As ADODB.Recordset) As String

    On Error GoTo errhandler

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
errhandler:

End Function

Public Function toUnicode(str As String) As String

    Dim X        As Long

    Dim uStr     As New cStringBuilder

    Dim uChrCode As Integer
    
    Dim Len_str As Long
        Len_str = LenB(str)

    For X = 1 To Len_str
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
