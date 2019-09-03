VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*** DECLARATIONS FOUND IN WININET.H ***
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetQueryDataAvailable Lib "wininet.dll" (ByVal hHttpRequest As Long, ByRef lpdwNumberOfBytesAvailable As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" (ByVal hHttpRequest As Long, ByVal dwOption As Long, lpBuffer As Any, ByRef dwBufferLength As Long) As Long
Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hHttpRequest As Long, ByVal dwOption As Long, lpBuffer As Any, ByVal dwBufferLength As Long) As Long
Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" (ByVal hInternetSession As Long, ByVal lpHeaders As String, ByVal lHeadersLength As Long, ByVal lModifiers As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal sAcceptTypes As String, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Long '**********************************************

'**********************************************
' Constants for InternetOpen()
'**********************************************
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
'**********************************************
' Constants for InternetConnect()
'**********************************************
Private Const INTERNET_DEFAULT_HTTP_PORT = 80
Private Const INTERNET_DEFAULT_HTTPS_PORT = 443
Private Const INTERNET_SERVICE_HTTP = 3
'**********************************************
' Constants for HttpOpenRequest()
'**********************************************
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100
Private Const INTERNET_FLAG_NO_COOKIES = &H80000
Private Const INTERNET_FLAG_SECURE = &H800000
'**********************************************
' Constants for InternetQueryOption() or InternetSetOption()
'**********************************************
Private Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Private Const INTERNET_OPTION_SEND_TIMEOUT = 5
Private Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Private Const INTERNET_OPTION_SECURITY_FLAGS = 31
Private Const SECURITY_FLAG_IGNORE_UNKNOWN_CA = &H100
'**********************************************
' Constants for HttpAddRequestHeaders()
'**********************************************
Private Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000
Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000
Private Const HTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA = &H40000000
Private Const HTTP_ADDREQ_FLAG_COALESCE_WITH_SEMICOLON = &H1000000
Private Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000
Private Const ERROR_INSUFFICIENT_BUFFER = 122
 
Public Enum enumPostHeader
    HTTP_QUERY_ACCEPT = 24
    HTTP_QUERY_ACCEPT_CHARSET = 25
    HTTP_QUERY_ACCEPT_ENCODING = 26
    HTTP_QUERY_ACCEPT_LANGUAGE = 27
    HTTP_QUERY_ACCEPT_RANGES = 42
    HTTP_QUERY_AGE = 48
    HTTP_QUERY_ALLOW = 7
    HTTP_QUERY_AUTHORIZATION = 28
    HTTP_QUERY_CACHE_CONTROL = 49
    HTTP_QUERY_CONNECTION = 23
    HTTP_QUERY_CONTENT_BASE = 50
    HTTP_QUERY_CONTENT_DESCRIPTION = 4
    HTTP_QUERY_CONTENT_DISPOSITION = 47
    HTTP_QUERY_CONTENT_ENCODING = 29
    HTTP_QUERY_CONTENT_ID = 3
    HTTP_QUERY_CONTENT_LANGUAGE = 6
    HTTP_QUERY_CONTENT_LENGTH = 5
    HTTP_QUERY_CONTENT_LOCATION = 51
    HTTP_QUERY_CONTENT_MD5 = 52
    HTTP_QUERY_CONTENT_RANGE = 53
    HTTP_QUERY_CONTENT_TRANSFER_ENCODING = 2
    HTTP_QUERY_CONTENT_TYPE = 1
    HTTP_QUERY_COOKIE = 44
    HTTP_QUERY_COST = 15
    HTTP_QUERY_CUSTOM = 65535
    HTTP_QUERY_DATE = 9
    HTTP_QUERY_DERIVED_FROM = 14
    HTTP_QUERY_ECHO_HEADERS = 73
    HTTP_QUERY_ECHO_HEADERS_CRLF = 74
    HTTP_QUERY_ECHO_REPLY = 72
    HTTP_QUERY_ECHO_REQUEST = 71
    HTTP_QUERY_ETAG = 54
    HTTP_QUERY_EXPECT = 68
    HTTP_QUERY_EXPIRES = 10
    HTTP_QUERY_FORWARDED = 30
    HTTP_QUERY_FROM = 31
    HTTP_QUERY_HOST = 55
    HTTP_QUERY_IF_MATCH = 56
    HTTP_QUERY_IF_MODIFIED_SINCE = 32
    HTTP_QUERY_IF_NONE_MATCH = 57
    HTTP_QUERY_IF_RANGE = 58
    HTTP_QUERY_IF_UNMODIFIED_SINCE = 59
    HTTP_QUERY_LAST_MODIFIED = 11
    HTTP_QUERY_LINK = 16
    HTTP_QUERY_LOCATION = 33
    HTTP_QUERY_MAX = 78
    HTTP_QUERY_MAX_FORWARDS = 60
    HTTP_QUERY_MESSAGE_ID = 12
    HTTP_QUERY_MIME_VERSION = 0
    HTTP_QUERY_ORIG_URI = 34
    HTTP_QUERY_PRAGMA = 17
    HTTP_QUERY_PROXY_AUTHENTICATE = 41
    HTTP_QUERY_PROXY_AUTHORIZATION = 61
    HTTP_QUERY_PROXY_CONNECTION = 69
    HTTP_QUERY_PUBLIC = 8
    HTTP_QUERY_RANGE = 62
    HTTP_QUERY_RAW_HEADERS = 21
    HTTP_QUERY_RAW_HEADERS_CRLF = 22
    HTTP_QUERY_REFERER = 35
    HTTP_QUERY_REFRESH = 46
    HTTP_QUERY_REQUEST_METHOD = 45
    HTTP_QUERY_RETRY_AFTER = 36
    HTTP_QUERY_SERVER = 37
    HTTP_QUERY_SET_COOKIE = 43
    HTTP_QUERY_STATUS_CODE = 19
    HTTP_QUERY_STATUS_TEXT = 20
    HTTP_QUERY_TITLE = 38
    HTTP_QUERY_TRANSFER_ENCODING = 63
    HTTP_QUERY_UNLESS_MODIFIED_SINCE = 70
    HTTP_QUERY_UPGRADE = 64
    HTTP_QUERY_URI = 13
    HTTP_QUERY_USER_AGENT = 39
    HTTP_QUERY_VARY = 65
    HTTP_QUERY_VERSION = 18
    HTTP_QUERY_VIA = 66
    HTTP_QUERY_WARNING = 67
    HTTP_QUERY_WWW_AUTHENTICATE = 40
End Enum
 
Private Const CONNECTION_TIMEOUT_MILLIS As Long = 60000
 
Private m_sURLDomain                    As String
Private m_sURLPort                      As String
Private m_sURLPath                      As String
Private m_bURLSSL                       As Boolean
Private m_sPOSTData                     As String
Private m_hInetSession                  As Long
Private m_hInetConnect                  As Long
Private m_hInetRequest                  As Long
Private m_sHttpMethod                   As String
 
Public Property Let PostParameters(ByVal sData As String)
    m_sPOSTData = sData
End Property
 
Public Property Get PostParameters() As String
    PostParameters = m_sPOSTData
End Property
 
Public Function OpenRequest(ByVal sURL As String, _
                            Optional ByVal sHttpMethod As String = "POST", _
                            Optional ByVal sUserName As String = vbNullString, _
                            Optional ByVal sPassword As String = vbNullString, _
                            Optional ByRef sErrorMsg As Variant) As Boolean
 
    m_sPOSTData = ""
    m_sHttpMethod = sHttpMethod

    If Not CrackURL(sURL) Then Exit Function
 
    m_hInetSession = InternetOpen("Http Client", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

    If m_hInetSession <= 0 Then
        sErrorMsg = "InternetOpen() failed (System Error " & Err.LastDllError & ")"
        GoTo EXIT_LABEL
    End If
 
    m_hInetConnect = InternetConnect(m_hInetSession, m_sURLDomain, m_sURLPort, sUserName, sPassword, INTERNET_SERVICE_HTTP, 0, 0)

    If m_hInetSession <= 0 Then
        sErrorMsg = "InternetConnect() failed (System Error " & Err.LastDllError & ")"
        GoTo EXIT_LABEL
    End If
 
    Dim lFlags As Long
    lFlags = (INTERNET_FLAG_NO_COOKIES Or INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE)

    If m_bURLSSL Then lFlags = (lFlags Or INTERNET_FLAG_SECURE)
 
    m_hInetRequest = HttpOpenRequest(m_hInetConnect, sHttpMethod, m_sURLPath, "HTTP/1.0", vbNullString, vbNullString, lFlags, 0)

    If m_hInetRequest <= 0 Then
        sErrorMsg = "HttpOpenRequest() failed (System Error " & Err.LastDllError & ")"
        GoTo EXIT_LABEL
    End If
    
    If Not CBool(InternetSetOption(m_hInetRequest, INTERNET_OPTION_CONNECT_TIMEOUT, CONNECTION_TIMEOUT_MILLIS, 4)) Then
        sErrorMsg = "InternetSetOption(INTERNET_OPTION_CONNECT_TIMEOUT) failed (System Error " & Err.LastDllError & ")"
        GoTo EXIT_LABEL
    End If
    
    If Not CBool(InternetSetOption(m_hInetRequest, INTERNET_OPTION_SEND_TIMEOUT, CONNECTION_TIMEOUT_MILLIS, 4)) Then
        sErrorMsg = "InternetSetOption(INTERNET_OPTION_SEND_TIMEOUT) failed (System Error " & Err.LastDllError & ")"
        GoTo EXIT_LABEL
    End If
    
    If Not CBool(InternetSetOption(m_hInetRequest, INTERNET_OPTION_RECEIVE_TIMEOUT, CONNECTION_TIMEOUT_MILLIS, 4)) Then
        sErrorMsg = "InternetSetOption(INTERNET_OPTION_RECEIVE_TIMEOUT) failed (System Error " & Err.LastDllError & ")"
        GoTo EXIT_LABEL
    End If
    
    If sHttpMethod = "POST" Then
        If Not AddHeader("Content-Type", "application/x-www-form-urlencoded") Then
            sErrorMsg = "AddHeader('Content-Type') failed"
            GoTo EXIT_LABEL
        End If
    End If
    
    If m_bURLSSL Then
        Dim lBuffer As Long
        Call InternetQueryOption(m_hInetRequest, INTERNET_OPTION_SECURITY_FLAGS, lFlags, lBuffer)
        lFlags = lFlags Or SECURITY_FLAG_IGNORE_UNKNOWN_CA
        Call InternetSetOption(m_hInetRequest, INTERNET_OPTION_SECURITY_FLAGS, lFlags, lBuffer)
    End If
    
    OpenRequest = True
    
EXIT_LABEL:

End Function
    
Public Function AddHeader(ByVal sName As String, ByVal sValue As String) As Boolean
    Dim sHeaderFmt As String
    
    sHeaderFmt = Trim$(sName) & ": " & Trim$(sValue) & vbCrLf
    AddHeader = CBool(HttpAddRequestHeaders(m_hInetRequest, sHeaderFmt, Len(sHeaderFmt), HTTP_ADDREQ_FLAG_ADD Or HTTP_ADDREQ_FLAG_REPLACE))
    
End Function
    
Public Function AddPOSTParameter(ByVal Name As String, ByVal Value As String) As Boolean

    Dim sParam As String

    Name = Trim$(Name)
    Value = Trim$(Value)

    If m_sHttpMethod = "POST" Then Value = UrlEncode(Value)

    If Len(Name) > 0 Then
        sParam = Name & "=" & Value
    Else
        sParam = Value
    End If
 
    If Len(m_sPOSTData) = 0 Then
        m_sPOSTData = sParam
    Else
        m_sPOSTData = m_sPOSTData & "&" & sParam
    End If
    
    AddPOSTParameter = True
End Function
 
Public Function Execute() As Boolean

    If Len(m_sPOSTData) > 0 Then
        Call AddHeader("Content-Length", Len(m_sPOSTData))
        Execute = CBool(HttpSendRequest(m_hInetRequest, vbNullString, 0, m_sPOSTData, Len(m_sPOSTData)))
    Else
        Execute = CBool(HttpSendRequest(m_hInetRequest, vbNullString, 0, vbNullString, 0))
    End If
End Function
 
Public Function GetResponseAsString() As String
    On Error GoTo errhandler
    Dim sBody      As String
    Dim sBuffer    As String
    Dim lBytesRead As Long
    Do
        sBuffer = Space$(4048)
        InternetReadFile m_hInetRequest, sBuffer, Len(sBuffer), lBytesRead

        If lBytesRead <= 0 Then Exit Do
        sBody = sBody & Left$(sBuffer, lBytesRead)
    Loop
    GetResponseAsString = sBody
    Exit Function

errhandler:

    GetResponseAsString = ""
End Function

Public Function GetResponseHeader(ByVal lHeaderType As enumPostHeader) As String
    On Error GoTo errhandler
    Dim sValue       As String
    Dim lBytesRead   As Long
    Dim lHeaderIndex As Long
    Dim bSuccess     As Boolean
    
    sValue = String$(2000, Chr$(0))
    lBytesRead = Len(sValue)
    bSuccess = CBool(HttpQueryInfo(m_hInetRequest, lHeaderType, ByVal sValue, lBytesRead, lHeaderIndex))

    If bSuccess And lBytesRead > 0 Then
        GetResponseHeader = Left$(sValue, InStr(1, sValue, Chr$(0)) - 1)
    End If
    Exit Function
 
errhandler:
End Function
 
Private Function CrackURL(ByVal strURL As String) As Boolean
    Dim intPos      As Integer
    Dim strProtocol As String
 
    m_sURLDomain = vbNullString
    m_sURLPort = vbNullString
    m_sURLPath = vbNullString
    m_bURLSSL = False
 
    ' search for double forward slash and remove protocol prefix
    intPos = InStr(1, strURL, "//")

    If intPos > 0 Then
        strProtocol = Left$(strURL, intPos - 1)
        strURL = mid$(strURL, intPos + 2)
    End If
 
    m_bURLSSL = (InStr(LCase$(strProtocol), "https") > 0)
 
    ' split the URL into domain and path using the first forward slash found
    ' If no slash is found, then the url is the domain only
    intPos = InStr(1, strURL, "/")

    If intPos > 0 Then
        m_sURLDomain = Left$(strURL, intPos - 1)
        m_sURLPath = mid$(strURL, intPos + 1)
    Else
        m_sURLDomain = strURL
    End If
 
    ' Now see if a port is attached to the domain; split if found
    intPos = InStr(1, m_sURLDomain, ":")

    If intPos > 0 Then
        m_sURLPort = mid$(m_sURLDomain, intPos + 1)
        m_sURLDomain = Left$(m_sURLDomain, intPos - 1)
    Else
        m_sURLPort = IIf(m_bURLSSL, INTERNET_DEFAULT_HTTPS_PORT, INTERNET_DEFAULT_HTTP_PORT)
    End If
    CrackURL = True
End Function
 
Private Function TrimString(ByVal s As String) As String
    TrimString = Left$(s, InStr(1, s, Chr$(0)) - 1)
End Function
 
Private Function UrlEncode(sText As String) As String
    Dim sResult As String
    Dim sFinal  As String
    Dim sChar   As String
    Dim i       As Long
 
    sResult = ""
    sFinal = ""

    For i = 1 To Len(sText)
        sChar = mid$(sText, i, 1)

        If InStr(1, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", sChar) <> 0 Then
            sResult = sResult & sChar
        ElseIf sChar = " " Then
            sResult = sResult & "+"
        ElseIf sChar = vbCr Or sChar = vbLf Or sChar = vbCrLf Then
        Else
            sResult = sResult & "%" & Hex$(Asc(sChar))
        End If

        If Len(sResult) >= 1000 Then
            sFinal = sFinal & sResult
            sResult = ""
        End If
    Next
    UrlEncode = sFinal & sResult
End Function
 
Public Sub CloseConnection()
    On Error GoTo ERROR_HANDLER

    If m_hInetRequest > 0 Then Call InternetCloseHandle(m_hInetRequest)
    If m_hInetConnect > 0 Then Call InternetCloseHandle(m_hInetConnect)
    If m_hInetSession > 0 Then Call InternetCloseHandle(m_hInetSession)
    m_hInetRequest = 0
    m_hInetConnect = 0
    m_hInetSession = 0
    m_sPOSTData = ""
    m_sURLDomain = ""
    m_sURLPort = ""
    m_sURLPath = ""
    m_bURLSSL = False
ERROR_HANDLER:
End Sub
 
Private Sub Class_Terminate()
    Call CloseConnection
End Sub

