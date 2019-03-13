Attribute VB_Name = "modPing"
Option Explicit

Private Const IP_SUCCESS As Long = 0
Private Const IP_BUF_TOO_SMALL As Long = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE As Long = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE As Long = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE As Long = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE As Long = (11000 + 5)
Private Const IP_NO_RESOURCES As Long = (11000 + 6)
Private Const IP_BAD_OPTION As Long = (11000 + 7)
Private Const IP_HW_ERROR As Long = (11000 + 8)
Private Const IP_PACKET_TOO_BIG As Long = (11000 + 9)
Private Const IP_REQ_TIMED_OUT As Long = (11000 + 10)
Private Const IP_BAD_REQ As Long = (11000 + 11)
Private Const IP_BAD_ROUTE As Long = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT As Long = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM As Long = (11000 + 14)
Private Const IP_PARAM_PROBLEM As Long = (11000 + 15)
Private Const IP_SOURCE_QUENCH As Long = (11000 + 16)
Private Const IP_OPTION_TOO_BIG As Long = (11000 + 17)
Private Const IP_BAD_DESTINATION As Long = (11000 + 18)
Private Const IP_ADDR_DELETED As Long = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE As Long = (11000 + 20)
Private Const IP_MTU_CHANGE As Long = (11000 + 21)
Private Const IP_UNLOAD As Long = (11000 + 22)
Private Const IP_ADDR_ADDED As Long = (11000 + 23)
Private Const IP_GENERAL_FAILURE As Long = (11000 + 50)
Private Const IP_PENDING As Long = (11000 + 255)
Private Const PING_TIMEOUT As Long = 500
Private Const WS_VERSION_REQD As Long = &H101
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Private Type ICMP_OPTIONS
   Ttl             As Byte
   Tos             As Byte
   Flags           As Byte
   OptionsSize     As Byte
   OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
   Address         As Long
   status          As Long
   RoundTripTime   As Long
   DataSize        As Long
  'Reserved        As Integer
   DataPointer     As Long
   Options         As ICMP_OPTIONS
   data            As String * 250
End Type

Private Declare Function gethostbyname Lib "WSOCK32.DLL" _
  (ByVal hostname As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (XDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)

Private Declare Function WSAStartup Lib "WSOCK32.DLL" _
   (ByVal wVersionRequired As Long, _
    lpWSADATA As WSADATA) As Long

Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Private Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long

Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long

Private Declare Function inet_addr Lib "WSOCK32.DLL" _
   (ByVal s As String) As Long

'*****************************************
Public Function Ping(strAddress As String, strDataToSend As String, ECHO As ICMP_ECHO_REPLY) As Long


  'If Ping succeeds :
  '.RoundTripTime = time in ms for the ping to complete,
  '.Data is the data returned (NULL terminated)
  '.Address is the Ip address that actually replied
  '.DataSize is the size of the string in .Data
  '.Status will be 0
  '
  'If Ping fails .Status will be the error code

   Dim lngPortHandle As Long
   Dim lngAddress As Long

   'convert the address into a long representation
   lngAddress = inet_addr(strAddress)

   'if dwAddress is valid
   If lngAddress <> INADDR_NONE Then

     'open a port
      lngPortHandle = IcmpCreateFile()

     'and if successful,
      If lngPortHandle Then

        'ping it.
         Call IcmpSendEcho(lngPortHandle, _
                           lngAddress, _
                           strDataToSend, _
                           Len(strDataToSend), _
                           0, _
                           ECHO, _
                           Len(ECHO), _
                           PING_TIMEOUT)

        'return the status as ping success
         Ping = ECHO.status

        'close the port handle
         Call IcmpCloseHandle(lngPortHandle)

      End If  'If hPort

   Else:

         'the address format was probably invalid
         Ping = INADDR_NONE

   End If

End Function


'*****************************************
Public Function GetStatusCode(lngStatus As Long) As String

   Dim msg As String

   Select Case lngStatus
      Case IP_SUCCESS:               msg = "ip success"
      Case INADDR_NONE:              msg = "inet_addr: bad IP format"
      Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
      Case IP_NO_RESOURCES:          msg = "ip no resources"
      Case IP_BAD_OPTION:            msg = "ip bad option"
      Case IP_HW_ERROR:              msg = "ip hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
      Case IP_BAD_REQ:               msg = "ip bad req"
      Case IP_BAD_ROUTE:             msg = "ip bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
      Case IP_SOURCE_QUENCH:         msg = "ip source quench"
      Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
      Case IP_BAD_DESTINATION:       msg = "ip bad destination"
      Case IP_ADDR_DELETED:          msg = "ip addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
      Case IP_MTU_CHANGE:            msg = "ip mtu_change"
      Case IP_UNLOAD:                msg = "ip unload"
      Case IP_ADDR_ADDED:            msg = "ip addr added"
      Case IP_GENERAL_FAILURE:       msg = "ip general failure"
      Case IP_PENDING:               msg = "ip pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown  msg returned"
   End Select

   GetStatusCode = CStr(lngStatus) & "   [ " & msg & " ]"

End Function

'*****************************************
Public Function GetIPFromHostName(ByVal strHostName As String) As String

   'converts a host name to an IP address.

   Dim ptrHosent As Long  'address of hostent structure
   Dim ptrName As Long    'address of name pointer
   Dim ptrAddress As Long 'address of address pointer
   Dim ptrIPAddress As Long
   Dim strAddress As String

   strAddress = Space$(4)

   ptrHosent = gethostbyname(strHostName & vbNullChar)

   If ptrHosent <> 0 Then

      'assign pointer addresses and offset

      'ptrName is the official name of the host (PC).
      'If using the DNS or similar resolution system,
      'it is the Fully Qualified Domain Name (FQDN)
      'that caused the server to return a reply.
      'If using a local hosts file, it is the first
      'entry after the IP address.
      ptrName = ptrHosent

      'Null-terminated list of addresses for the host.
      'The Address is offset 12 bytes from the start of
      'the HOSENT structure. Note: Here we are retrieving
      'only the first address returned. To return more than
      'one, define sAddress as a string array and loop through
      'the 4-byte ptrIPAddress members returned. The last
      'item is a terminating null. All addresses are returned
      'in network byte order.
      ptrAddress = ptrHosent + 12

      'get the IP address
      CopyMemory ptrName, ByVal ptrName, 4
      CopyMemory ptrAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress, ByVal ptrAddress, 4
      CopyMemory ByVal strAddress, ByVal ptrIPAddress, 4

      GetIPFromHostName = IPToText(strAddress)

   End If

End Function


'*****************************************
Private Function IPToText(ByVal IPAddress As String) As String

   IPToText = CStr(Asc(IPAddress)) & "." & _
              CStr(Asc(mid$(IPAddress, 2, 1))) & "." & _
              CStr(Asc(mid$(IPAddress, 3, 1))) & "." & _
              CStr(Asc(mid$(IPAddress, 4, 1)))

End Function

'*****************************************
Public Sub SocketsCleanup()

   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If

End Sub


'*****************************************
Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA

   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS

End Function
'*****************************************

'*****************************************
Public Function PingAddress(strAddressToResolve As String, strDataToSend As String) As String

   Dim ECHO As ICMP_ECHO_REPLY
   Dim logPos As Long
   Dim lngSuccess As Long
   Dim strIPAddress As String

   If SocketsInitialize() Then

     'convert the host name into an IP address
      strIPAddress = GetIPFromHostName(strAddressToResolve)
      Debug.Print "Resolved IP Address : " & strIPAddress

     'ping the ip passing the address, text
     'to use, and the ECHO structure
      lngSuccess = Ping(strIPAddress, strDataToSend, ECHO)

      'display the results
      Debug.Print "Ping Status Code : " & GetStatusCode(lngSuccess)
      Debug.Print "Echo Addess : " & ECHO.Address
      Debug.Print "Round Trip Time : " & ECHO.RoundTripTime & " ms"
      Debug.Print "Data Size : " & ECHO.DataSize & " bytes"

      If Left$(ECHO.data, 1) <> Chr$(0) Then
         logPos = InStr(ECHO.data, Chr$(0))
         Debug.Print "Echo Data : " & Left$(ECHO.data, logPos - 1)
      End If

      Debug.Print "Data Pointer : " & ECHO.DataPointer

      Call SocketsCleanup
      PingAddress = ECHO.RoundTripTime & " ms"

   Else

        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."

   End If

End Function

