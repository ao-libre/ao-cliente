Attribute VB_Name = "Cripto"
'--- mdAesEcb.bas
Option Explicit
DefObj A-Z
 
#Const ImplUseShared = False
 
'=========================================================================
' API
'=========================================================================
 
'--- for CNG
Private Const MS_PRIMITIVE_PROVIDER         As String = "Microsoft Primitive Provider"
Private Const BCRYPT_CHAIN_MODE_ECB         As String = "ChainingModeECB"
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG   As Long = 8
 
'--- for CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1
'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
 
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbOutput As Any, ByVal cbOutput As Long, cbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptSetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGenerateSymmetricKey Lib "bcrypt" (ByVal hAlgorithm As Long, phKey As Long, pbKeyObject As Any, ByVal cbKeyObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
Private Declare Function BCryptEncrypt Lib "bcrypt" (ByVal hKey As Long, pbInput As Any, ByVal cbInput As Long, ByVal pPaddingInfo As Long, ByVal pbIV As Long, ByVal cbIV As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDeriveKeyPBKDF2 Lib "bcrypt" (ByVal pPrf As Long, pbPassword As Any, ByVal cbPassword As Long, pbSalt As Any, ByVal cbSalt As Long, ByVal cIterations As Long, ByVal dwDummy As Long, pbDerivedKey As Any, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCreateHash Lib "bcrypt" (ByVal hAlgorithm As Long, phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyHash Lib "bcrypt" (ByVal hHash As Long) As Long
Private Declare Function BCryptHashData Lib "bcrypt" (ByVal hHash As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptFinishHash Lib "bcrypt" (ByVal hHash As Long, pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
#If Not ImplUseShared Then
    Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
    Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, pcchString As Long) As Long
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
#End If
 
'=========================================================================
' Constants and member variables
'=========================================================================
 
Private Const ERR_UNSUPPORTED_ENCR  As String = "Unsupported encryption"
Private Const AES_BLOCK_SIZE        As Long = 16
Private Const AES_KEYLEN            As Long = 32                    '-- 32 -> AES-256, 24 -> AES-196, 16 -> AES-128
Private Const AES_SALT              As String = "SaltVb6CryptoAes"  '-- at least 16 chars
 
Private Type UcsZipCryptoType
    hPbkdf2Alg          As Long
    hHmacAlg            As Long
    hHmacHash           As Long
    HmacHashLen         As Long
    hAesAlg             As Long
    hAesKey             As Long
    AesKeyObjData()     As Byte
    AesKeyObjLen        As Long
    Nonce(0 To 1)       As Long
    EncrData()          As Byte
    EncrPos             As Long
    LastError           As String
End Type
 
'=========================================================================
' Functions
'=========================================================================
 
Public Function AesEncryptString(sText As String, sPassword As String) As String
    Dim baData()        As Byte
    Dim sError          As String
    
    baData = ToUtf8Array(sText)
    If Not AesCryptArray(baData, ToUtf8Array(sPassword), Error:=sError) Then
        Err.Raise vbObjectError, , sError
    End If
    AesEncryptString = ToBase64Array(baData)
End Function
 
Public Function AesDecryptString(sEncr As String, sPassword As String) As String
    Dim baData()        As Byte
    Dim sError          As String
    
    baData = FromBase64Array(sEncr)
    If Not AesCryptArray(baData, ToUtf8Array(sPassword), Error:=sError) Then
        Err.Raise vbObjectError, , sError
    End If
    AesDecryptString = FromUtf8Array(baData)
End Function
 
Public Function AesCryptArray( _
            baData() As Byte, _
            baPass() As Byte, _
            Optional Salt As String, _
            Optional ByVal KeyLen As Long, _
            Optional Error As String, _
            Optional HmacSha1 As Variant) As Boolean
    Const VT_BYREF      As Long = &H4000
    Dim uCtx            As UcsZipCryptoType
    Dim vErr            As Variant
    Dim bHashBefore     As Boolean
    Dim bHashAfter      As Boolean
    Dim baTemp()        As Byte
    Dim lPtr            As Long
    
    On Error GoTo EH
    If Not IsMissing(HmacSha1) Then
        bHashBefore = (HmacSha1(0) <= 0)
        bHashAfter = (HmacSha1(0) > 0)
    End If
    If LenB(Salt) > 0 Then
        baTemp = ToUtf8Array(Salt)
    Else
        baTemp = ToUtf8Array(AES_SALT)
    End If
    If KeyLen <= 0 Then
        KeyLen = AES_KEYLEN
    End If
    If Not pvCryptoAesInit(uCtx, baPass, baTemp, KeyLen, 0) Then
        Error = uCtx.LastError
        GoTo QH
    End If
    If Not pvCryptoAesCrypt(uCtx, baData, Size:=UBound(baData) + 1, HashBefore:=bHashBefore, HashAfter:=bHashAfter) Then
        Error = uCtx.LastError
        GoTo QH
    End If
    If Not IsMissing(HmacSha1) Then
        baTemp = pvCryptoAesGetFinalHash(uCtx, UBound(HmacSha1) + 1)
        lPtr = Peek((VarPtr(HmacSha1) Xor &H80000000) + 8 Xor &H80000000)
        If (Peek(VarPtr(HmacSha1)) And VT_BYREF) <> 0 Then
            lPtr = Peek(lPtr)
        End If
        lPtr = Peek((lPtr Xor &H80000000) + 12 Xor &H80000000)
        Call CopyMemory(ByVal lPtr, baTemp(0), UBound(baTemp) + 1)
    End If
    '--- success
    AesCryptArray = True
QH:
    pvCryptoAesTerminate uCtx
    Exit Function
EH:
    vErr = Array(Err.number, Err.source, Err.Description)
    pvCryptoAesTerminate uCtx
    Err.Raise vErr(0), vErr(1), vErr(2)
End Function
 
'= private ===============================================================
 
Private Function pvCryptoAesInit(uCrypto As UcsZipCryptoType, baPass() As Byte, baSalt() As Byte, ByVal lKeyLen As Long, nPassVer As Integer) As Boolean
    Dim baDerivedKey()  As Byte
    Dim lDummy          As Long '--- discarded
    Dim hResult         As Long
    Dim sApiSource      As String
    
    '--- init member vars
    uCrypto.Nonce(0) = 0
    uCrypto.Nonce(1) = 0
    uCrypto.EncrData = vbNullString
    uCrypto.EncrPos = 0
    '--- generate RFC 2898 based derived key
    On Error GoTo EH_Unsupported '--- CNG API missing on XP
    hResult = BCryptOpenAlgorithmProvider(uCrypto.hPbkdf2Alg, StrPtr("SHA1"), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG)
    If hResult <> 0 Then
        sApiSource = "BCryptOpenAlgorithmProvider(SHA1)"
        GoTo QH
    End If
    On Error GoTo 0
    ReDim baDerivedKey(0 To 2 * lKeyLen + 1) As Byte
    On Error GoTo EH_Unsupported '--- PBKDF2 API missing on Vista
    hResult = BCryptDeriveKeyPBKDF2(uCrypto.hPbkdf2Alg, baPass(0), UBound(baPass) + 1, baSalt(0), UBound(baSalt) + 1, 1000, 0, baDerivedKey(0), UBound(baDerivedKey) + 1, 0)
    If hResult <> 0 Then
        sApiSource = "BCryptDeriveKeyPBKDF2"
        GoTo QH
    End If
    On Error GoTo 0
    '--- extract Password Verification Value from last 2 bytes of derived key
    Call CopyMemory(nPassVer, baDerivedKey(2 * lKeyLen), 2)
    '--- init AES w/ ECB from first half of derived key
    hResult = BCryptOpenAlgorithmProvider(uCrypto.hAesAlg, StrPtr("AES"), StrPtr(MS_PRIMITIVE_PROVIDER), 0)
    If hResult <> 0 Then
        sApiSource = "BCryptOpenAlgorithmProvider(AES)"
        GoTo QH
    End If
    hResult = BCryptGetProperty(uCrypto.hAesAlg, StrPtr("ObjectLength"), uCrypto.AesKeyObjLen, 4, lDummy, 0)
    If hResult <> 0 Then
        sApiSource = "BCryptGetProperty(ObjectLength)"
        GoTo QH
    End If
    hResult = BCryptSetProperty(uCrypto.hAesAlg, StrPtr("ChainingMode"), StrPtr(BCRYPT_CHAIN_MODE_ECB), LenB(BCRYPT_CHAIN_MODE_ECB), 0)
    If hResult <> 0 Then
        sApiSource = "BCryptSetProperty(ChainingMode)"
        GoTo QH
    End If
    ReDim uCrypto.AesKeyObjData(0 To uCrypto.AesKeyObjLen - 1) As Byte
    hResult = BCryptGenerateSymmetricKey(uCrypto.hAesAlg, uCrypto.hAesKey, uCrypto.AesKeyObjData(0), uCrypto.AesKeyObjLen, baDerivedKey(0), lKeyLen, 0)
    If hResult <> 0 Then
        sApiSource = "BCryptGenerateSymmetricKey"
        GoTo QH
    End If
    '-- init HMAC from second half of derived key
    hResult = BCryptOpenAlgorithmProvider(uCrypto.hHmacAlg, StrPtr("SHA1"), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG)
    If hResult <> 0 Then
        sApiSource = "BCryptOpenAlgorithmProvider(SHA1)"
        GoTo QH
    End If
    hResult = BCryptGetProperty(uCrypto.hHmacAlg, StrPtr("HashDigestLength"), uCrypto.HmacHashLen, 4, lDummy, 0)
    If hResult <> 0 Then
        sApiSource = "BCryptGetProperty(HashDigestLength)"
        GoTo QH
    End If
    hResult = BCryptCreateHash(uCrypto.hHmacAlg, uCrypto.hHmacHash, 0, 0, baDerivedKey(lKeyLen), lKeyLen, 0)
    If hResult <> 0 Then
        sApiSource = "BCryptCreateHash"
        GoTo QH
    End If
    '--- success
    pvCryptoAesInit = True
    Exit Function
QH:
    If Err.LastDllError <> 0 Then
        uCrypto.LastError = GetSystemMessage(Err.LastDllError)
    Else
        uCrypto.LastError = "[" & Hex$(hResult) & "] Error in " & sApiSource
    End If
    Exit Function
EH_Unsupported:
    uCrypto.LastError = ERR_UNSUPPORTED_ENCR
End Function
 
Private Sub pvCryptoAesTerminate(uCrypto As UcsZipCryptoType)
    If uCrypto.hPbkdf2Alg <> 0 Then
        Call BCryptCloseAlgorithmProvider(uCrypto.hPbkdf2Alg, 0)
        uCrypto.hPbkdf2Alg = 0
    End If
    If uCrypto.hHmacHash <> 0 Then
        Call BCryptDestroyHash(uCrypto.hHmacHash)
        uCrypto.hHmacHash = 0
    End If
    If uCrypto.hHmacAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(uCrypto.hHmacAlg, 0)
        uCrypto.hHmacAlg = 0
    End If
    If uCrypto.hAesKey <> 0 Then
        Call BCryptDestroyKey(uCrypto.hAesKey)
        uCrypto.hAesKey = 0
    End If
    If uCrypto.hAesAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(uCrypto.hAesAlg, 0)
        uCrypto.hAesAlg = 0
    End If
End Sub
 
Private Function pvCryptoAesCrypt( _
            uCrypto As UcsZipCryptoType, _
            baData() As Byte, _
            Optional ByVal Offset As Long, _
            Optional ByVal Size As Long, _
            Optional ByVal HashBefore As Boolean, _
            Optional ByVal HashAfter As Boolean) As Boolean
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lPadSize        As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If Size < 0 Then
        Size = UBound(baData) + 1 - Offset
    End If
    If HashBefore Then
        hResult = BCryptHashData(uCrypto.hHmacHash, baData(Offset), Size, 0)
        If hResult <> 0 Then
            sApiSource = "BCryptHashData"
            GoTo QH
        End If
    End If
    With uCrypto
        '--- reuse EncrData from prev call until next AES_BLOCK_SIZE boundary
        For lIdx = Offset To Offset + Size - 1
            If (.EncrPos And (AES_BLOCK_SIZE - 1)) = 0 Then
                Exit For
            End If
            baData(lIdx) = baData(lIdx) Xor .EncrData(.EncrPos)
            .EncrPos = .EncrPos + 1
        Next
        If lIdx < Offset + Size Then
            '--- pad remaining input size to AES_BLOCK_SIZE
            lPadSize = (Offset + Size - lIdx + AES_BLOCK_SIZE - 1) And -AES_BLOCK_SIZE
            If UBound(.EncrData) + 1 < lPadSize Then
                ReDim .EncrData(0 To lPadSize - 1) As Byte
            End If
            '--- encrypt incremental nonces in EncrData
            For lJdx = 0 To lPadSize - 1 Step 16
                If .Nonce(0) <> -1 Then
                    .Nonce(0) = (.Nonce(0) Xor &H80000000) + 1 Xor &H80000000
                Else
                    .Nonce(0) = 0
                    .Nonce(1) = (.Nonce(1) Xor &H80000000) + 1 Xor &H80000000
                End If
                Call CopyMemory(.EncrData(lJdx), .Nonce(0), 8)
            Next
            hResult = BCryptEncrypt(.hAesKey, .EncrData(0), lPadSize, 0, 0, 0, .EncrData(0), lPadSize, lJdx, 0)
            If hResult <> 0 Then
                sApiSource = "BCryptEncrypt"
                GoTo QH
            End If
            '--- xor remaining input and leave anything extra of EncrData for reuse
            For .EncrPos = 0 To Offset + Size - lIdx - 1
                baData(lIdx) = baData(lIdx) Xor .EncrData(.EncrPos)
                lIdx = lIdx + 1
            Next
        End If
    End With
    If HashAfter Then
        hResult = BCryptHashData(uCrypto.hHmacHash, baData(Offset), Size, 0)
        If hResult <> 0 Then
            sApiSource = "BCryptHashData"
            GoTo QH
        End If
    End If
    '--- success
    pvCryptoAesCrypt = True
    Exit Function
QH:
    If Err.LastDllError <> 0 Then
        uCrypto.LastError = GetSystemMessage(Err.LastDllError)
    Else
        uCrypto.LastError = "[" & Hex$(hResult) & "] Error in " & sApiSource
    End If
End Function
 
Private Function pvCryptoAesGetFinalHash(uCrypto As UcsZipCryptoType, ByVal lSize As Long) As Byte()
    Dim baResult()      As Byte
    
    ReDim baResult(0 To uCrypto.HmacHashLen - 1) As Byte
    Call BCryptFinishHash(uCrypto.hHmacHash, baResult(0), uCrypto.HmacHashLen, 0)
    ReDim Preserve baResult(0 To lSize - 1) As Byte
    pvCryptoAesGetFinalHash = baResult
End Function
 
'= shared ================================================================
 
#If Not ImplUseShared Then
Public Function ToBase64Array(baData() As Byte) As String
    Dim lSize           As Long
    
    If UBound(baData) >= 0 Then
        ToBase64Array = String$(2 * UBound(baData) + 6, 0)
        lSize = Len(ToBase64Array) + 1
        Call CryptBinaryToString(VarPtr(baData(0)), UBound(baData) + 1, CRYPT_STRING_BASE64, StrPtr(ToBase64Array), lSize)
        ToBase64Array = Left$(ToBase64Array, lSize)
    End If
End Function
 
Public Function FromBase64Array(sText As String) As Byte()
    Dim lSize           As Long
    Dim baOutput()      As Byte
    
    lSize = Len(sText) + 1
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, 0)
    If lSize > 0 Then
        ReDim Preserve baOutput(0 To lSize - 1) As Byte
        FromBase64Array = baOutput
    Else
        FromBase64Array = vbNullString
    End If
End Function
 
Public Function ToUtf8Array(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
    Else
        baRetVal = vbNullString
    End If
    ToUtf8Array = baRetVal
End Function
 
Public Function FromUtf8Array(baText() As Byte) As String
    Dim lSize           As Long
    
    If UBound(baText) >= 0 Then
        FromUtf8Array = String$(2 * UBound(baText), 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, StrPtr(FromUtf8Array), Len(FromUtf8Array))
        FromUtf8Array = Left$(FromUtf8Array, lSize)
    End If
End Function
 
Public Function GetSystemMessage(ByVal lLastDllError As Long) As String
    Dim lSize            As Long
   
    GetSystemMessage = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lLastDllError, 0&, GetSystemMessage, Len(GetSystemMessage), 0&)
    If lSize > 2 Then
        If mid$(GetSystemMessage, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetSystemMessage = "[" & lLastDllError & "] " & Left$(GetSystemMessage, lSize)
End Function
 
Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function
#End If

