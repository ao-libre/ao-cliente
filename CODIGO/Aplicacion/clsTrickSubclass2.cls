VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrickSubclass2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsTrickSubclass2.cls - class for window subclassing
' © Krivous Anatolii Anatolevich (The trick), 2015-2016
' Version 2.2

Private Type PROCESS_HEAP_ENTRY
    lpData              As Long
    cbData              As Long
    cbOverhead          As Byte
    iRegionIndex        As Byte
    wFlags              As Integer
    dwCommittedSize     As Long
    dwUnCommittedSize   As Long
    lpFirstBlock        As Long
    lpLastBlock         As Long
End Type

Private Declare Function SetWindowSubclass Lib "Comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, dwRefData As Any) As Long
Private Declare Function RemoveWindowSubclass Lib "Comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "Comctl32" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Private Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function HeapWalk Lib "kernel32" (ByVal hHeap As Long, ByRef lpEntry As PROCESS_HEAP_ENTRY) As Long
Private Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpValue As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (pSrc As Any, pDst As Any) As Long

Private Const WM_CREATE                     As Long = &H1
Private Const WM_DESTROY                    As Long = &H2
Private Const GCL_WNDPROC                   As Long = (-24)
Private Const GWL_WNDPROC                   As Long = (-4)
Private Const HEAP_CREATE_ENABLE_EXECUTE    As Long = &H40000
Private Const HEAP_NO_SERIALIZE             As Long = &H1
Private Const HEAP_ZERO_MEMORY              As Long = &H8
Private Const PROCESS_HEAP_ENTRY_BUSY       As Long = &H4
Private Const WNDPROCINDEX                  As Long = 8
Private Const EnvName                       As String = "TrickSubclass"

Public Event WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, Ret As Long, DefCall As Boolean)

Private mIsSubclassed   As Boolean
Private mhWnd           As Long
Private mIsPaused       As Boolean
Private mTerminateFlag  As Boolean
Private mDepth          As Long
Private mSelf           As clsTrickSubclass2

Dim hHeap   As Long
Dim lpAsm   As Long

' Return a window handle
Public Property Get hWnd() As Long
    hWnd = mhWnd
End Property
' Subclassing state (True - subclassing on)
Public Property Get IsSubclassed() As Boolean
    IsSubclassed = mIsSubclassed
End Property
' Pause subclassing
Public Function PauseSubclass() As Boolean
    If mIsSubclassed And Not mIsPaused Then mIsPaused = True: PauseSubclass = True
End Function
' Resume
Public Function ResumeSubclass() As Boolean
    If mIsSubclassed And mIsPaused Then mIsPaused = False: ResumeSubclass = True
End Function
' If pause then return True
Public Property Get IsPaused() As Boolean
    IsPaused = mIsPaused
End Property
' Set subclassing to window (if subclassing already enabled then remove it)
Public Function Hook(ByVal hWnd As Long) As Boolean

    If mIsSubclassed Then
        If Not UnHook Then Exit Function
    End If
    
    If CreateAsm Then
        
        mIsSubclassed = SetWindowSubclass(hWnd, lpAsm, ObjPtr(Me), 0)
        
        If mIsSubclassed Then
            Hook = True
            mhWnd = hWnd
        End If
        
    End If
    
End Function
' Remove subclassing
Public Function UnHook() As Boolean
    If Not mIsSubclassed Then Exit Function
    UnHook = RemoveWindowSubclass(mhWnd, lpAsm, ObjPtr(Me))
    If UnHook Then mhWnd = 0: mIsSubclassed = False
End Function
' Call default procedure
Public Function CallDef(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef status As Boolean) As Long
    If Not mIsSubclassed Then Exit Function
    CallDef = DefSubclassProc(hWnd, msg, wParam, lParam)
    status = True
End Function

' --------------------------------------------------------------------------------------------------------------------------------------
Private Function SUBCLASSPROC(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    Dim inIDE   As Boolean
    Dim retAddr As Long
    Dim addr    As Long
    
    mDepth = mDepth + 1
    
    If mIsPaused Then
        SUBCLASSPROC = DefSubclassProc(hWnd, msg, wParam, lParam)
    Else
        Dim DefCall As Boolean
        DefCall = True
        RaiseEvent WndProc(hWnd, msg, wParam, lParam, SUBCLASSPROC, DefCall)
        If DefCall Then SUBCLASSPROC = DefSubclassProc(hWnd, msg, wParam, lParam)
    End If
     
    mDepth = mDepth - 1
    
    Debug.Assert MakeTrue(inIDE)
    
    If inIDE Then
        Dim refDat  As Long
        GetMem4 ByVal ObjPtr(Me) + 8, refDat
        GetMem4 ByVal refDat + 4, refDat
        If refDat = 1 Then
            addr = VarPtr(hWnd) + &H20
            GetMem4 ByVal addr, ByVal addr - &H28
        End If
    Else
        If mTerminateFlag And mDepth = 0 Then
            addr = VarPtr(hWnd) + &H20
            GetMem4 ByVal addr, ByVal addr - &H28
            ' // Clean
            Call Class_Terminate
        End If
    End If
    
End Function

Private Sub Class_Terminate()

    If hHeap = 0 Then Exit Sub
    
    UnHook
    
    If mDepth Then
        Set mSelf = Me
        mTerminateFlag = True
    Else
        If CountHooks = 1 Then
            HeapDestroy hHeap
            hHeap = 0
            SaveCurHeap
        Else
            HeapFree hHeap, HEAP_NO_SERIALIZE, ByVal lpAsm
        End If
        Set mSelf = Nothing
    End If
    
End Sub
Private Function CreateAsm() As Boolean
    Dim inIDE   As Boolean
    Dim AsmSize As Long
    Dim ptr     As Long
    Dim isFirst As Boolean

    Debug.Assert MakeTrue(inIDE)
    
    If lpAsm = 0 Then
        If inIDE Then AsmSize = &H5E Else AsmSize = &H1D
        hHeap = GetPrevHeap()
        
        If hHeap Then
            If inIDE Then
                Dim flag    As Long
                ptr = GetFlagPointer()
                GetMem4 ByVal ptr, flag
                If flag Then
                    HeapDestroy hHeap
                    isFirst = True
                End If
            End If
        Else: isFirst = True
        End If
        
        If isFirst Then
            hHeap = HeapCreate(HEAP_CREATE_ENABLE_EXECUTE Or HEAP_NO_SERIALIZE, 0, 0)
            If hHeap = 0 Then Err.Raise 7: Exit Function
            If Not SaveCurHeap() Then HeapDestroy hHeap: hHeap = 0: Err.Raise 7: Exit Function
            AsmSize = AsmSize + &H4
        End If
        
        lpAsm = HeapAlloc(hHeap, HEAP_NO_SERIALIZE Or HEAP_ZERO_MEMORY, AsmSize)
        
        If lpAsm = 0 Then
            If isFirst Then HeapDestroy hHeap
            hHeap = 0
            Err.Raise 7
            Exit Function
        End If
        
        Dim prv As Long
        Dim i   As Long
        
        If inIDE Then
            If isFirst Then
                GetMem4 0&, ByVal lpAsm
                lpAsm = lpAsm + 4
            End If
        End If
        
    End If
    
    ptr = lpAsm
    
    If inIDE Then
        CreateIDEStub (ptr): ptr = ptr + &H40
    End If
    
    CreateStackConv ptr
    CreateAsm = True
    
End Function
Private Function GetFlagPointer() As Long
    Dim he  As PROCESS_HEAP_ENTRY
    HeapLock hHeap
    Do While HeapWalk(hHeap, he)
        If he.wFlags And PROCESS_HEAP_ENTRY_BUSY Then GetFlagPointer = he.lpData: Exit Function
    Loop
    HeapUnlock hHeap
End Function
Private Function CountHooks() As Long
    Dim he  As PROCESS_HEAP_ENTRY
    HeapLock hHeap
    Do While HeapWalk(hHeap, he)
        If he.wFlags And PROCESS_HEAP_ENTRY_BUSY Then CountHooks = CountHooks + 1
    Loop
    HeapUnlock hHeap
End Function
Private Function SaveCurHeap() As Boolean
    Dim i   As Long
    Dim out As String
    out = Hex(hHeap)
    For i = Len(out) + 1 To 8: out = "0" & out: Next
    SaveCurHeap = SetEnvironmentVariable(StrPtr(EnvName), StrPtr(out))
End Function
Private Function GetPrevHeap() As Long
    Dim out         As String
    out = Space(&H8)
    If GetEnvironmentVariable(StrPtr(EnvName), StrPtr(out), LenB(out)) Then GetPrevHeap = Val("&H" & out)
End Function
Private Function CreateStackConv(ByVal ptr As Long) As Boolean
    Dim lpMeth      As Long
    Dim vTable      As Long
    
    GetMem4 ByVal ObjPtr(Me), vTable
    GetMem4 ByVal vTable + WNDPROCINDEX * 4 + &H1C, lpMeth
    
    GetMem4 &H5450C031, ByVal ptr + &H0:    GetMem4 &H488DE409, ByVal ptr + &H4:    GetMem4 &H2474FF06, ByVal ptr + &H8
    GetMem4 &H68FAE020, ByVal ptr + &HC:    GetMem4 &H12345678, ByVal ptr + &H10:   GetMem4 &HFFFFE7E8, ByVal ptr + &H14
    GetMem4 &H18C258FF, ByVal ptr + &H18:   GetMem4 &H0, ByVal ptr + &H1C

    GetMem4 ObjPtr(Me), ByVal ptr + &H10                    ' Push Me
    GetMem4 lpMeth - (ptr + &H14) - 5, ByVal ptr + &H14 + 1 ' Call SUBCLASSPROC
    
End Function
 
Private Function CreateIDEStub(ByVal ptr As Long) As Boolean
    Dim hInstVB6    As Long
    Dim lpEbMode    As Long
    Dim hComctl32   As Long
    Dim lpDefProc   As Long
    Dim lpRemove    As Long
    
    hInstVB6 = GetModuleHandle(StrPtr("vba6"))
    If hInstVB6 = 0 Then Exit Function
    hComctl32 = GetModuleHandle(StrPtr("Comctl32"))
    If hComctl32 = 0 Then
        hComctl32 = LoadLibrary(StrPtr("Comctl32"))
        If hComctl32 = 0 Then Exit Function
    End If
    
    lpEbMode = GetProcAddress(hInstVB6, "EbMode")
    If lpEbMode = 0 Then Exit Function
    lpDefProc = GetProcAddress(hComctl32, "DefSubclassProc")
    If lpDefProc = 0 Then Exit Function
    lpRemove = GetProcAddress(hComctl32, "RemoveWindowSubclass")
    If lpRemove = 0 Then Exit Function
    
    GetMem4 &HFFFFFBE8, ByVal ptr + &H0:    GetMem4 &H74C084FF, ByVal ptr + &H4:    GetMem4 &H74013C1C, ByVal ptr + &H8
    GetMem4 &H2474FF33, ByVal ptr + &HC:    GetMem4 &H2474FF10, ByVal ptr + &H10:   GetMem4 &H2474FF10, ByVal ptr + &H14
    GetMem4 &H2474FF10, ByVal ptr + &H18:   GetMem4 &HFFDEE810, ByVal ptr + &H1C:   GetMem4 &H18C2FFFF, ByVal ptr + &H20
    GetMem4 &HDFF00, ByVal ptr + &H24:      GetMem4 &H68000000, ByVal ptr + &H28:   GetMem4 &H12345678, ByVal ptr + &H2C
    GetMem4 &H34567868, ByVal ptr + &H30:   GetMem4 &H2474FF12, ByVal ptr + &H34:   GetMem4 &HFFC2E80C, ByVal ptr + &H38
    GetMem4 &HCDEBFFFF, ByVal ptr + &H3C:

    GetMem4 lpEbMode - ptr - 5, ByVal ptr + 1 + 0                   ' Call EbMode
    GetMem4 lpDefProc - (ptr + &H1D) - 5, ByVal ptr + &H1D + 1      ' Call DefSubclassProc
    GetMem4 lpRemove - (ptr + &H39) - 5, ByVal ptr + &H39 + 1       ' Call RemoveWindowSubclass
    GetMem4 ObjPtr(Me), ByVal ptr + &H2C                            ' Push uIdSubclass
    GetMem4 ptr, ByVal ptr + &H31                                   ' Push pfnSubclass
    GetMem4 GetFlagPointer(), ByVal ptr + &H27                      ' dec dword [flag]
    
    CreateIDEStub = True
End Function
Private Function MakeTrue(Value As Boolean) As Boolean: Value = True: MakeTrue = True: End Function
