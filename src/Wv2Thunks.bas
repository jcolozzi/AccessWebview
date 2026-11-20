Attribute VB_Name = "Wv2Thunks"
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas N-4 __ (_______) ---
'
'   N-4 ____:
'     - HandlerKind _ HK_ResponseGetContent = 13 .
'       *1 ___________* = HK_ExecuteScriptCompleted .
'       (__________. CallbackId __________)
'     - m_iidTable ______ HK_ResponseGetContent .
'     - IID {875738e1-9fa2-40e3-8b74-2e8972dd6fe7}
'
'   *Invoke ___ ExecuteScriptCompleted ___*
'     Invoke(HRESULT errorCode, IStream* result)
'     -> arg1 = errorCode (__ 32bit), arg2 = IStream* (0 ______)
'     _9.8c ____ CallbackId ____________.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas N-3 __ (_______________) ---
'
'   N-3 ____:
'     - HandlerKind _ HK_WebResourceResponseReceived = 12 .
'       *__, ICoreWebView2_2 (vtable 61/62)* = __ IF .
'     - m_iidTable ______ HK_WebResourceRequested <- HK_WebResourceResponseReceived
'     - IID {7de9898a-24f5-40c3-a2de-d4f458e69828}
'       *___ SDK ____ MIDL_INTERFACE _ (____115)*
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas N-2 __ (_____________) ---
'
'   N-2 ____:
'     - `MultiByteToWideChar` ______ (kernel32)
'     - `Utf8BytesToString` ___. *UTF-8 ______ VBA String ___*
'
'   *__ ADODB.Stream ______*
'     ________ COM ___________. ADODB .
'     (Wv2Tests _ WriteUtf8NoBom) _________. `MultiByteToWideChar` .
'     Win32 API ___, _______ `lstrlenW` / `RtlMoveMemory` ____.
'
'   *StrConv(b, vbUnicode) ____*
'     ________________ (_______ CP932) _____.
'     UTF-8 ____________. ____________ API ___.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas N-1 __ (WebResourceRequested ___) ---
'
'   N-1 ____:
'     - HandlerKind _ HK_WebResourceRequested = 11 ___ (__, __ ICoreWebView2)
'     - m_iidTable ______ HK_DOMContentLoaded <- HK_WebResourceRequested .
'     - InitIIDTable _ ICoreWebView2WebResourceRequestedEventHandler _ IID .
'       {ab00b74c-15f1-4646-80e8-e76341d25d71}
'       *___ SDK ___ WebView2.h _ MIDL_INTERFACE _ (____115)*
'     - ______ (Handler_QueryInterface / AddRef / Release / dcf) .
'       (kind ___________ kind=11 _______)
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas _9.16__ (__: Module9_9_9.bas) ---
'
'   _9.16 ____ (__: ________ Sub _ Wv2Tests.bas ___):
'     *__________________ (_1-A)*
'       Wv2Thunks ___________vtable_SAFEARRAY .
'       ___________. .
'       (Test_9_14_SearchFallback / Test_9_15_Reorder / Test_9_16_SearchEngine .
'        ____ Check / CheckOrder) _, ____ Wv2Tests.bas ____.
'       ____ Wv2Browser _ Debug_* ___________________,
'       ____________ (Wv2Json _____________________).
'
'     *_____________*
'       Test_TwoHandlers_Mock / Test_Environment_Real / Test_ResetSurvival_Manual /
'       Test_VirtualAllocFree_Roundtrip / Test_SentinelStatus _, .
'       Private __ (Thunks_* / Sentinel_* / VirtualAlloc _) .
'       _____. ___________ Private _ Public ________,
'       _____________ (________________________).
'
'     - ______ (______ Public/Private API) ______.
'
' --- Wv2Thunks.bas _9.15__ (__: Module9_9_9.bas) ---
'
'   _9.15 ____ (_____________ Sub __):
'     - Test_9_15_Reorder ___. Wv2Browser ____ (Debug_SeedDummyTabs /
'       MoveTab / Debug_TabOrderTitles / Debug_ActiveIndex) ___, .
'       ___________ WebView2 ___________ (________).
'       ___ OnPaneWebMessage(reorder) ________ MoveTab .
'       _________ (Test_9_14 _ Debug_NormalizeUrl _______).
'     - CheckOrder _______ (______ activeIndex ______).
'     - ________________. _____ Sub _____.
'
' --- Wv2Thunks.bas _9.14__ (__: Module9_9_9.bas) ---
'
'   _9.14 ____ (_________________ Sub __):
'     - Test_9_14_SearchFallback ___. Wv2Browser.Debug_NormalizeUrl ___,
'       __________URL___Google .
'       WebView2 ___________ (________).
'       ___ Wv2Browser _ Debug_NormalizeUrl _____, ___ NavigateActive .
'       __ NormalizeUrl __________ (_________).
'     - _______: "localhost:8080" __________ (IsAlphaOnly) .
'       ____________, Edge/WebView2 _ http:// .
'       (=____________). __ Sub _______________.
'     - ________________. _____ Sub _____.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas _9.13__ (__: Module9_9_9.bas) ---
'
'   _9.13 ____ (_________, ______):
'     - UserForm ________ Sub 4 ______. _____9.13 .
'       UserForm1 _______ StartWebView2 (___ IsPageLoaded /
'       GetActivePane / OnExecuteScriptResult) ______, .
'       ________________:
'         * Test_Controller_Real   (_9.3b)  <- StartWebView2 .
'         * Test_Navigate_Real     (_9.3c)  <- StartWebView2 .
'         * Test_AllEvents_Real    (_9.4b)  <- StartWebView2 .
'         * Test_ExecuteScript_Parallel (_9.9c) <- StartWebView2 .
'       ExecuteScript _________9.9c _____. __ SPA .
'       _______ StartWebView2_Spa .
'       __, _________ (____).
'     - _____________ 5 _ (____vtable__________.
'       GUI _____________):
'         * Test_TwoHandlers_Mock           (vtable/_________)
'         * Test_Environment_Real           (Environment ____)
'         * Test_VirtualAllocFree_Roundtrip (_________)
'         * Test_SentinelStatus             (_____)
'         * Test_ResetSurvival_Manual       (VBA Reset _____)
'     - Test_ResetSurvival_Manual ____________ Test_Navigate_Real .
'       ________, UserForm1.StartWebView2_Full .
'       (_______________).
'     - ______ (Handler_xxx / dcf / Thunks_Init / __________) .
'       ______. ___________ Sub ______.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas _9.10a__ (__: Module9_9_9.bas) ---
'
'   _9.10a ____ (______, Sentinel MEM_FREE ____ = ____ 10):
'     - ____: MEM_FREE (&H10000&) / PAGE_READWRITE (&H4&)
'     - Sentinel_RecoverIfNeeded __State ___________:
'         VirtualQuery ___ (State=MEM_COMMIT __ AllocationBase=prevBase .
'         Protect ______) ___________________,
'         ReadLongPtr / MemLongPtr / VirtualFree _____.
'         _______ Sentinel_ClearPrevRegion ________ Exit Sub.
'       -> _____________ ReadLongPtr / MemLongPtr _______,
'         SAFEARRAY _______________ AV (= SEH __) .
'         __. __ AV _ On Error Resume Next _______ Excel .
'         ____ (____ 10). ____ 55 ______.
'     - __ Private Function: IsProtectReadWritable(prot) As Boolean
'         (PAGE_GUARD _____ + PAGE_READWRITE / PAGE_EXECUTE_READWRITE __)
'     - ______ (Handler_xxx / dcf / Thunks_Init _) .
'
'   _9.9c ____ (__, ___):
'     - Test_ExecuteScript_Parallel _____ (__ D)
'         5 __ ExecuteScript ______, callback ID .
'         _____ (_ a, 9.8c) _________________.
'         ____ ((1+1), (2+1), ... ) ____ callback ID _________.
'     - ______ (Handler_xxx / dcf / Thunks_Init _) .
'         (9.9c _ Wv2Thunks _______ 1 _______)
'
'   _9.9b ____ (__, SPA __):
'     - HandlerKind _ 2 ____ (HK_HistoryChanged=9 / HK_DOMContentLoaded=10)
'     - m_iidTable ______ HK_DOMContentLoaded .
'     - InitIIDTable _ 2 IID .
'     - GetLongLongProperty ___ (#If Win64 ___ = 64bit __)
'     - BuildSpaTestHtml / Test_SpaTest_Real ___ (SPA ____)
'     - _ R': BuildSpaTestHtml _ doPush/doReplace _ try-catch .
'              (about:blank origin _ SecurityError ______)
'
'   * _9.9 (a + b + c) _ 32bit ________ (__ X) *
'     ___ 64bit __ (#Else #Error __) ________,
'     ______ LongLong _______32bit _____________.
'     _______ x86 _____________ (__) _____.
'     EventRegistrationToken _ x86 __ 64bit (LongLong) ___________.
'
'   _9.9a ____ (__):
'     - FillGUID _ Private -> Public .
'         (Wv2Pane.EnsureView2 _____ GUID ______________)
'     - ICoreWebView2_2 _______ Wv2Pane ____ (m_pView2 + EnsureView2)
'
'   _9.8c ____ (__):
'     - HandlerKind _ HK_ExecuteScriptCompleted = 8 .
'         (1 _____, _____, _ a __)
'     - m_iidTable ______ HK_NewWindowRequested -> HK_ExecuteScriptCompleted .
'     - InitIIDTable _ IID_ICoreWebView2ExecuteScriptCompletedHandler .
'         (49511172-CC67-4BCA-9923-137112F4C4CC)
'     - Handler_QueryInterface / Handler_AddRef / Handler_Release / Handler_Invoke
'       ___________ (kind ___________ kind=8 _______)
'
'   _9.8a/9.8b ____ (__):
'     - ______ Module9_9_9 -> Wv2Thunks _____ (9.8a)
'     - ___ m_handlers() As ComCallbackHandler (9.8a)
'     - _9.8b __ Wv2Thunks _____ (Settings __ Wv2Pane ____)
'
'   ________:
'     WebView2 ______________. ________ (VirtualAlloc),
'     ______ (m_handlers / m_freeSlots), ______ (HandlerAddRefInternal /
'     HandlerReleaseInternal), __________ Handler_QueryInterface /
'     Handler_AddRef / Handler_Release / Handler_Invoke, IID .
'     (m_iidTable + IsEqualGUIDInPlace, _9.7a), .
'     VBA __________ (_9.5), Win32 API _____ (_9.7b .
'     ________ API ___) ___.
'
'   ____________ (_______.md __):
'     _9.2: ________ (DispCallFunc -> VBA __________)
'     _9.3: _____ (Wv2Environment / Wv2Pane) ___, _ D1-.
'     _9.4: ________ (_ P1), HandlerKind ____ (_ r)
'     _9.5: ___________________, &H .
'            (MEM_RELEASE = &H8000& __)
'     _9.6: NewWindowRequested ___, ____________ (View_/Ctrl_)
'     _9.7a: Handler_QueryInterface ____ (IID ______)
'     _9.7b: UserForm ________ API (GetAncestor / SetWindowLongPtrW /
'             SetWindowPos) _____ (GA_ROOT / WS_THICKFRAME _)
'
'   * _9.7a __________ *
'     Test_AllEvents_Real ________, _________:
'     1. WebView2 ______________________,
'        IID_IUnknown ____ IID ___ QueryInterface .
'     2. WebView2 ______ QI __ vtable ______, Invoke .
'        QI ____ Invoke.
'     3. 9.6b _____ IID __________________,
'        WebView2 _ QI ____________, ______ IMarshal
'        .
'
'   * _____________ (_9.7a __) *
'     ________ (= _______, Invoke ______,
'     AddRef/Release ___, Cleanup _____, _) _____:
'       (1) Handler_QueryInterface .
'           _ppvObject = this : HandlerAddRefInternal this :
'            Handler_QueryInterface = S_OK : Exit Function.
'           _________ 9.6b .
'       (2) ____________, IID .
'       (3) ______ IID ________, .
'
'   _____ (_9.9c __ = _9.9 __________):
'     - TabWebView2 _____ (_9.10 _____, ____)
'     - _ SPA _____ (http(s) origin __ pushState/replaceState ____, 9.10 __)
'     - _______ IF (ICoreWebView2_3 __, ______)
'     - 32bit (x86) __: _______ x86 _ + ____ (______)
'     _ SPA __ (HistoryChanged / DOMContentLoaded / postMessage ______) .
'        ExecuteScript ___ (D + E + F + G) _______ (_9.9c) _____.
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/refTypes/discussions/3
''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize As Long = 8
#Else
    Private Const NullPtr As Long = 0&
    Private Const PtrSize As Long = 4
#End If

' --- SafeArray / PointerAccessor ____ ---
Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

' --- WebView2Loader.dll ---
'   _9.3a ___ Public Declare ___ (Wv2Environment.Init ______)
Public Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

' --- Win32 API ---
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As LongPtr

Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long

' --- DispCallFunc (oleaut32) ---
'   _9.3a _______.
'   ICoreWebView2Environment::Release ___ VBA .
'   ______ (ComRelease / ComAddRef) _____.
'
'   ByRef ... As Any ________, __ 0 .
'   _______ByVal 0&_____ NULL ___.
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Any, _
    ByRef prgpvarg As Any, _
    ByRef pvargResult As Any) As Long

' --- user32: GetClientRect (_9.3c _____) ---
'   HWND ______________. _____ (0,0,width,height).
'   put_Bounds _______________.
'   Wv2Pane _________ Public Declare.
Public Declare PtrSafe Function GetClientRect Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByRef lpRect As RECT) As Long

' --- UserForm ________ API _ (_9.7b _____) ---
'   UserForm1.frm ___Frame1 _ HWND -> UserForm ___ HWND_____,
'   UserForm ___ HWND _ WS_THICKFRAME (= __ = ________) .
'   _________.
'
'   Declare _ Module9 ___ (______ Win32 API _ Module9 _____).
'   ___ Public Const ___ Module9 ___ (UserForm1 ______).
'
'   GetAncestor:
'     ____ HWND ___ HWND _____.
'     gaFlags = GA_ROOT (=2) ____, __ HWND .
'     (= ___________________/___________) .
'     ____. Frame1 -> UserForm ___ HWND ______.
'
'   GetWindowLongPtrW / SetWindowLongPtrW:
'     __________ (Window Style _) ___/____.
'     ____ nIndex _ GWL_STYLE (=-16) ____, .
'     LongPtr ___/_____ (x64 __ LONG_PTR _, x86 __ LONG _,
'     PtrSafe + LongPtr _________).
'     WS_THICKFRAME (=&H40000) _ OR _____, _____________.
'
'   SetWindowPos:
'     ________/___/Z ____/________.
'     WS_THICKFRAME ___, SWP_FRAMECHANGED _________,
'     _____________ (= _______) ______.
'     __/___/Z _____________ SWP_NOMOVE | SWP_NOSIZE |
'     SWP_NOZORDER _______.
Public Declare PtrSafe Function GetAncestor Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal gaFlags As Long) As LongPtr

Public Declare PtrSafe Function GetWindowLongPtrW Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIndex As Long) As LongPtr

Public Declare PtrSafe Function SetWindowLongPtrW Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As LongPtr) As LongPtr

Public Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal uFlags As Long) As Long

' --- UserForm ____________ (_9.7b _____) ---
'   GetAncestor / GetWindowLongPtrW / SetWindowLongPtrW / SetWindowPos .
'   ____. Win32 ______________.
'
'   ____ Integer ________, ____ 29 .
'   ____ &H... __ & __________ (Integer .
'   _______, Long _______).
Public Const GA_ROOT          As Long = 2
Public Const GWL_STYLE        As Long = -16
Public Const WS_THICKFRAME    As Long = &H40000
Public Const SWP_NOSIZE       As Long = &H1&
Public Const SWP_NOMOVE       As Long = &H2&
Public Const SWP_NOZORDER     As Long = &H4&
Public Const SWP_FRAMECHANGED As Long = &H20&

' --- ________ API (_9.4a _____, _ mWebView2Helper __) ---
'   LPWSTR (UTF-16 _______) _ VBA String __________.
'   ________:
'     lstrlenW       : LPWSTR ____ (NULL _____) .
'     CoTaskMemFree  : COM _ out __ LPWSTR* (CoTaskMemAlloc .
'                      ___) _____. COM ___, ________.
'     RtlMoveMemory  : _______ (_________)
Public Declare PtrSafe Function lstrlenW Lib "kernel32" ( _
    ByVal lpString As LongPtr) As Long

Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32" ( _
    ByVal pv As LongPtr)

Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal destination As LongPtr, _
    ByVal source As LongPtr, _
    ByVal length As LongPtr)

' --- _______ (N-2 ___) ---
'   UTF-8 ______ UTF-16 (VBA _ String) ___.
'   *_______________* StrConv(b, vbUnicode) .
'   _____________ (CP932) _______ UTF-8 _____.
Public Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal codePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long) As Long

Public Const CP_UTF8 As Long = 65001

' --- ____ API (_9.5 _____, ________) ---
'   _____________ Thunks_Init .
'   __________. MouseScroll.bas (Cristian Buse _) _____.
'
'   W _ (UTF-16) ___. StrPtr() _ VBA String ___ UTF-16 .
'   _________, A/W ________.
'
'   SetEnvironmentVariableW(lpName, NULL) _____________.
'   VBA __ NULL ______________ 0 .
'   (LongPtr _ ByVal ___ 0 _ NULL _______).
Private Declare PtrSafe Function GetEnvironmentVariableW Lib "kernel32" ( _
    ByVal lpName As LongPtr, _
    ByVal lpBuffer As LongPtr, _
    ByVal nSize As Long) As Long

Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" ( _
    ByVal lpName As LongPtr, _
    ByVal lpValue As LongPtr) As Long

' --- Win32 _____ (_9.5 _____, ____) ---
'   ___ Win32 API ______________________.
'   VirtualFree _ 0 ___________ (ERROR_BUSY=170 /
'   ERROR_INVALID_ADDRESS=487 / ERROR_INVALID_PARAMETER=87 __) ___.
'
'   __: GetLastError _ per-thread ______. VBA __ Win32 API .
'         ____________, _____ API __________.
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

' --- VirtualQuery ____ (_9.5 _____) ---
'   x64 ____ MEMORY_BASIC_INFORMATION ______ (__ 48 bytes):
'     +0  BaseAddress       LongPtr (8)
'     +8  AllocationBase    LongPtr (8)
'     +16 AllocationProtect Long    (4)
'     +20 (pad1)            Long    (4) <- .
'     +24 RegionSize        LongPtr (8)
'     +32 State             Long    (4)
'     +36 Protect           Long    (4)
'     +40 Type              Long    (4)  (VBA _________ Type_ ___)
'     +44 (pad2)            Long    (4) <- .
'   VBA Type ____ + _________________, .
'   pad1/pad2 ______________________.
#If Win64 Then
Private Type MEMORY_BASIC_INFORMATION
    BaseAddress       As LongPtr
    AllocationBase    As LongPtr
    AllocationProtect As Long
    pad1              As Long
    RegionSize        As LongPtr
    state             As Long
    Protect           As Long
    Type_             As Long
    pad2              As Long
End Type
#Else
Private Type MEMORY_BASIC_INFORMATION
    BaseAddress       As LongPtr
    AllocationBase    As LongPtr
    AllocationProtect As Long
    RegionSize        As LongPtr
    state             As Long
    Protect           As Long
    Type_             As Long
End Type
#End If

' --- VirtualQuery (_9.5 _____, ____) ---
'   _______________________.
'   __________ MEMORY_BASIC_INFORMATION _____ (___ 0).
'   Type ______________, Declare ________.
Private Declare PtrSafe Function VirtualQuery Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByRef lpBuffer As MEMORY_BASIC_INFORMATION, _
    ByVal dwLength As LongPtr) As LongPtr

' MEMORY_BASIC_INFORMATION.State __:
'   MEM_COMMIT  = &H1000  (4096)   <- ___ MEM_COMMIT .
'   MEM_RESERVE = &H2000  (8192)   <- ___ MEM_RESERVE .
'   MEM_FREE    = &H10000 (65536)
' MEMORY_BASIC_INFORMATION.Protect ___ FormatMemProtect .
' __________.

' --- RECT ___ (_9.3c _____) ---
'   ICoreWebView2Controller::put_Bounds _____ 16 bytes ___.
'   x64 ABI ____16 bytes _____________, .
'   ___ VarPtr(rect) _________________, .
'   ____ (= _______ dcf _ LongPtr _______) ___.
'   Wv2Pane _________ Public Type.
Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Const MEM_COMMIT             As Long = &H1000&
Private Const MEM_RESERVE            As Long = &H2000&
' _ 9.5g ____&H8000_->_&H8000&____. .
'   Integer _ &H8000 (= -32768) _________ Long ________,
'   VirtualFree _ -32768 _______ ERROR_INVALID_PARAMETER _______.
'   ____ VBA 7 + x64 __________________.
Private Const MEM_RELEASE            As Long = &H8000&
' MEM_FREE (_9.10a __): VirtualQuery _ State ___. _________.
'   &H10000 = 65536. ______ & ______ (MEM_RELEASE ________).
Private Const MEM_FREE               As Long = &H10000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&
' PAGE_READWRITE (_9.10a __): Sentinel ___________________.
'   ______ PAGE_EXECUTE_READWRITE (&H40) _______, .
'   ____ PAGE_READWRITE (&H4) _____________.
Private Const PAGE_READWRITE         As Long = &H4&
Private Const S_OK                   As Long = 0
Private Const E_NOINTERFACE          As Long = &H80004002
Private Const CC_STDCALL             As Long = 4

' --- __________ (_9.5 _____) ---
'
'   SENTINEL_ENV_NAME:
'     ___________. 1 _______ Excel .
'     _____________, .
'     ____________ (= __________).
'
'   SENTINEL_BUFFER_SIZE:
'     GetEnvironmentVariableW _____________ (WCHAR __).
'     x64 ______ CStr ____________ 19 .
'     (2^63 - 1 = 9223372036854775807), _________ 20 .
'     ____, __________ 32.
Private Const SENTINEL_ENV_NAME    As String = "WV2_VBA_LastRegion"
Private Const SENTINEL_BUFFER_SIZE As Long = 32

' --- ____ / _________ ---

#If Win64 Then
Private Const LATE_BIND_OFFSET As Long = 55
Private Const STUB_LEN As Long = 91
Private Const THUNK_LEN As Long = 74
#Else
' 32-bit VBE7.DLL EntryPoint stub (verified on VBE7 7.01.1118):
'   +0..+20 : EbMode gate (skips jump when EbMode == 2 Break mode)
'   +21     : B8 <imm32>  MOV EAX, targetAddr  (imm32 at +22)
'   +26     : FF E0       JMP EAX
'   +28     : 33 C0 C2 00 00  (note: patched to C2 0C 00 so Break mode cleans 3 DWORDs!)
Private Const LATE_BIND_OFFSET As Long = 22
Private Const STUB_LEN As Long = 33
Private Const THUNK_LEN As Long = 46
#End If

' ___________ (___ 91 + _____ 5 = 96)
Private Const THUNK_OFFSET   As Long = 96

' vtableObj ________ (= _ SLOT_SIZE __________)
'   ______ WebView2VTable ___ (40 bytes) ___:
'     +0..+7   : pVTable (= ___ +8 ___)
'     +8..+15  : Functions(0) = Handler_QueryInterface
'     +16..+23 : Functions(1) = Handler_AddRef
'     +24..+31 : Functions(2) = Handler_Release
'     +32..+39 : Functions(3) = pSlot (_________)
'
'   _9.3a ___ Public Const ___ (Wv2Environment.Init ____)
Public Const VTABLE_OBJ_OFFSET As Long = 176

' 1 _____________ (8 _______ 224 bytes)
'   +0..+90      _______ (91 bytes)
'   +91..+95     .
'   +96..+169    ___ (74 bytes)
'   +170..+175   _____ (8 byte _______)
'   +176..+183   vtableObj.pVTable
'   +184..+215   vtableObj.Functions(0..3)
'   +216..+223   _______ (8 byte _______)
Private Const SLOT_SIZE      As Long = 224

' _____ (__)
Private Const SLOT_COUNT     As Long = 512

' ___________ (_____ + ____)
'   +0          : _____ (1 byte, ___ 8 byte _____)
'   +1..+63     : __ (____________________)
Private Const HEADER_SIZE    As Long = 64

' ________ (= 64 + 224_512 = 114,752 bytes _ 112 KB)
Private Const REGION_SIZE    As Long = HEADER_SIZE + SLOT_SIZE * SLOT_COUNT


' ============================================================
' GUID ___ (Win32 __, _9.7a _____)
'
'   COM _ REFIID _______ 16 byte ____.
'   QueryInterface _ riid ______________.
'
'   _____ x64 _____ (16 bytes, Win32 __):
'     +0..+3   : Data1     (Long, little-endian)
'     +4..+5   : Data2     (Integer, little-endian)
'     +6..+7   : Data3     (Integer, little-endian)
'     +8..+15  : Data4(0..7) (_ Byte _____, big-endian _)
'
'   _: IID {4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d} .
'     Data1 = &H4E8A3389
'     Data2 = &HC9D8 (Integer ______________ &HC9D8 ___)
'     Data3 = &H4BD2
'     Data4 = (&HB6, &HB5, &H12, &H4F, &HEE, &H6C, &HC1, &H4D)
'
'   GUID ___ memcmp __ (16 byte _______) ___. .
'   LongLong _ 2 ____________ IsEqualGUIDInPlace ___.
'
'   Public _______, ________ IF QueryInterface .
'   __ (_9.8 __________ IF _) ___________.
' ============================================================
Public Type GUID
    data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


' ============================================================
' HandlerKind ___ (_____________)
'   WebView2 ______________. ComCallbackHandler.cls .
'   m_kind _____, Handler_Invoke _________.
' ============================================================
Public Enum HandlerKind
    HK_None = 0
    HK_EnvironmentCompleted = 1
    HK_ControllerCompleted = 2
    HK_NavigationStarting = 3
    HK_NavigationCompleted = 4
    HK_WebMessageReceived = 5
    HK_DocumentTitleChanged = 6
    HK_NewWindowRequested = 7
    HK_ExecuteScriptCompleted = 8   ' _9.8c ___, 1 _____, .
    HK_HistoryChanged = 9           ' _9.9b ___, __, __ ICoreWebView2 (vtable 13/14)
    HK_DOMContentLoaded = 10        ' _9.9b ___, __, ICoreWebView2_2 (vtable 64/65)
    HK_WebResourceRequested = 11    ' N-1 ___, __, __ ICoreWebView2 (vtable 55/56)
    HK_WebResourceResponseReceived = 12  ' N-3 ___, __, ICoreWebView2_2 (vtable 61/62)
    HK_ResponseGetContent = 13      ' N-4 ___, *1 ___________*
End Enum


' --- ____ 1 _______________ ---
'   ______ vtableObj.Functions(0..2) ______.
'   AddressOf __________________, GetAddr _______.
'   Thunks_Init _ 1 ______, Acquire __________.
Private m_pHandler_QI       As LongPtr
Private m_pHandler_AddRef   As LongPtr
Private m_pHandler_Release  As LongPtr

' --- _________ ---
Private m_pRegionBase As LongPtr        ' VirtualAlloc .
Private m_freeHead    As Long           ' _______________ (-1 __)
Private m_freeNext()  As Long           ' freeNext(i) = i .
Private m_inUse       As Long           ' ___________ (_____)

' --- ____ index -> Handler __________ ---
'   Handler_AddRef / Release / QueryInterface .
'   _this (= pSlot + VTABLE_OBJ_OFFSET) -> idx (SlotIndexFromVTableObjAddr)
'    -> m_handlers(idx)_________ ComCallbackHandler ___________.
'   refcount = 0 ___________________.
Private m_handlers(0 To SLOT_COUNT - 1) As ComCallbackHandler

' --- IID ____ (_9.7a _______) ---
'   HandlerKind __________ IID______.
'   Handler_QueryInterface _ riid ________, .
'   S_OK + __, _____ ppvObject = 0 + E_NOINTERFACE ___.
'
'   Thunks_Init ____ InitIIDTable __________.
'   _______ HandlerKind ____ (HK_None = 0 ? HK_ResponseGetContent = 13).
'   HK_None _________ (____________, ___ 0 ___).
Private m_iidTable(HK_None To HK_ResponseGetContent) As GUID

' --- IID_IUnknown (Win32 __, Thunks_Init ____) ---
'   {00000000-0000-0000-C000-000000000046}. Handler_QueryInterface .
'   _____. ___________ COM ___.
Private m_iidIUnknown As GUID

Private Const THUNK_BUF_SIZE As Long = 80     ' 74 _ 8 .


' ============================================================
' EntryPoint .
'   ___________ VirtualAlloc ________.
'   Force-Compile _______ 1 _____________.
' ============================================================
Private Sub EntryPoint(): End Sub
' ============================================================
' AcquireHandlerFor (_9.3a _____)
'
'   _____ (Wv2Environment __) ____________ 1 .
'   ___________. __________:
'     1. Thunks_Init() ___ (____)
'     2. ComCallbackHandler _ New
'     3. ComCallbackHandler.Handler_Invoke ______ vTable .
'     4. Thunks_AcquireSlot __________ + _______ + m_handlers .
'     5. ComCallbackHandler.Init(kind, owner, pSlot) .
'     6. ______ ComCallbackHandler .
'
'   __________________, _____________.
'   ___, ______ ComCallbackHandler.Slot + VTABLE_OBJ_OFFSET _ WebView2 .
'   _______________.
'
'   ___ (______) _ Nothing ___. ______ Nothing ______.
' ============================================================
Public Function AcquireHandlerFor( _
    ByVal kind As HandlerKind, _
    ByVal owner As Object) As ComCallbackHandler

    ' --- 1. ______ ---
    If m_pRegionBase = 0 Then
        If Not Thunks_Init() Then Exit Function
    End If

    ' --- 2. ComCallbackHandler _ New ---
    Dim h As ComCallbackHandler
    Set h = New ComCallbackHandler

    ' --- 3. ComCallbackHandler.Handler_Invoke ________ ---
    Dim pHandlerInvoke As LongPtr
    pHandlerInvoke = GetClassMethodAddrAtFixedSlot(h, 7)
    If pHandlerInvoke = 0 Then Exit Function

    ' --- 4. ______ + _______ + m_handlers __ ---
    Dim pSlot As LongPtr
    pSlot = Thunks_AcquireSlot(h, ObjPtr(h), pHandlerInvoke)
    If pSlot = 0 Then Exit Function

    ' --- 5. ComCallbackHandler ______ ---
    h.Init kind, owner, pSlot

    ' --- 6. _____ ComCallbackHandler ___ ---
    Set AcquireHandlerFor = h
End Function


Public Sub EnsureFolder(ByVal path As String)
    If LenB(path) = 0 Then Exit Sub

    Dim cleanPath As String
    cleanPath = path
    If Right$(cleanPath, 1) = "\" Then cleanPath = Left$(cleanPath, Len(cleanPath) - 1)
    If Len(cleanPath) <= 3 Then Exit Sub

    On Error Resume Next
    If Len(Dir$(cleanPath, vbDirectory)) > 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Dim slashPos As Long
    slashPos = InStrRev(cleanPath, "\")
    If slashPos > 3 Then
        EnsureFolder Left$(cleanPath, slashPos - 1)
    End If

    On Error Resume Next
    MkDir cleanPath
    On Error GoTo 0
End Sub


' ============================================================
' FinalizeWebView2Environment (_9.3a ______)
'
'   ___ gHandler _ Nothing _____________,
'   gHandler ________, Thunks_Shutdown _________.
'
'   .
'   _____. ___________ Test _____ Thunks_Shutdown .
'   ____ (Test_TwoHandlers_Mock _______).
' ============================================================
Public Sub FinalizeWebView2Environment()
    Thunks_Shutdown
End Sub


' ============================================================
' dcf (_9.3b _____)
'
'   __ DispCallFunc ____. _____ (ParamArray) .
'   __ 0?N __ COM ________.
'
'   ___ dcf ______:
'     - _CopyMemory ___ pVTable / pMethod _________.
'       _________________________ (DispCallFunc
'       _______________).
'     - _Case Else __ CLngPtr(CLng(...))____. .
'       Case vbLong: ___________________.
'
'   __:
'     pInterface : COM ____________ (this)
'     vtblIndex  : vtable _______ (IUnknown _ 0=QI, 1=AddRef, 2=Release,
'                  _______________________ 3 ____)
'     funcName   : _____ (__________)
'     args       : ____________ (LongPtr/Long/Double/___)
'
'   _________ (VarType ___):
'     vbLong       -> 4 bytes .
'     vbLongLong   -> 8 bytes ____/_______ (Win64 LongPtr __)
'     vbDouble     -> 8 bytes .
'     _____   -> vbLongLong __ (LongPtr _______)
'
'   ___: _____ COM ________ (Long _____)
'           _DispCallFunc __________ hr ___, Debug.Print .
'           _____ COM .
'             (HRESULT ____, AddRef/Release ____ ULONG .
'              ___, ______________)
'           _____ = HRESULT_________________:
'             AddRef/Release _______________,
'             _ _ 0 ___________________ (_______).
'             _______________ HRESULT .
'             ___________.
'
'   __: out __ (_: pp* ________) _____________,
'         ______ Variant _____ LongPtr ___ VarPtr _____.
'         Variant ________________, ______ CLngPtr .
'         ____. _9.3c ___ Controller::get_CoreWebView2 _____.
' ============================================================
Public Function dcf( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    ByVal funcName As String, _
    ParamArray args() As Variant) As Long

    If pInterface = 0 Then
        Debug.Print "dcf: null interface - " & funcName
        dcf = &H80004003   ' E_POINTER
        Exit Function
    End If

    Dim argc As Long
    argc = UBound(args) - LBound(args) + 1
    If argc < 0 Then argc = 0

    Dim res As Variant
    Dim hr As Long

    If argc = 0 Then
        ' ____: vt/vp _ NULL .
        hr = DispCallFunc(pInterface, vtblIndex * PtrSize, _
                          CC_STDCALL, vbLong, _
                          0, ByVal 0&, ByVal 0&, res)
    Else
        ' ____: vt / vp .
        Dim vt() As Integer
        Dim vp() As LongPtr
        Dim vals() As Variant
        ReDim vt(0 To argc - 1)
        ReDim vp(0 To argc - 1)
        ReDim vals(0 To argc - 1)

        Dim i As Long
        For i = 0 To argc - 1
            vals(i) = args(LBound(args) + i)
            Select Case VarType(vals(i))
                Case vbLong:     vt(i) = vbLong
                Case vbDouble:   vt(i) = vbDouble
                Case vbCurrency: vt(i) = vbCurrency
                #If Win64 Then
                Case vbLongLong: vt(i) = vbLongLong   ' = vbLongPtr on Win64
                Case Else:       vt(i) = vbLongLong   ' _______ (LongPtr __)
                #Else
                Case Else:       vt(i) = vbLong       ' 32bit LongPtr = vbLong
                #End If
            End Select
            vp(i) = VarPtr(vals(i))
        Next i

        hr = DispCallFunc(pInterface, vtblIndex * PtrSize, _
                          CC_STDCALL, vbLong, _
                          argc, vt(0), vp(0), res)
    End If

    If hr <> 0 Then
        If LenB(funcName) > 0 Then _
            Debug.Print "dcf CALL failed: " & funcName & " hr=&H" & Hex(hr)
        dcf = hr
    Else
        ' ____ COM _______________.
        ' HRESULT __________, ______ <> 0 ______________.
        ' AddRef/Release ____ ULONG ______, .
        ' (_______<> 0 _____________).
        dcf = CLng(res)
    End If
End Function


' ============================================================
' ComRelease / ComAddRef
'
'   IUnknown::Release / AddRef ________.
'   _9.3b _ dcf _______, ______ Public Function .
'   (Wv2Environment.Class_Terminate ___________Release .
'   _____________).
' ============================================================
Public Function ComRelease(ByVal pInterface As LongPtr) As Long
    If pInterface <> 0 Then ComRelease = dcf(pInterface, 2, "Release")
End Function

Public Function ComAddRef(ByVal pInterface As LongPtr) As Long
    If pInterface <> 0 Then ComAddRef = dcf(pInterface, 1, "AddRef")
End Function


' ============================================================
' PtrToString (_9.4a _____, _ mWebView2Helper __)
'
'   LPWSTR (UTF-16 ____________) _ VBA _ String _____.
'   COM _ out ________ LPWSTR ________.
'
'   ____:
'     RtlMoveMemory _______________ (_ _).
'     ____________CopyMemory ____________,
'     PointerAccessor _ 1 ____ ChrW _____________,
'     URL ___________. _9.4a _______.
'
'   __:
'     ______ p _ CoTaskMemAlloc __________.
'     ______ PtrToString _ String ______, .
'     CoTaskMemFree(p) __________.
'     GetStringProperty ______ -> __ -> __________.
' ============================================================
Public Function PtrToString(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function
    Dim cch As Long
    cch = lstrlenW(p)
    If cch = 0 Then Exit Function
    PtrToString = String$(cch, vbNullChar)
    RtlMoveMemory StrPtr(PtrToString), p, CLngPtr(cch * 2)
End Function


' ============================================================
' Utf8BytesToString (N-2 _____)
'
'   UTF-8 ______ VBA _ String _____.
'   ________ (IStream _________) _____________.
'
'   __:
'     b  : _____ (0 __)
'     cb : ________ (____________ = ____________)
'
'   ___: _______. cb <= 0 _____.
'
'   *______________*
'     MB_ERR_INVALID_CHARS _____ (dwFlags = 0). .
'     *_________________*___________,
'     ___________________________.
'     ______ U+FFFD _____________.
'
'   *32bit _____* ____ (LongPtr __).
' ============================================================
Public Function Utf8BytesToString(ByRef b() As Byte, ByVal cb As Long) As String
    If cb <= 0 Then Exit Function

    Dim cch As Long
    cch = MultiByteToWideChar(CP_UTF8, 0, VarPtr(b(LBound(b))), cb, 0, 0)
    If cch <= 0 Then Exit Function

    Dim outStr As String
    outStr = String$(cch, vbNullChar)
    cch = MultiByteToWideChar(CP_UTF8, 0, VarPtr(b(LBound(b))), cb, StrPtr(outStr), cch)
    If cch <= 0 Then Exit Function

    Utf8BytesToString = Left$(outStr, cch)
End Function


' ============================================================
' GetStringProperty (_9.4a _____, _ mWebView2Helper __)
'
'   COM _________ get_*String ____________ String ___.
'   ______: vtable ______ HRESULT get_Xxx([out, retval] LPWSTR *value)
'   _______________.
'
'   __:
'     1. dcf _ out __ LPWSTR* ___ (pStr _______)
'     2. PtrToString _ VBA String .
'     3. CoTaskMemFree _ LPWSTR ___ (COM ___, ________)
'
'   __:
'     pInterface : COM ____ (_: ICoreWebView2NavigationStartingEventArgs*)
'     vtblIndex  : _________ vtable index (_: get_Uri _ 3)
'     funcName   : _____ (__________, HRESULT _________)
'
'   ___: _______ (_________ or ____)
'
'   _9.4a _____:
'     Dim url As String
'     url = GetStringProperty(pArgs, 3, "NavStartingArgs.get_Uri")
' ============================================================
Public Function GetStringProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As String

    If pInterface = 0 Then Exit Function

    Dim pStr As LongPtr
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(pStr))
    If hr = 0 And pStr <> 0 Then
        GetStringProperty = PtrToString(pStr)
        CoTaskMemFree pStr
    End If
End Function


' ============================================================
' GetBoolProperty (_9.4b _____, _ mWebView2Helper __)
'
'   COM _________ get_*Bool ____________ Boolean ___.
'   vtable _____: HRESULT get_Xxx([out, retval] BOOL *value)
'
'   BOOL _ Win32 ___ 4 bytes __ (0 = FALSE, ____ = TRUE).
'   VBA _ Boolean _______ 0 _____.
'
'   _9.4b _____:
'     Dim isOk As Boolean
'     isOk = GetBoolProperty(pArgs, 3, "NavCompletedArgs.get_IsSuccess")
' ============================================================
Public Function GetBoolProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As Boolean

    If pInterface = 0 Then Exit Function

    Dim value As Long
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(value))
    If hr = 0 Then GetBoolProperty = (value <> 0)
End Function


' ============================================================
' GetLongProperty (_9.4b _____, _ mWebView2Helper __)
'
'   COM _________ get_*Long ____________ Long ___.
'   vtable _____: HRESULT get_Xxx([out, retval] long *value)
'
'   ___ (_: COREWEBVIEW2_WEB_ERROR_STATUS) ____________.
' ============================================================
Public Function GetLongProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As Long

    If pInterface = 0 Then Exit Function

    Dim value As Long
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(value))
    If hr = 0 Then GetLongProperty = value
End Function


' ============================================================
' GetLongLongProperty (_9.9b _____)
'
'   COM _________ get_*UINT64 / get_*INT64 .
'   LongLong (____ 64 bit) ___.
'   vtable _____: HRESULT get_Xxx([out, retval] UINT64 *value)
'
'   __:
'     ICoreWebView2DOMContentLoadedEventArgs.get_NavigationId (UINT64) ___.
'     NavigationId ___________, ___ LongLong _________.
'
'   * 32bit _____ (__ X, _ a) *
'     32bit VBA __ LongLong _________, ____ #If Win64 .
'     ___ 64bit _______. 32bit ______, UINT64 .
'     _________. __________:
'       (a) Currency _ (8 byte ____) ____ 10000 .
'       (b) Long 2 _ (lo/hi) .
'       (c) NavigationId __________ 32bit __ Long .
'     _______ 32bit _____________ (________).
' ============================================================
#If Win64 Then
Public Function GetLongLongProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As LongLong

    If pInterface = 0 Then Exit Function

    Dim value As LongLong
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(value))
    If hr = 0 Then GetLongLongProperty = value
End Function
#Else
Public Function GetLongLongProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As Currency

    If pInterface = 0 Then Exit Function

    Dim value As Currency
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(value))
    If hr = 0 Then GetLongLongProperty = value
End Function
#End If


' ============================================================
' Test_TwoHandlers_Mock (_9.2______________)
'
'   _9.2__________: 2 _______ Acquire _,
'   _____ vtableObj _________ AddRef/Release ____,
'   idx _________________. WebView2 .
'   (_______________).
'
'   _9.3a __________:
'     - gHandler ________ pass ____ (_______)
'     - owner = Nothing _ ComCallbackHandler _ Acquire .
'     - HandlerReleaseInternal _________gHandler Is h.
'       .
'
'   ______ Wv2Environment _______. AcquireHandlerFor ____,
'   ____ Thunks_AcquireSlot .
'   (__________________).
' ============================================================
Public Sub Test_TwoHandlers_Mock()

    Debug.Print String(60, "=")
    Debug.Print "Test_TwoHandlers_Mock __"
    Debug.Print String(60, "=")

    ' --- 1. ______ ---
    If Not Thunks_Init() Then
        Debug.Print "Thunks_Init __"
        Exit Sub
    End If

    Debug.Print "Thunks_Init OK, m_inUse = " & m_inUse

    ' --- 2. 2 _________ + Acquire ---
    Dim h1 As ComCallbackHandler, h2 As ComCallbackHandler
    Set h1 = New ComCallbackHandler
    Set h2 = New ComCallbackHandler

    Dim pInvoke1 As LongPtr, pInvoke2 As LongPtr
    pInvoke1 = GetClassMethodAddrAtFixedSlot(h1, 7)
    pInvoke2 = GetClassMethodAddrAtFixedSlot(h2, 7)

    Dim pSlot1 As LongPtr, pSlot2 As LongPtr
    pSlot1 = Thunks_AcquireSlot(h1, ObjPtr(h1), pInvoke1)
    pSlot2 = Thunks_AcquireSlot(h2, ObjPtr(h2), pInvoke2)

    If pSlot1 = 0 Or pSlot2 = 0 Then
        Debug.Print "Acquire __"
        Exit Sub
    End If

    h1.Init HK_EnvironmentCompleted, Nothing, pSlot1
    h2.Init HK_ControllerCompleted, Nothing, pSlot2

    Dim idx1 As Long, idx2 As Long
    idx1 = SlotIndexFromAddr(pSlot1)
    idx2 = SlotIndexFromAddr(pSlot2)
    Debug.Print "h1 -> idx " & idx1 & ", pSlot " & pSlot1
    Debug.Print "h2 -> idx " & idx2 & ", pSlot " & pSlot2
    Debug.Print "m_inUse = " & m_inUse & " (___ 2)"

    If idx1 = idx2 Then
        Debug.Print "[NG] __________ idx _____"
        Exit Sub
    End If

    ' --- 3. ______ vtableObj _________ ---
    Dim pObj1 As LongPtr, pObj2 As LongPtr
    pObj1 = pSlot1 + VTABLE_OBJ_OFFSET
    pObj2 = pSlot2 + VTABLE_OBJ_OFFSET

    ' --- 4. ___ AddRef ____, ________________ ---
    Debug.Print "--- AddRef sequence ---"
    Dim r As Long
    r = Handler_AddRef(pObj1)
    Debug.Print "  Handler_AddRef(pObj1) -> " & r
    r = Handler_AddRef(pObj1)
    Debug.Print "  Handler_AddRef(pObj1) -> " & r
    r = Handler_AddRef(pObj2)
    Debug.Print "  Handler_AddRef(pObj2) -> " & r
    Debug.Print "h1.RefCount = " & h1.RefCount & " (___ 2)"
    Debug.Print "h2.RefCount = " & h2.RefCount & " (___ 1)"

    If h1.RefCount <> 2 Or h2.RefCount <> 1 Then
        Debug.Print "[NG] AddRef _____________"
    Else
        Debug.Print "[OK] AddRef __________"
    End If

    ' --- 5. Release ____, _____ refcount=0 ________ ---
    Debug.Print "--- Release sequence ---"
    r = Handler_Release(pObj1)
    Debug.Print "  Handler_Release(pObj1) -> " & r
    Debug.Print "h1.RefCount = " & h1.RefCount & " (___ 1)"
    Debug.Print "m_inUse = " & m_inUse & " (___ 2, _______)"

    r = Handler_Release(pObj1)
    Debug.Print "  Handler_Release(pObj1) -> " & r
    Debug.Print "(h1 ___) m_inUse = " & m_inUse & " (___ 1)"

    Debug.Print "h1.RefCount (___) = " & h1.RefCount & " (___ 0)"
    Debug.Print "h2.RefCount (h1_____) = " & h2.RefCount & " (___ 1)"

    r = Handler_Release(pObj2)
    Debug.Print "  Handler_Release(pObj2) -> " & r
    Debug.Print "(h2 ___) m_inUse = " & m_inUse & " (___ 0)"
    Debug.Print "h2.RefCount (___) = " & h2.RefCount & " (___ 0)"

    ' --- 6. ________________ no-op ________ ---
    Debug.Print "--- Post-release safety check ---"
    r = HandlerAddRefInternal(pObj1)
    Debug.Print "____ pObj1 __ AddRef ___ = " & r & " (___ 0)"
    r = HandlerReleaseInternal(pObj2)
    Debug.Print "____ pObj2 __ Release ___ = " & r & " (___ 0)"

    ' --- 7. ____ ---
    Set h1 = Nothing
    Set h2 = Nothing

    Debug.Print "--- ____ ---"
    Debug.Print "m_inUse = " & m_inUse & " (___ 0)"
    Debug.Print "m_freeHead = " & m_freeHead

    Thunks_Shutdown

    Debug.Print "Test_TwoHandlers_Mock __"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_Environment_Real (_9.3a _____)
'
'   Wv2Environment ___ _________.
'   _ WebView2Loader.dll _____ Environment __ -> ______ ->
'   AddRef -> Class_Terminate __ Release _____________.
'
'   ____:
'     1. Wv2Environment.Init _ ComCallbackHandler _ 1 _ Acquire __, API .
'     2. WebView2 __ Invoke ___ ComCallbackHandler.Handler_Invoke .
'        Wv2Environment.OnEnvironmentCompleted _______ (_________)
'     3. errorCode = 0, pEnvironment != 0 _____ IsReady .
'     4. m_handler _ Nothing _______ (____________)
'     5. m_pEnvironment __ AddRef _______ (= 1 ________)
'     6. End Sub _ Class_Terminate ___, ComRelease .
'
'   _______ 5 _. WebView2 ___________ (______)
'   ___________.
' ============================================================
Public Sub Test_Environment_Real()

    Debug.Print String(60, "=")
    Debug.Print "Test_Environment_Real __"
    Debug.Print String(60, "=")

    Dim env As Wv2Environment
    Set env = New Wv2Environment

    Debug.Print "env.Init _____ ..."
    env.Init "C:\Temp\VBA_WebView2"

    If env.IsFailed Then
        Debug.Print "[NG] env.Init ______ (LastError = &H" & Hex(env.LastError) & ")"
        GoTo Cleanup
    End If

    Debug.Print "Init ___ State = " & env.state & _
                " (1=Es_Waiting / 2=Es_Ready ________ / 3=Es_Failed)"
    Debug.Print "_________ (__ 5 _, _______________) ..."

    ' ______ IsReady / IsFailed _______.
    ' (Wv2Environment _ Public Enum .
    '  ________, Bool ____________)
    Dim t As Single: t = Timer
    Do While (Not env.IsReady) And (Not env.IsFailed) And ((Timer - t) < 5#)
        DoEvents
    Loop

    If env.IsReady Then
        Debug.Print "[OK] env.IsReady = True"
        Debug.Print "  m_pEnvironment   = " & env.EnvironmentPtr
        Debug.Print "  ____          = " & Format$(Timer - t, "0.000") & " _"
    ElseIf env.IsFailed Then
        Debug.Print "[NG] env.IsFailed = True (LastError = &H" & Hex(env.LastError) & ")"
    Else
        Debug.Print "[NG] ______ (state = " & env.state & ")"
    End If

Cleanup:
    Debug.Print "--- env ___ (Class_Terminate _ ComRelease _____) ---"
    Set env = Nothing
    Debug.Print "Set env = Nothing __"

    ' ____ Thunks_Shutdown .
    Debug.Print "Thunks_Shutdown __"
    Thunks_Shutdown

    Debug.Print "Test_Environment_Real __"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_ResetSurvival_Manual (_9.4a _____)
'
'   _________ (_ R4) ________.
'   _______________________, .
'   ____ VBE _______.
'
'   ____ (_________):
'     1. ________ Test_ResetSurvival_Manual ___ (____)
'     2. UserForm1.Show vbModeless -> UserForm1.StartWebView2_Full .
'        Bing __ + ________ (NavigationStarting __) .
'     3. Bing _______, _________ 1?2 .
'        NavigationStarting _______________ Debug.Print .
'     4. **_____ VBE ______ (____) ___**
'        -> m_pRegionBase ______________ 0 .
'        -> ___ VirtualAlloc ___ OS .
'        -> ___________ (Thunks_Shutdown ________)
'     5. Excel .
'     6. _______ (____________________) ,
'        WebView2 _ ComCallbackHandler ________ Excel .
'        (____________, .
'         _____ return ________________. .
'         _____________________)
'     7. ________ Test_Navigate_Real .
'        (Thunks_Init __m_pRegionBase = 0 __________,
'         ___________________)
'
'   _______:
'     ? ______ Excel .
'     ? .
'     ? _9.5 _____, .
'       ____ VirtualFree __ (_______)
'
'   ______________ Test_SentinelStatus _______.
' ============================================================
Public Sub Test_ResetSurvival_Manual()
    Debug.Print String(60, "=")
    Debug.Print "Test_ResetSurvival_Manual (_________)"
    Debug.Print String(60, "=")
    Debug.Print "______________:"
    Debug.Print ""
    Debug.Print "1. __ Test_Navigate_Real _____ Bing ___"
    Debug.Print "2. Bing _____ 1?2 _______ NavigationStarting _____"
    Debug.Print "   Debug.Print ________"
    Debug.Print "3. _____ VBE ______ (____) ___"
    Debug.Print "4. Excel __________"
    Debug.Print "5. _______, _______________"
    Debug.Print "   Excel ______________"
    Debug.Print "6. _______ Test_Navigate_Real ____________"
    Debug.Print ""
    Debug.Print "____:"
    Debug.Print "  ? ______ Excel _________ (________)"
    Debug.Print "  ? ______________"
    Debug.Print "  ? _9.5 _____, _____________"
    Debug.Print "    ________ VirtualFree __ (_______)"
    Debug.Print ""
    Debug.Print "______________ Test_SentinelStatus ___:"
    Debug.Print "  - ______________________________"
    Debug.Print "  - _ Init _______________________"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Sentinel_LoadPrevRegion (_9.5 _____)
'
'   ____ SENTINEL_ENV_NAME _____ Thunks_Init .
'   _____________. _____ (= ______, .
'   ____) ___ 0 ___.
'
'   _ ___________________________.
'      _______________ Shutdown ____________.
'      (______________, _ Excel .
'       _____)
'
'   ___: ____________ (____ 0)
' ============================================================
Private Function Sentinel_LoadPrevRegion() As LongPtr
    Sentinel_LoadPrevRegion = NullPtr

    Dim buff As String
    buff = String$(SENTINEL_BUFFER_SIZE, vbNullChar)

    Dim n As Long
    n = GetEnvironmentVariableW(StrPtr(SENTINEL_ENV_NAME), StrPtr(buff), SENTINEL_BUFFER_SIZE)

    ' n = 0 : ________ or .
    If n = 0 Then Exit Function

    ' n >= SENTINEL_BUFFER_SIZE : ______ (= ___, __)
    If n >= SENTINEL_BUFFER_SIZE Then Exit Function

    Dim s As String
    s = Left$(buff, n)

    ' _______ (___, _________) _____ 0 .
    On Error Resume Next
    Sentinel_LoadPrevRegion = CLngPtr(s)
    On Error GoTo 0
End Function


' ============================================================
' Sentinel_StorePrevRegion (_9.5 _____)
'
'   ______________________.
'   ________ (MouseScroll.bas __).
'
'   _________:
'     Thunks_Init _ VirtualAlloc ____. _____, Init .
'     __________ Init ____________.
' ============================================================
Private Sub Sentinel_StorePrevRegion(ByVal addr As LongPtr)
    Dim s As String
    s = CStr(addr)
    SetEnvironmentVariableW StrPtr(SENTINEL_ENV_NAME), StrPtr(s)
End Sub


' ============================================================
' Sentinel_ClearPrevRegion (_9.5 _____)
'
'   __________.
'   Win32 __: SetEnvironmentVariableW(lpName, NULL) ________.
'   VBA ________ ByVal 0 ______ (LongPtr _ ByVal _ 0 .
'   _______ NULL ___________).
'
'   _________:
'     1. Thunks_Shutdown ___ (______)
'     2. Sentinel_RecoverIfNeeded ___ (_________ Init ___)
' ============================================================
Private Sub Sentinel_ClearPrevRegion()
    SetEnvironmentVariableW StrPtr(SENTINEL_ENV_NAME), 0
End Sub


' ============================================================
' Sentinel_RecoverIfNeeded (_9.5 _____)
'
'   Thunks_Init ________. _________________.
'
'   ____ (_ _, _____ + VirtualFree _____):
'     1. __________ 0 .
'        - __ WebView2 _________________ S_OK .
'        - VirtualFree .
'        - _______/.
'          ____ On Error .
'     2. VirtualFree .
'        - ____ GetLastError .
'     3. .
'        - _____ Thunks_Init _______________, .
'          VirtualAlloc .
'
'   ________________________________:
'     - ___________ (_->.
'       _________)
'     - Excel _ VBA ________________ (________)
'     - __________ = .
'
'   ____ (___________________):
'     - "VirtualQuery result ...": OS .
'         State = MEM_COMMIT (4096)  -> _____ (__)
'         State = MEM_FREE   (65536) -> .
'         State = MEM_RESERVE(8192)  -> .
'     - "header LongPtr before clear: X" : .
'         X = 1   -> _____, _________ (__)
'         X = 0   -> .
'         ______ -> _______/.
'     - "VirtualFree ... LastError: Y" : VirtualFree .
'         Y = 0   -> ________ (succeeded)
'         Y = 87  -> ERROR_INVALID_PARAMETER (_9.5 _________,
'                  __________ & ___________)
'         Y = 170 -> ERROR_BUSY (________________)
'         Y = 487 -> ERROR_INVALID_ADDRESS (______, _______)
' ============================================================
Private Sub Sentinel_RecoverIfNeeded()
    Dim prevBase As LongPtr
    prevBase = Sentinel_LoadPrevRegion()
    If prevBase = 0 Then Exit Sub

    Debug.Print "Sentinel: detected previous region = " & prevBase & _
                " (decimal), recovering ..."

    ' --- __0 + ____ (_9.5 ____ / _9.10a _________) ---
    '   VirtualQuery _ OS ____________, .
    '   ____. _______________:
    '     State = MEM_COMMIT (4096)  -> _____ (VirtualFree ______)
    '     State = MEM_FREE   (65536) -> ______ (_____ Free __)
    '     State = MEM_RESERVE(8192)  -> Reserve ____ (____)
    '   AllocationBase _ prevBase _________ (= ________).
    '
    '   * _9.10a ______ (____ 10 ___) *
    '     _____VirtualQuery _ State ___________, .
    '     ______ ReadLongPtr / MemLongPtr / VirtualFree .
    '     __. State=MEM_FREE __________ ReadLongPtr / MemLongPtr .
    '     ___ SAFEARRAY _______________ AV (______) .
    '     ____. __ AV _ CPU _ SEH _____ VBA _ On Error Resume Next
    '     ________ (= Excel _______, ____ 10).
    '     ___ VirtualQuery ___________________,
    '     ReadLongPtr / MemLongPtr / VirtualFree _____ (____ 55).
    Dim mbi As MEMORY_BASIC_INFORMATION
    Dim qSize As LongPtr
    Dim regionAlive As Boolean
    regionAlive = False

    qSize = VirtualQuery(prevBase, mbi, LenB(mbi))
    If qSize = 0 Then
        Debug.Print "Sentinel: VirtualQuery FAILED, LastError = " & GetLastError()
    Else
        Debug.Print "Sentinel: VirtualQuery result (" & qSize & " bytes copied):"
        Debug.Print "  BaseAddress       = " & mbi.BaseAddress
        Debug.Print "  AllocationBase    = " & mbi.AllocationBase & _
                    IIf(mbi.AllocationBase = prevBase, " (= prevBase, OK)", " (_ prevBase, ____)")
        Debug.Print "  AllocationProtect = &H" & Hex$(mbi.AllocationProtect) & _
                    " (" & FormatMemProtect(mbi.AllocationProtect) & ")"
        Debug.Print "  RegionSize        = " & mbi.RegionSize
        Debug.Print "  State             = " & mbi.state & " (" & FormatMemState(mbi.state) & ")"
        Debug.Print "  Protect           = &H" & Hex$(mbi.Protect) & _
                    " (" & FormatMemProtect(mbi.Protect) & ")"
        Debug.Print "  Type              = &H" & Hex$(mbi.Type_)

        ' ____ (__): 3 ____________________.
        '   (1) State = MEM_COMMIT       : ____ commit __ (__/_______)
        '   (2) AllocationBase = prevBase : .
        '                                   .
        '   (3) Protect _______    : PAGE_GUARD / PAGE_NOACCESS ___.
        '       ______ PAGE_EXECUTE_READWRITE _______, .
        '       PAGE_READWRITE _____. PAGE_GUARD .
        '       _____ STATUS_GUARD_PAGE_VIOLATION _________.
        regionAlive = (mbi.state = MEM_COMMIT) _
                      And (mbi.AllocationBase = prevBase) _
                      And IsProtectReadWritable(mbi.Protect)
    End If

    If Not regionAlive Then
        ' ________ (= ______ / ____ / ____ / ____).
        ' ReadLongPtr / MemLongPtr / VirtualFree ______, .
        ' _________. VirtualFree _______: .
        ' MEM_RELEASE _ ERROR_INVALID_PARAMETER ____________,
        ' _____________________________.
        Debug.Print "Sentinel: ____________ -> ________ (____ 10 __)"
        Sentinel_ClearPrevRegion
        Exit Sub
    End If

    ' --- _________________ VirtualQuery _____ ---

    ' __1: _____________ (_________ AV _____)
    '   _: MemLongPtr _ Property Let __ (______). .
    '       ReadLongPtr _____.
    Dim headerByte As LongPtr
    headerByte = ReadLongPtr(prevBase)
    Debug.Print "Sentinel: header LongPtr before clear = " & headerByte & _
                " (_____ = " & (headerByte And &HFF&) & ")"

    ' 1. ______ 0 ___ (____)
    MemLongPtr(prevBase) = NullPtr

    ' 2. VirtualFree _______ (GetLastError __________)
    '    Win32 __: VirtualFree ____ 0 ___, ___ GetLastError _____.
    Dim freeResult As Long
    Dim lastErr As Long
    freeResult = VirtualFree(prevBase, 0, MEM_RELEASE)
    lastErr = GetLastError()
    If freeResult <> 0 Then
        Debug.Print "Sentinel: VirtualFree succeeded for " & prevBase
    Else
        Debug.Print "Sentinel: VirtualFree returned 0, LastError = " & lastErr & _
                    " (" & VirtualFreeErrorName(lastErr) & ")"
    End If

    ' 3. __________ (_____________)
    Sentinel_ClearPrevRegion
End Sub

' ============================================================
' IsProtectReadWritable (_9.10a __)
' ============================================================
'   MEMORY_BASIC_INFORMATION.Protect __VBA .
'   _____. Sentinel_RecoverIfNeeded ________ (____ 55).
'
'   __________: PAGE_READWRITE (&H4) / PAGE_EXECUTE_READWRITE (&H40).
'   ___ PAGE_GUARD (&H100) ________________________,
'   ______________.
Private Function IsProtectReadWritable(ByVal prot As Long) As Boolean
    ' PAGE_GUARD .
    If (prot And &H100&) <> 0 Then
        IsProtectReadWritable = False
        Exit Function
    End If
    ' .
    Dim base As Long
    base = prot And Not &H100& And Not &H200& And Not &H400&
    Select Case base
        Case PAGE_READWRITE, PAGE_EXECUTE_READWRITE
            IsProtectReadWritable = True
        Case Else
            IsProtectReadWritable = False
    End Select
End Function


' ============================================================
' VirtualFreeErrorName (_9.5 _____, ___)
'
'   GetLastError ___ Win32 _________. .
'   _____, _____ "UNKNOWN" ___.
' ============================================================
Private Function VirtualFreeErrorName(ByVal errCode As Long) As String
    Select Case errCode
        Case 0: VirtualFreeErrorName = "ERROR_SUCCESS (= _______?)"
        Case 5: VirtualFreeErrorName = "ERROR_ACCESS_DENIED"
        Case 6: VirtualFreeErrorName = "ERROR_INVALID_HANDLE"
        Case 8: VirtualFreeErrorName = "ERROR_NOT_ENOUGH_MEMORY"
        Case 87: VirtualFreeErrorName = "ERROR_INVALID_PARAMETER"
        Case 170: VirtualFreeErrorName = "ERROR_BUSY (________)"
        Case 487: VirtualFreeErrorName = "ERROR_INVALID_ADDRESS"
        Case Else: VirtualFreeErrorName = "UNKNOWN"
    End Select
End Function


' ============================================================
' FormatMemState (_9.5 _____, ___)
'
'   MEMORY_BASIC_INFORMATION.State _________.
' ============================================================
Private Function FormatMemState(ByVal state As Long) As String
    Select Case state
        Case &H1000: FormatMemState = "MEM_COMMIT (_____)"
        Case &H2000: FormatMemState = "MEM_RESERVE (Reserve __, Commit _)"
        Case &H10000: FormatMemState = "MEM_FREE (______)"
        Case Else: FormatMemState = "UNKNOWN(&H" & Hex$(state) & ")"
    End Select
End Function


' ============================================================
' FormatMemProtect (_9.5 _____, ___)
'
'   MEMORY_BASIC_INFORMATION.Protect / AllocationProtect _________.
'   PAGE_GUARD _________ OR _________________.
' ============================================================
Private Function FormatMemProtect(ByVal prot As Long) As String
    If prot = 0 Then
        FormatMemProtect = "NONE (= ________ Protect=0)"
        Exit Function
    End If

    Dim s As String

    Select Case prot And &HFF&
        Case &H1: s = "PAGE_NOACCESS"
        Case &H2: s = "PAGE_READONLY"
        Case &H4: s = "PAGE_READWRITE"
        Case &H8: s = "PAGE_WRITECOPY"
        Case &H10: s = "PAGE_EXECUTE"
        Case &H20: s = "PAGE_EXECUTE_READ"
        Case &H40: s = "PAGE_EXECUTE_READWRITE"
        Case &H80: s = "PAGE_EXECUTE_WRITECOPY"
        Case Else: s = "UNKNOWN(&H" & Hex$(prot And &HFF&) & ")"
    End Select

    If (prot And &H100&) <> 0 Then s = s & "|PAGE_GUARD"
    If (prot And &H200&) <> 0 Then s = s & "|PAGE_NOCACHE"
    If (prot And &H400&) <> 0 Then s = s & "|PAGE_WRITECOMBINE"

    FormatMemProtect = s
End Function


' ============================================================
' Test_VirtualAllocFree_Roundtrip (_9.5 _____)
'
'   VBA ________________ VirtualAlloc -> VirtualFree .
'   ______ 5 ____________.
'   9.5c ________________ ERROR_INVALID_PARAMETER .
'   ____________________.
'
'   ___:
'     1. Excel + VBE ______________ (__)
'     2. ________ Test_VirtualAllocFree_Roundtrip .
'     3. _?_ .
'     4. (_____) ___ Test_AllEvents_Real ___ Thunks_Init ____,
'        m_pRegionBase _________ Test_VirtualAllocFree_Eps .
'        _______ _ .
'
'   _________ -> .
'                       (VBA _______________________)
'   ______   -> .
' ============================================================
Public Sub Test_VirtualAllocFree_Roundtrip()
    Debug.Print String(60, "=")
    Debug.Print "Test_VirtualAllocFree_Roundtrip (___________)"
    Debug.Print String(60, "=")

    Dim p As LongPtr, p2 As LongPtr
    Dim r As Long
    Dim errCode As Long

    ' --- ____ _: alloc -> _ free ---
    Debug.Print "--- ____ _: alloc -> _ free ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  VirtualAlloc -> " & p
    If p = 0 Then
        Debug.Print "  [SKIP] alloc ___, ____________"
        Exit Sub
    End If
    r = VirtualFree(p, 0, MEM_RELEASE)
    errCode = GetLastError()
    Debug.Print "  VirtualFree returned " & r & ", LastError = " & errCode & _
                IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")

    ' --- ____ _: alloc -> ______ -> free ---
    Debug.Print "--- ____ _: alloc -> ______ -> free ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  VirtualAlloc -> " & p
    If p <> 0 Then
        MemLongPtr(p) = CLngPtr(1)               ' .
        MemLongPtr(p + PtrSize) = CLngPtr(&H42&) ' ____ (Long ______)
        Debug.Print "  ___________ (header LongPtr = " & ReadLongPtr(p) & ")"
        r = VirtualFree(p, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
    End If

    ' --- ____ _: alloc -> __________ -> free ---
    Debug.Print "--- ____ _: alloc -> __________ -> free ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  VirtualAlloc -> " & p
    If p <> 0 Then
        Dim s As String
        s = CStr(p)
        Dim pRestored As LongPtr
        pRestored = CLngPtr(s)
        Debug.Print "  CStr/CLngPtr round-trip: " & p & " -> '" & s & "' -> " & pRestored & _
                    IIf(p = pRestored, " [__]", " [___!]")
        r = VirtualFree(pRestored, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree(pRestored) returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
    End If

    ' --- ____ _: alloc(A) -> alloc(B) -> free(A) -> free(B) ---
    '   ___ Sentinel_RecoverIfNeeded _______. .
    '   ____.
    Debug.Print "--- ____ _: 2 ____ -> ______ free -> _____ free ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  __ A alloc -> " & p
    p2 = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  __ B alloc -> " & p2
    If p <> 0 And p2 <> 0 Then
        MemLongPtr(p) = CLngPtr(1)             ' __ A .
        r = VirtualFree(p, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree(A) returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
        r = VirtualFree(p2, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree(B) returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
    End If

    ' --- ____ _: ___ Thunks_Init ______ VirtualFree ---
    '   __ m_pRegionBase ________ (= Thunks_Init __) __,
    '   __________ VirtualQuery -> VirtualFree ____.
    '   __: ________ Thunks_Shutdown .
    '         m_pRegionBase _ 0 _____ (= _______).
    '         ______ Thunks_Shutdown ______________,
    '         m_pRegionBase = 0 _________.
    Debug.Print "--- ____ _: ___ Thunks_Init _____ ---"
    If m_pRegionBase <> 0 Then
        Debug.Print "  [SKIP] m_pRegionBase = " & m_pRegionBase & " (= __ Init __)"
        Debug.Print "         _______ Shutdown _ (= m_pRegionBase = 0) __________"
    Else
        Debug.Print "  Thunks_Init ___ ..."
        If Not Thunks_Init() Then
            Debug.Print "  [SKIP] Thunks_Init __"
        Else
            Dim baseBak As LongPtr
            baseBak = m_pRegionBase
            Debug.Print "  m_pRegionBase = " & baseBak
            ' Thunks_Shutdown _______ VirtualFree ____.
            ' ___________, m_pRegionBase .
            ' _____ Sentinel_ClearPrevRegion ___.
            r = VirtualFree(baseBak, 0, MEM_RELEASE)
            errCode = GetLastError()
            Debug.Print "  VirtualFree(__) returned " & r & ", LastError = " & errCode & _
                        IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
            ' .
            m_pRegionBase = 0
            m_freeHead = -1
            m_inUse = 0
            Erase m_freeNext
            Sentinel_ClearPrevRegion
            Debug.Print "  _____ (m_pRegionBase = 0, _________)"
        End If
    End If

    Debug.Print String(60, "=")
    Debug.Print "Test_VirtualAllocFree_Roundtrip __"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_SentinelStatus (_9.5 _____, _____)
'
'   ___________ m_pRegionBase _ Debug.Print __________.
'   _______________Test_SentinelStatus_______.
'
'   ______:
'     - ______: ______, m_pRegionBase = 0
'     - StartWebView2 _: ____ = m_pRegionBase = .
'     - _____: ______, m_pRegionBase = 0
'     - ______: ____ = _______, m_pRegionBase = 0
'         (VBA _________________________)
'     - _______ Init _: ____ = m_pRegionBase = .
'         (______ VirtualFree __)
' ============================================================
Public Sub Test_SentinelStatus()
    Dim prev As LongPtr
    prev = Sentinel_LoadPrevRegion()

    Debug.Print String(60, "=")
    Debug.Print "Test_SentinelStatus"
    Debug.Print "  Env var name : " & SENTINEL_ENV_NAME
    If prev = 0 Then
        Debug.Print "  Env var value: (empty or unset) -> no leftover region"
    Else
        Debug.Print "  Env var value: " & prev & " (decimal pointer)"
        If prev = m_pRegionBase Then
            Debug.Print "  -> ___________ (____)"
        ElseIf m_pRegionBase = 0 Then
            Debug.Print "  -> __ Thunks_Init ________ (_____)"
        Else
            Debug.Print "  -> _____ m_pRegionBase ________ (__)"
        End If
    End If
    Debug.Print "  m_pRegionBase: " & m_pRegionBase
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Thunks_Init
'   ___ VirtualAlloc ____, ____________,
'   ______________, vtableObj _____,
'   ____________.
'
'   _9.2______ (_9.3a ______):
'     - Handler_QI / AddRef / Release ______ 1 .
'     - _ 512 _____ vtableObj .
'         pVTable      = pFunctions
'         Functions(0) = m_pHandler_QI
'         Functions(1) = m_pHandler_AddRef
'         Functions(2) = m_pHandler_Release
'         Functions(3) = pSlot (= _________)
'
'   _9.5 ________:
'     - ___ Sentinel_RecoverIfNeeded .
'         (___________________ VirtualFree ____)
'     - VirtualAlloc _____ Sentinel_StorePrevRegion .
'         (_______________, ____________)
'
'   ___: __ = True, __ = False
' ============================================================
Public Function Thunks_Init() As Boolean

    If m_pRegionBase <> 0 Then
        Thunks_Init = True
        Exit Function
    End If

    ' --- _______ (_9.5 ____) ---
    '   _________________ Shutdown _________,
    '   __________________. .
    '   ________.
    Sentinel_RecoverIfNeeded

    ' Force-Compile EntryPoint (___________)
    EntryPoint

    ' EntryPoint .
    Dim pStubSrc As LongPtr
    pStubSrc = VBA.Int(AddressOf EntryPoint)
    If pStubSrc = 0 Then Exit Function

    ' Handler_QI / AddRef / Release __________ 1 .
    m_pHandler_QI = GetAddr(AddressOf Handler_QueryInterface)
    m_pHandler_AddRef = GetAddr(AddressOf Handler_AddRef)
    m_pHandler_Release = GetAddr(AddressOf Handler_Release)
    If m_pHandler_QI = 0 Or m_pHandler_AddRef = 0 Or m_pHandler_Release = 0 Then
        Exit Function
    End If

    ' .
    m_pRegionBase = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If m_pRegionBase = 0 Then Exit Function

    ' --- _______ (_9.5 ____) ---
    '   _________________.
    '   _____, ___ Init .
    '   __ Init _________.
    Sentinel_StorePrevRegion m_pRegionBase

    ' _________ (VirtualAlloc ________________)
    Dim k As Long
    For k = 0 To HEADER_SIZE - 1 Step PtrSize
        MemLongPtr(m_pRegionBase + k) = NullPtr
    Next k

    ' _________ (_____ 1 byte = 1)
    MemLongPtr(m_pRegionBase) = CLngPtr(1)

    ' ________ (______ + vtableObj ___)
    Dim i As Long, pSlot As LongPtr, pVTableObj As LongPtr, pFunctions As LongPtr
    For i = 0 To SLOT_COUNT - 1
        pSlot = SlotAddrAt(i)

        ' --- ______ ---
        #If Win64 Then
        For k = 0 To 95 Step 8
            MemLongPtr(pSlot + k) = ReadLongPtr(pStubSrc + k)
        Next k
        #Else
        ' 32-bit VBE7 EntryPoint stub is 33 bytes (+0..+32); copy 36 bytes (9 DWORDs)
        ' and patch +30..+31 from C2 00 00 (RET 0) to C2 0C 00 (RET 12) so Break mode
        ' cleans the 3 __stdcall DWORDs (this, arg1, arg2) pushed by WebView2!
        For k = 0 To 35 Step 4
            MemLongPtr(pSlot + k) = ReadLongPtr(pStubSrc + k)
        Next k
        ' At +28 bytes: 33 C0 C2 00 (0x00C2C033) -> 33 C0 C2 0C (0x0CC2C033)
        ' At +32 bytes: 00 ... (high byte of imm16 = 0x00, already 0 from stub)
        MemLongPtr(pSlot + 28) = &HCC2C033
        #End If

        ' --- vtableObj ___ ---
        pVTableObj = pSlot + VTABLE_OBJ_OFFSET
        pFunctions = pVTableObj + PtrSize        ' pVTable .

        MemLongPtr(pVTableObj) = pFunctions
        MemLongPtr(pFunctions + 0 * PtrSize) = m_pHandler_QI
        MemLongPtr(pFunctions + 1 * PtrSize) = m_pHandler_AddRef
        MemLongPtr(pFunctions + 2 * PtrSize) = m_pHandler_Release
        MemLongPtr(pFunctions + 3 * PtrSize) = pSlot
    Next i

    ' .
    ReDim m_freeNext(0 To SLOT_COUNT - 1)
    For i = 0 To SLOT_COUNT - 2
        m_freeNext(i) = i + 1
    Next i
    m_freeNext(SLOT_COUNT - 1) = -1
    m_freeHead = 0
    m_inUse = 0

    ' Handler .
    For i = 0 To SLOT_COUNT - 1
        Set m_handlers(i) = Nothing
    Next i

    ' --- IID _______ (_9.7a _____) ---
    '   IID_IUnknown _ HandlerKind ___ IID __________.
    InitIIDTable

    Thunks_Init = True
End Function


' ============================================================
' Thunks_AcquireSlot
'   _______________ 1 ____, .
'   _______pSelfObj _ Me _____ pTargetFunc .
'   ___________, +55 ______________.
'
'   _9.3a ____ AcquireHandlerFor _____________,
'   Test_TwoHandlers_Mock __________________.
'
'   ___: __________, ___ 0
' ============================================================
Public Function Thunks_AcquireSlot( _
    ByVal handler As ComCallbackHandler, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr

    If m_pRegionBase = 0 Then Exit Function
    If m_freeHead < 0 Then Exit Function     ' .
    If handler Is Nothing Then Exit Function

    ' ________ 1 .
    Dim idx As Long
    idx = m_freeHead
    m_freeHead = m_freeNext(idx)
    m_freeNext(idx) = -1                      ' -1 = .

    Dim pSlot As LongPtr
    pSlot = SlotAddrAt(idx)

    ' ________ (offset +96 __ 74 bytes)
    WriteThunkMachineCode pSlot + THUNK_OFFSET, pSelfObj, pTargetFunc, m_pRegionBase

    ' ________ +55 .
    MemLongPtr(pSlot + LATE_BIND_OFFSET) = pSlot + THUNK_OFFSET

    ' Handler _______ idx .
    Set m_handlers(idx) = handler

    m_inUse = m_inUse + 1
    Thunks_AcquireSlot = pSlot
End Function


' ============================================================
' Thunks_ReleaseSlot
'   _____________________.
'   m_handlers(idx) ______.
' ============================================================
Public Sub Thunks_ReleaseSlot(ByVal pSlot As LongPtr)
    If m_pRegionBase = 0 Then Exit Sub
    If pSlot = 0 Then Exit Sub

    Dim idx As Long
    idx = SlotIndexFromAddr(pSlot)
    If idx < 0 Then Exit Sub
    If m_freeNext(idx) <> -1 Then Exit Sub           ' ____ = .

    ' Handler .
    Set m_handlers(idx) = Nothing

    ' .
    m_freeNext(idx) = m_freeHead
    m_freeHead = idx

    m_inUse = m_inUse - 1
End Sub


' ============================================================
' SlotIndexFromAddr (_____)
'   ________ (= __________) .
'   ____ index _____.
' ============================================================
Private Function SlotIndexFromAddr(ByVal pSlot As LongPtr) As Long
    SlotIndexFromAddr = -1
    If m_pRegionBase = 0 Then Exit Function

    Dim offset As LongPtr
    offset = pSlot - (m_pRegionBase + HEADER_SIZE)
    If offset < 0 Then Exit Function
    If (offset Mod SLOT_SIZE) <> 0 Then Exit Function

    Dim idx As Long
    idx = CLng(offset \ SLOT_SIZE)
    If idx < 0 Or idx >= SLOT_COUNT Then Exit Function

    SlotIndexFromAddr = idx
End Function


' ============================================================
' SlotIndexFromVTableObjAddr (_____)
'   vtableObj _____ (= WebView2 _ AddRef/Release _ this _________,
'                     = pSlot + VTABLE_OBJ_OFFSET) __, ____ index _____.
' ============================================================
Private Function SlotIndexFromVTableObjAddr(ByVal pVTableObj As LongPtr) As Long
    SlotIndexFromVTableObjAddr = -1
    If m_pRegionBase = 0 Then Exit Function

    Dim offset As LongPtr
    offset = pVTableObj - (m_pRegionBase + HEADER_SIZE + VTABLE_OBJ_OFFSET)
    If offset < 0 Then Exit Function
    If (offset Mod SLOT_SIZE) <> 0 Then Exit Function

    Dim idx As Long
    idx = CLng(offset \ SLOT_SIZE)
    If idx < 0 Or idx >= SLOT_COUNT Then Exit Function

    SlotIndexFromVTableObjAddr = idx
End Function


' ============================================================
' Thunks_Shutdown
'   ______ 0 ______ VirtualFree ________.
'
'   _9.5 ________:
'     - ___ Sentinel_ClearPrevRegion ______________.
'       ________ Init .
'       VirtualFree _____________. VirtualFree .
'       0 _______, ___________.
' ============================================================
Public Sub Thunks_Shutdown()
    If m_pRegionBase = 0 Then Exit Sub

    ' .
    MemLongPtr(m_pRegionBase) = NullPtr

    ' _ Handler _ owner _________ (______)
    Dim i As Long
    For i = 0 To SLOT_COUNT - 1
        If Not (m_handlers(i) Is Nothing) Then
            m_handlers(i).ClearOwner
            Set m_handlers(i) = Nothing
        End If
    Next i

    ' _____ (_9.5 _________: ____ LastError ___)
    Dim shutFreeResult As Long
    Dim shutFreeErr As Long
    shutFreeResult = VirtualFree(m_pRegionBase, 0, MEM_RELEASE)
    shutFreeErr = GetLastError()
    Debug.Print "Thunks_Shutdown: VirtualFree(" & m_pRegionBase & ") returned " & _
                shutFreeResult & ", LastError = " & shutFreeErr & _
                IIf(shutFreeResult <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(shutFreeErr) & ")")

    m_pRegionBase = 0
    m_freeHead = -1
    m_inUse = 0
    Erase m_freeNext

    ' --- _______ (_9.5 ____) ---
    '   _______________. _________ Init .
    '   ____________ VirtualFree .
    '   (______ Debug.Print __ "VirtualFree returned 0" .
    '    _______) _____.
    Sentinel_ClearPrevRegion
End Sub


' ============================================================
' __________ -> .
' ============================================================
Private Function SlotAddrAt(ByVal idx As Long) As LongPtr
    SlotAddrAt = m_pRegionBase + HEADER_SIZE + CLngPtr(idx) * SLOT_SIZE
End Function


' ============================================================
' ______________________ (x64: 74 bytes / x86: 46 bytes)
' ============================================================
Private Sub WriteThunkMachineCode( _
    ByVal addr As LongPtr, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr, _
    ByVal pRegionBase As LongPtr)

    Dim b(0 To THUNK_BUF_SIZE - 1) As Byte, i As Long
    i = 0

#If Win64 Then
    ' --- _________ (18 bytes) ---
    ' mov rax, imm64   (= pAliveFlag = pRegionBase + 0)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pRegionBase: i = i + 8
    ' cmp byte ptr [rax], 1
    b(i) = &H80: b(i + 1) = &H38: b(i + 2) = &H1: i = i + 3
    ' je +3 (= alive ___, xor + ret _ 3 bytes _____)
    b(i) = &H74: b(i + 1) = &H3: i = i + 2
    ' xor eax, eax
    b(i) = &H33: b(i + 1) = &HC0: i = i + 2
    ' ret
    b(i) = &HC3: i = i + 1

    ' --- __, __?_______ 56 bytes _____ ---

    ' sub rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4
    ' mov r9, r8
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    ' mov rcx, imm64   (pSelfObj)
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pSelfObj: i = i + 8
    ' lea rax, [rsp+0x28]
    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    ' mov [rsp+0x20], rax
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5
    ' mov rax, imm64   (pTargetFunc)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pTargetFunc: i = i + 8
    ' call rax
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    ' mov eax, [rsp+0x28]
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4
    ' add rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4
    ' ret
    b(i) = &HC3: i = i + 1
    ' int3 / int3 (_____)
    b(i) = &HCC: i = i + 1
    b(i) = &HCC: i = i + 1
#Else
    ' --- 32-bit x86 __stdcall Thunk (46 bytes) ---
    ' On entry from JMP EAX in VBE7 EntryPoint stub:
    '   [esp+0]  = return address to WebView2 caller
    '   [esp+4]  = this (vtableObj*)
    '   [esp+8]  = arg1 (errorCode or sender)
    '   [esp+12] = arg2 (result ptr or eventArgs)
    '
    ' Alive flag check (15 bytes):
    ' mov eax, imm32 (pRegionBase)
    b(i) = &HB8: i = i + 1
    MemLongPtr(VarPtr(b(i))) = pRegionBase: i = i + 4
    ' cmp byte ptr [eax], 1
    b(i) = &H80: b(i + 1) = &H38: b(i + 2) = &H1: i = i + 3
    ' je +5 (skip xor eax,eax + ret 0x000C)
    b(i) = &H74: b(i + 1) = &H5: i = i + 2
    ' xor eax, eax (S_OK = 0)
    b(i) = &H33: b(i + 1) = &HC0: i = i + 2
    ' ret 0x000C (clean 3 DWORDs)
    b(i) = &HC2: b(i + 1) = &HC: b(i + 2) = &H0: i = i + 3

    ' Allocate 4-byte retval slot on stack (push 0 -> esp -= 4)
    b(i) = &H6A: b(i + 1) = &H0: i = i + 2
    ' mov eax, esp (&retval); now [esp+8]=this, [esp+12]=arg1, [esp+16]=arg2
    b(i) = &H89: b(i + 1) = &HE0: i = i + 2
    ' push eax (&retval, 5th param of VBA Function Handler_Invoke)
    b(i) = &H50: i = i + 1
    ' push dword ptr [esp+20] (arg2; esp shifted -4 after push eax)
    b(i) = &HFF: b(i + 1) = &H74: b(i + 2) = &H24: b(i + 3) = &H14: i = i + 4
    ' push dword ptr [esp+20] (arg1; esp shifted -8)
    b(i) = &HFF: b(i + 1) = &H74: b(i + 2) = &H24: b(i + 3) = &H14: i = i + 4
    ' push dword ptr [esp+20] (this; esp shifted -12)
    b(i) = &HFF: b(i + 1) = &H74: b(i + 2) = &H24: b(i + 3) = &H14: i = i + 4
    ' push imm32 (pSelfObj = ObjPtr(ComCallbackHandler), 1st param Me)
    b(i) = &H68: i = i + 1
    MemLongPtr(VarPtr(b(i))) = pSelfObj: i = i + 4
    ' mov eax, imm32 (pTargetFunc = ComCallbackHandler.Handler_Invoke)
    b(i) = &HB8: i = i + 1
    MemLongPtr(VarPtr(b(i))) = pTargetFunc: i = i + 4
    ' call eax (__stdcall cleans 5*4 = 20 bytes, leaving retval at [esp])
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    ' pop eax (EAX = retval Long HRESULT; esp restored to caller return address)
    b(i) = &H58: i = i + 1
    ' ret 0x000C (clean 3 DWORDs: this, arg1, arg2)
    b(i) = &HC2: b(i + 1) = &HC: b(i + 2) = &H0: i = i + 3
#End If

    ' --- PtrSize _______ ---
    Dim k As Long
    For k = 0 To THUNK_BUF_SIZE - 1 Step PtrSize
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k
End Sub


' ============================================================
' _______________ 8 bytes (LongPtr) .
' ============================================================
Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(offset)))
End Function


' ============================================================
' ____vTable.
' ============================================================
Private Function GetClassMethodAddrAtFixedSlot( _
    ByVal cls As Object, ByVal slotIndex As Long) As LongPtr

    If cls Is Nothing Then Exit Function
    Dim pObj As LongPtr: pObj = ObjPtr(cls)
    If pObj = 0 Then Exit Function

    Dim pVTable As LongPtr
    pVTable = ReadLongPtr(pObj)
    If pVTable = 0 Then Exit Function

    GetClassMethodAddrAtFixedSlot = ReadLongPtr(pVTable + slotIndex * PtrSize)
End Function


' ============================================================
' IUnknown ____(_______)
'
'   _9.2______:
'     - this __ SlotIndexFromVTableObjAddr _ idx .
'     - m_handlers(idx) .
'     - .
'
'   _9.7a ________ IID ____:
'     - Handler_QueryInterface _ riid ____ IID .
'     - IID_IUnknown _______ HandlerKind ________ IID___ S_OK
'     - _____ ppvObject = 0 + E_NOINTERFACE
'     - ____ riid.Data1 _ Hex _____ (________________)
'
'   COM ________:
'     - ___ (S_OK ___) _ AddRef .
'     - ___ (E_NOINTERFACE ___) _ AddRef ___, ppvObject _ NULL .
'
'   ___________________________.
' ============================================================
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

Private Function Handler_QueryInterface( _
    ByVal this As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long

    ' --- __: riid _ NULL __ E_POINTER ---
    If riid = 0 Then
        ppvObject = 0
        Handler_QueryInterface = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' --- __: this _____ E_NOINTERFACE ---
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(this)
    If idx < 0 Then
        ppvObject = 0
        Handler_QueryInterface = E_NOINTERFACE
        Exit Function
    End If
    If m_handlers(idx) Is Nothing Then
        ppvObject = 0
        Handler_QueryInterface = E_NOINTERFACE
        Exit Function
    End If

    ' --- riid _ IID_IUnknown ___ ---
    If IsEqualGUIDInPlace(riid, m_iidIUnknown) Then
        ppvObject = this
        HandlerAddRefInternal this
        Handler_QueryInterface = S_OK
        Exit Function
    End If

    ' --- riid ________ IID____ ---
    Dim kind As HandlerKind
    kind = m_handlers(idx).kind
    If IsEqualGUIDInPlace(riid, m_iidTable(kind)) Then
        ppvObject = this
        HandlerAddRefInternal this
        Handler_QueryInterface = S_OK
        Exit Function
    End If

    ' --- __ IID _______: E_NOINTERFACE ---
    '   riid.Data1 (__ 4 byte, little-endian _ Long) ______.
    '   QueryInterface ________ WebView2 ___ IID .
    '   _____ (___ IID = IMarshal / IAgileObject ________).
    Dim data1 As Long
    data1 = LongPtrLowDword(ReadLongPtr(riid))
    Debug.Print "  QI rejected [idx=" & idx & " kind=" & kind & _
                "] riid.Data1=&H" & Hex(data1) & " -> E_NOINTERFACE"
    ppvObject = 0
    Handler_QueryInterface = E_NOINTERFACE
End Function

Private Function Handler_AddRef(ByVal this As LongPtr) As Long
    Handler_AddRef = HandlerAddRefInternal(this)
End Function

Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = HandlerReleaseInternal(this)
End Function


' ============================================================
' IsEqualGUIDInPlace (_9.7a _____)
'
'   2 __ GUID (16 byte) _____.
'   ___ riid (____, WebView2 ______),
'   _____ m_iidTable / m_iidIUnknown (VBA __ GUID ___).
'
'   ____: LongPtr (= x64 _ 8 byte) _ 2 ___________.
'     - __ 8 byte: Data1 (4) + Data2 (2) + Data3 (2)
'     - __ 8 byte: Data4(0..7)
'
'   VBA __ GUID _________ VarPtr _____.
'   x64 __ GUID __ 16 byte ____ (_______, .
'   Win32 ___ little-endian ___). ___ Win32 GUID .
'   VBA _ Type _____ (_______) ________.
'
'   ______:
'     - VBA _ Type ____________ 16 byte .
'     - Long / Integer _ x64 _____ little-endian .
'   ____ x64 VBA _______, _______ IID .
'   ____ (riid _ m_iidTable ______) __, _______.
' ============================================================
Private Function IsEqualGUIDInPlace(ByVal pRiid As LongPtr, ByRef refGuid As GUID) As Boolean
    Dim pRef As LongPtr
    pRef = VarPtr(refGuid)

#If Win64 Then
    If ReadLongPtr(pRiid) <> ReadLongPtr(pRef) Then Exit Function
    If ReadLongPtr(pRiid + 8) <> ReadLongPtr(pRef + 8) Then Exit Function
#Else
    If ReadLongPtr(pRiid) <> ReadLongPtr(pRef) Then Exit Function
    If ReadLongPtr(pRiid + 4) <> ReadLongPtr(pRef + 4) Then Exit Function
    If ReadLongPtr(pRiid + 8) <> ReadLongPtr(pRef + 8) Then Exit Function
    If ReadLongPtr(pRiid + 12) <> ReadLongPtr(pRef + 12) Then Exit Function
#End If

    IsEqualGUIDInPlace = True
End Function


Private Function LongPtrLowDword(ByVal v As LongPtr) As Long
#If Win64 Then
    Dim u As LongLong
    u = CLngLng(v) And &HFFFFFFFF^
    If u > &H7FFFFFFF^ Then
        LongPtrLowDword = CLng(u - &H100000000^)
    Else
        LongPtrLowDword = CLng(u)
    End If
#Else
    LongPtrLowDword = v
#End If
End Function


' ============================================================
' InitIIDTable (_9.7a _____)
'
'   m_iidIUnknown _ m_iidTable(HK_xxx) ______.
'   Thunks_Init _____ 1 _______.
'
'   IID ____: WebView2.h (Microsoft Edge WebView2 SDK).
'   _ IID _ WebView2_IID_vtable_reference.md _____.
'
'   GUID ______________:
'     Data1 (Long, 4 byte little-endian): .
'     Data2 (Integer, 2 byte little-endian): ______ (_________,
'       &H8000 ___ & _______ Long _____ CInt _ Integer .
'       _______. _____ FillGUID _ String _________,
'       __________________ &H _____ (____ 29) ___)
' ============================================================
Private Sub InitIIDTable()
    ' --- IID_IUnknown (Win32 __, {00000000-0000-0000-C000-000000000046}) ---
    FillGUID m_iidIUnknown, "00000000-0000-0000-C000-000000000046"

    ' --- HandlerKind ______ IID ---

    ' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
    FillGUID m_iidTable(HK_EnvironmentCompleted), _
             "4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d"

    ' ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
    FillGUID m_iidTable(HK_ControllerCompleted), _
             "6c4819f3-c9b7-4260-8127-c9f5bde7f68c"

    ' ICoreWebView2NavigationStartingEventHandler
    FillGUID m_iidTable(HK_NavigationStarting), _
             "9adbe429-f36d-432b-9ddc-f8881fbd76e3"

    ' ICoreWebView2NavigationCompletedEventHandler
    FillGUID m_iidTable(HK_NavigationCompleted), _
             "d33a35bf-1c49-4f98-93ab-006e0533fe1c"

    ' ICoreWebView2WebMessageReceivedEventHandler
    FillGUID m_iidTable(HK_WebMessageReceived), _
             "57213f19-00e6-49fa-8e07-898ea01ecbd2"

    ' ICoreWebView2DocumentTitleChangedEventHandler
    FillGUID m_iidTable(HK_DocumentTitleChanged), _
             "f5f2b923-953e-4042-9f95-f3a118e1afd4"

    ' ICoreWebView2NewWindowRequestedEventHandler
    FillGUID m_iidTable(HK_NewWindowRequested), _
             "d4c185fe-c81c-4989-97af-2d3fa7ab5651"

    ' ICoreWebView2ExecuteScriptCompletedHandler (_9.8c ___)
    FillGUID m_iidTable(HK_ExecuteScriptCompleted), _
             "49511172-cc67-4bca-9923-137112f4c4cc"

    ' ICoreWebView2HistoryChangedEventHandler (_9.9b ___)
    '   __: WebView2.h L3790
    FillGUID m_iidTable(HK_HistoryChanged), _
             "c79a420c-efd9-4058-9295-3e8b4bcab645"

    ' ICoreWebView2DOMContentLoadedEventHandler (_9.9b ___)
    '   __: WebView2.h L5718
    FillGUID m_iidTable(HK_DOMContentLoaded), _
             "4bac7e9c-199e-49ed-87ed-249303acf019"

    ' ICoreWebView2WebResourceRequestedEventHandler (N-1 ___)
    '   __: WebView2.h _ MIDL_INTERFACE _ (*______________*)
    FillGUID m_iidTable(HK_WebResourceRequested), _
             "ab00b74c-15f1-4646-80e8-e76341d25d71"

    ' ICoreWebView2WebResourceResponseReceivedEventHandler (N-3 ___)
    '   __: WebView2.h _ MIDL_INTERFACE .
    FillGUID m_iidTable(HK_WebResourceResponseReceived), _
             "7de9898a-24f5-40c3-a2de-d4f458e69828"

    ' ICoreWebView2WebResourceResponseViewGetContentCompletedHandler (N-4 ___)
    '   __: WebView2.h _ MIDL_INTERFACE .
    FillGUID m_iidTable(HK_ResponseGetContent), _
             "875738e1-9fa2-40e3-8b74-2e8972dd6fe7"
End Sub


' ============================================================
' FillGUID (_9.7a _____, _9.9a _ Public __)
'
'   "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" ________ GUID _______.
'   _____ &H ___________ (____ 29) ___________,
'   ________________.
'
'   ___: "4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d"
'   __:
'     g(0..7)   = "4e8a3389"  -> Data1 (Long)
'     g(9..12)  = "c9d8"      -> Data2 (Integer)
'     g(14..17) = "4bd2"      -> Data3 (Integer)
'     g(19..22) = "b6b5"      -> Data4(0..1)
'     g(24..35) = "124fee6cc14d" -> Data4(2..7)
'
'   _ 16 _____ HexStrToLong / HexStrToInt ____.
'   Data2 / Data3 _ Integer (____ 16 bit). &HC9D8 .
'   ___ CInt() ___ Overflow ____, __ Long .
'   _______________. .
'   HexStrToInt ___.
'
'   _9.9a _ Public __:
'     Wv2Pane.EnsureView2 __, QueryInterface ______ GUID .
'     ______________. Wv2Thunks _ InitIIDTable .
'     ______, _____________. GUID _ (Public Type) .
'     ______________, _______ (= ____ 50 .
'     __, 9.9a ________).
' ============================================================
Public Sub FillGUID(ByRef g As GUID, ByVal s As String)
    g.data1 = HexStrToLong(Mid$(s, 1, 8))
    g.Data2 = HexStrToInt(Mid$(s, 10, 4))
    g.Data3 = HexStrToInt(Mid$(s, 15, 4))

    ' Data4(0..7): __ 4 ____ (Mid 20..23) + __ 12 _ (Mid 25..36)
    g.Data4(0) = CByte("&H" & Mid$(s, 20, 2))
    g.Data4(1) = CByte("&H" & Mid$(s, 22, 2))
    g.Data4(2) = CByte("&H" & Mid$(s, 25, 2))
    g.Data4(3) = CByte("&H" & Mid$(s, 27, 2))
    g.Data4(4) = CByte("&H" & Mid$(s, 29, 2))
    g.Data4(5) = CByte("&H" & Mid$(s, 31, 2))
    g.Data4(6) = CByte("&H" & Mid$(s, 33, 2))
    g.Data4(7) = CByte("&H" & Mid$(s, 35, 2))
End Sub


Private Function HexStrToLong(ByVal s As String) As Long
    HexStrToLong = CLng("&H" & s)
End Function


' ============================================================
' HexStrToInt (_9.7a _____)
'
'   4 __ 16 _____ Integer (____ 16 bit) _____.
'   __ &H8000 ____________.
'
'   VBA _ CInt("&H...") _ &H8000 ___ Overflow .
'   (= _32768 _ Integer ______________, _______ 29
'   ________).
'
'   ______ CLng ___ Long ___, &H8000 .
'   2 __________ CInt __.
' ============================================================
Private Function HexStrToInt(ByVal s As String) As Integer
    Dim v As Long
    v = CLng("&H" & s)                        ' 0..&HFFFF _ Long .
    If v >= &H8000& Then
        HexStrToInt = CInt(v - &H10000)       ' __ Integer .
    Else
        HexStrToInt = CInt(v)
    End If
End Function


' ============================================================
' HandlerAddRefInternal / HandlerReleaseInternal
'
'   _9.2______:
'     - this (= pSlot + VTABLE_OBJ_OFFSET) __ idx .
'     - m_handlers(idx) .
'     - ___ this ____ 0 .
'
'   _9.3a ______:
'     - HandlerReleaseInternal _________, .
'       _gHandler Is h______ (gHandler __________)
'
'   ___: ____ refcount
' ============================================================
Private Function HandlerAddRefInternal(ByVal this As LongPtr) As Long
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(this)
    If idx < 0 Then Exit Function
    If m_handlers(idx) Is Nothing Then Exit Function

    Dim n As Long
    n = m_handlers(idx).RefCount + 1
    m_handlers(idx).RefCount = n
    Debug.Print "  AddRef [idx=" & idx & "] -> refCount = " & n
    HandlerAddRefInternal = n
End Function

Private Function HandlerReleaseInternal(ByVal this As LongPtr) As Long
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(this)
    If idx < 0 Then Exit Function
    If m_handlers(idx) Is Nothing Then Exit Function

    Dim n As Long
    n = m_handlers(idx).RefCount - 1
    If n < 0 Then n = 0                ' __ (_______)
    m_handlers(idx).RefCount = n
    Debug.Print "  Release [idx=" & idx & "] -> refCount = " & n

    If n = 0 Then
        ' refcount _ 0 ______, _________________.
        ' _____ (_9.1______):
        '   1. ClearOwner ________ (ComCallbackHandler -> ___________)
        '   2. _____________ m_handlers(idx) _ Nothing .
        '   3. Thunks_ReleaseSlot .
        '   4. _______ Nothing _ (___ ComCallbackHandler ________)
        '
        ' _9.3a ____: _ 4. __gHandler Is h __ gHandler .
        ' ____________ (gHandler ________).

        Dim h As ComCallbackHandler
        Set h = m_handlers(idx)

        Dim pSlot As LongPtr
        pSlot = h.Slot

        ' 1. _______ (ComCallbackHandler.m_owner = Nothing)
        '    __________ (Wv2Environment _) _______,
        '    ________ ComCallbackHandler .
        '    ___ GC ___________.
        h.ClearOwner

        ' 2. m_handlers(idx) _ Nothing .
        Set m_handlers(idx) = Nothing

        ' 3. .
        Thunks_ReleaseSlot pSlot

        ' 4. h .
        Set h = Nothing

        Debug.Print "  Release [idx=" & idx & "] -> refCount = 0 -> Slot " & idx & " auto-released"
    End If

    HandlerReleaseInternal = n
End Function


' ============================================================
' PointerAccessor .
' ============================================================
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal Ptr As LongPtr)
    ptrs(0) = Ptr
End Sub


