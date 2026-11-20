Attribute VB_Name = "Wv2Environment"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
''''''''''''''''''''''''''''''''''
' --- Wv2Environment.cls _9.8a__ (__: Class8) ---
'
'   _9.8a ____ (____, ________):
'     - ____ Class8 -> Wv2Environment .
'     - Class7 _____ ComCallbackHandler .
'     - _____________ "Class8.xxx" -> "Wv2Environment.xxx" .
'     - _9.3a .
'
'   __: ICoreWebView2Environment _ VBA ____.
'     - HK_EnvironmentCompleted ____ (ComCallbackHandler) _ 1 .
'     - CreateCoreWebView2EnvironmentWithOptions ____ WebView2 .
'     - ______ ComCallbackHandler.Handler_Invoke ____, __ m_kind
'       _______ OnEnvironmentCompleted .
'     - _____ pEnvironment ____ COM ____ AddRef ___,
'       Class_Terminate _ Release .
'
'   ____ (_ D1-_: 1 ____________):
'     - HK_EnvironmentCompleted __1 ______________.
'     - OnEnvironmentCompleted ____ Set m_handler = Nothing .
'       ComCallbackHandler _________.
'     - WebView2 _ ComCallbackHandler _ Release ___ refcount=0 ___,
'       _____________ (ComCallbackHandler.ClearOwner .
'       m_owner=Me ____, Wv2Environment <--> ComCallbackHandler .
'       _______).
'     - ___ Wv2Environment (Me) .
'       ________, Class_Terminate _ m_pEnvironment->Release ___.
'
'   ____:
'     Es_Created       : Class_Initialize .
'     Es_Waiting       : Init _ API _____, .
'     Es_Ready         : OnEnvironmentCompleted _ errorCode=0 .
'     Es_Failed        : OnEnvironmentCompleted _ errorCode<>0 .
''''''''''''''''''''''''''''''''''
Option Explicit

' Supports both 64-bit (x64) and 32-bit (x86) VBA7 Office hosts


Public Enum EnvironmentState
    Es_Created = 0
    Es_Waiting = 1
    Es_Ready = 2
    Es_Failed = 3
End Enum


' --- ____ ---
' m_handler:      HK_EnvironmentCompleted .
'                 (OnEnvironmentCompleted __ Nothing ___)
' m_pEnvironment: WebView2 _______ ICoreWebView2Environment
'                 (Class_Terminate _ Release _______)
Private m_handler      As ComCallbackHandler
Private m_pEnvironment As LongPtr
Private m_state        As EnvironmentState
Private m_lastError    As Long


' ============================================================
' Class_Initialize / Class_Terminate
' ============================================================
Private Sub Class_Initialize()
    m_state = Es_Created
    m_pEnvironment = 0
    m_lastError = 0
End Sub

Private Sub Class_Terminate()
    ' COM _____: Init ______ pEnvironment .
    ' OnEnvironmentCompleted _ AddRef ____ VBA _______________,
    ' ___ Release ________.
    If m_pEnvironment <> 0 Then
        Debug.Print "Wv2Environment.Terminate: Releasing m_pEnvironment = " & m_pEnvironment
        ComRelease m_pEnvironment
        m_pEnvironment = 0
    End If

    ' m_handler ___ OnEnvironmentCompleted _ Nothing __________,
    ' __ Init _________________, _____.
    ' (ComCallbackHandler _____ m_handlers(idx) ____________,
    '  ______________________. )
    If Not (m_handler Is Nothing) Then
        Debug.Print "Wv2Environment.Terminate: m_handler __________ (______)"
        Set m_handler = Nothing
    End If
End Sub


' ============================================================
' Init
'
'   ICoreWebView2Environment ________.
'   ___________________________ (___).
'   _____ OnEnvironmentCompleted ____.
'
'   ______ env.IsReady _ True _____ DoEvents ___.
' ============================================================
Public Sub Init(ByVal userDataPath As String)

    If m_state <> Es_Created Then
        MsgBox "Wv2Environment.Init _________", vbCritical
        Exit Sub
    End If

    ' --- 1. ____ (ComCallbackHandler) _ 1 _ Acquire ---
    '   AcquireHandlerFor ___:
    '     - ComCallbackHandler _ New
    '     - Thunks_AcquireSlot _______ + .
    '     - ComCallbackHandler.Init(HK_EnvironmentCompleted, owner=Me, pSlot)
    '     - m_handlers(idx) .
    '   _______, ______ ComCallbackHandler ______.
    Set m_handler = AcquireHandlerFor(HK_EnvironmentCompleted, Me)
    If m_handler Is Nothing Then
        MsgBox "Wv2Environment.Init: _____ Acquire ___", vbCritical
        m_state = Es_Failed
        Exit Sub
    End If

    ' --- 2. WebView2 ___ obj _______ ---
    '   pSlot + VTABLE_OBJ_OFFSET _ vtableObj ___.
    '   _______ obj ___ WebView2 ____, __ AddRef/Release/Invoke .
    '   this _______.
    Dim pVTableObj As LongPtr
    pVTableObj = m_handler.Slot + VTABLE_OBJ_OFFSET

    ' --- 3. _____________ ---
    Dim path As String
    path = userDataPath
    If LenB(path) = 0 Then path = "C:\Temp\VBA_WebView2"
    EnsureFolder path

    ' --- 4. CreateCoreWebView2EnvironmentWithOptions ___ ---
    '
    '   __: m_state = Es_Waiting _ API ____________.
    '   ___, WebView2Loader _ API .
    '   ________ (_________):
    '
    '     CreateCoreWebView2EnvironmentWithOptions .
    '       __ ___ WebView2 _ AddRef -> Invoke -> Release .
    '          __ Wv2Environment.OnEnvironmentCompleted _____ m_state = Es_Ready
    '       <- API _ S_OK .
    '
    '   ____, API _________ m_state = Es_Waiting _____,
    '   ____________ Es_Ready _ Es_Waiting __________,
    '   ______ IsReady _______ True _____.
    '
    '   __ Es_Waiting _______, .
    '   _____ OnEnvironmentCompleted ______ (Es_Ready / Es_Failed) .
    '   _______. API _______ Es_Failed ______.
    m_state = Es_Waiting

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(path), 0, pVTableObj)

    If hr <> 0 Then
        MsgBox "CreateCoreWebView2EnvironmentWithOptions failed. HRESULT: &H" & Hex(hr), _
               vbCritical
        m_state = Es_Failed
        m_lastError = hr
        ' ___________9.3a _____.
        ' WebView2 _ AddRef _____________, refcount=0 .
        ' m_handlers(idx) ______________.
        ' ___ Thunks_Shutdown _______________________.
        ' _9.3b ___ Controller _________________________.
        Exit Sub
    End If
End Sub


' ============================================================
' OnEnvironmentCompleted
'
'   ComCallbackHandler.Handler_Invoke _ HK_EnvironmentCompleted .
'   m_owner.OnEnvironmentCompleted _______________.
'
'   _______ Public Sub __:
'     ComCallbackHandler __ m_owner As Object .
'     (IDispatch::Invoke __) ___. Friend __ IDispatch .
'     ________, _________________________. .
'     ___________. Public ____ IDispatch ____________.
'
'   __:
'     errorCode    : HRESULT (S_OK = 0 ___)
'     pEnvironment : ICoreWebView2Environment .
'                    (WebView2 _____ AddRef _________.
'                     ____________ WebView2 _ Release ____,
'                     VBA _____________ AddRef ___. )
' ============================================================
Public Sub OnEnvironmentCompleted( _
    ByVal errorCode As Long, _
    ByVal pEnvironment As LongPtr)

    Debug.Print "Wv2Environment.OnEnvironmentCompleted: errorCode=" & errorCode & _
                ", pEnvironment=" & pEnvironment

    m_lastError = errorCode

    If errorCode = 0 And pEnvironment <> 0 Then
        ' VBA .
        m_pEnvironment = pEnvironment
        ComAddRef m_pEnvironment
        m_state = Es_Ready
    Else
        m_pEnvironment = 0
        m_state = Es_Failed
    End If

    ' 1 ____________ (_ D1-_):
    '   ____________________.
    '   _________ ComCallbackHandler.Handler_Invoke ___ WebView2 .
    '   __ refcount=1 _______ (WebView2 ____ Release ___).
    '   ___ m_handler ____, Wv2Environment -> ComCallbackHandler ________,
    '   ____ m_handlers(idx) ______________ ComCallbackHandler .
    '   ______. WebView2 _ Release ____ refcount=0 ______,
    '   ___ HandlerReleaseInternal __________,
    '   ComCallbackHandler.ClearOwner _ m_owner (= Me Wv2Environment) ________.
    Set m_handler = Nothing
End Sub


' ============================================================
' .
' ============================================================
Public Property Get state() As EnvironmentState
    state = m_state
End Property

Public Property Get IsReady() As Boolean
    IsReady = (m_state = Es_Ready)
End Property

Public Property Get IsFailed() As Boolean
    IsFailed = (m_state = Es_Failed)
End Property

Public Property Get LastError() As Long
    LastError = m_lastError
End Property

Public Property Get EnvironmentPtr() As LongPtr
    EnvironmentPtr = m_pEnvironment
End Property





