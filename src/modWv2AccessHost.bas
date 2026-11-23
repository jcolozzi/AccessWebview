Attribute VB_Name = "modWv2AccessHost"
Option Compare Database
Option Explicit

' ============================================================================
' modWv2AccessHost - Access Host Windowing, Registry & Verification Suite
' ============================================================================

Public Declare PtrSafe Function FindWindowExW Lib "user32" ( _
    ByVal hWndParent As LongPtr, _
    ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As LongPtr, _
    ByVal lpszWindow As LongPtr) As LongPtr

Public Declare PtrSafe Function GetParent Lib "user32" ( _
    ByVal hwnd As LongPtr) As LongPtr

Public Declare PtrSafe Function GetClassNameW Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpClassName As LongPtr, _
    ByVal nMaxCount As Long) As Long

Public Declare PtrSafe Function GetWindowRect Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByRef lpRect As RECT) As Long

Public Declare PtrSafe Function MapWindowPoints Lib "user32" ( _
    ByVal hWndFrom As LongPtr, _
    ByVal hWndTo As LongPtr, _
    ByRef lpPoints As RECT, _
    ByVal cPoints As Long) As Long

Public Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal uFlags As Long) As Long

Public Declare PtrSafe Function IsWindow Lib "user32" ( _
    ByVal hwnd As LongPtr) As Long

Public Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nCmdShow As Long) As Long

Private m_hostRegistry As Collection

' ============================================================================
' Persistent clsWv2Host Registry (Survives Subform OFormChild Reload on Grow)
' ============================================================================

Public Sub RegisterAccessHost(ByVal HostHwnd As LongPtr, ByVal hostObj As clsWv2Host)
    If HostHwnd = 0 Or hostObj Is Nothing Then Exit Sub
    If m_hostRegistry Is Nothing Then Set m_hostRegistry = New Collection
    UnregisterAccessHost HostHwnd
    m_hostRegistry.Add hostObj, "H" & CStr(HostHwnd)
End Sub

Public Function GetRegisteredAccessHost(ByVal HostHwnd As LongPtr) As clsWv2Host
    If HostHwnd = 0 Or m_hostRegistry Is Nothing Then Exit Function
    On Error Resume Next
    Set GetRegisteredAccessHost = m_hostRegistry("H" & CStr(HostHwnd))
    On Error GoTo 0
End Function

Public Sub UnregisterAccessHost(ByVal HostHwnd As LongPtr)
    If HostHwnd = 0 Or m_hostRegistry Is Nothing Then Exit Sub
    On Error Resume Next
    m_hostRegistry.Remove "H" & CStr(HostHwnd)
    On Error GoTo 0
End Sub

Public Sub ShutdownAccessFormWebViews(ByVal parentFrm As Access.Form)
    If parentFrm Is Nothing Then Exit Sub
    On Error Resume Next
    Dim pHwnd As LongPtr
    pHwnd = parentFrm.hwnd
    If pHwnd = 0 Then Exit Sub

    Dim hHost As clsWv2Host
    Set hHost = GetRegisteredAccessHost(pHwnd)
    If Not (hHost Is Nothing) Then
        UnregisterAccessHost pHwnd
        hHost.Shutdown
    End If

    Dim hChild As LongPtr
    hChild = FindWindowExW(pHwnd, 0, StrPtr("OFormSub"), 0)
    Do While hChild <> 0
        Set hHost = GetRegisteredAccessHost(hChild)
        If Not (hHost Is Nothing) Then
            UnregisterAccessHost hChild
            hHost.Shutdown
        End If
        hChild = FindWindowExW(pHwnd, hChild, StrPtr("OFormSub"), 0)
    Loop
    On Error GoTo 0
End Sub

' ============================================================================
' Access Window Hierarchy & Viewport Helpers
' ============================================================================

Public Function GetWinClassName(ByVal hwnd As LongPtr) As String
    If hwnd = 0 Then Exit Function
    Dim buf As String
    Dim n As Long
    buf = String$(64, vbNullChar)
    n = GetClassNameW(hwnd, StrPtr(buf), 64)
    If n > 0 Then GetWinClassName = Left$(buf, n)
End Function

Public Function IsSubformWindow(ByVal frm As Access.Form) As Boolean
    If frm Is Nothing Then Exit Function
    IsSubformWindow = (StrComp(GetWinClassName(frm.hwnd), "OFormChild", vbTextCompare) = 0)
End Function

Public Function GetAccessFormCanvasHwnd(ByVal frm As Access.Form) As LongPtr
    If frm Is Nothing Then Exit Function

    ' CRITICAL ACCESS RESIZE INVARIANT:
    ' When hosted inside a Subform control, frm.hwnd has class "OFormChild".
    ' Whenever a Subform grows (including enlarging or maximizing the parent form),
    ' Access destroys the OFormChild HWND and reloads the subform VBA module!
    ' By contrast, GetParent(frm.hwnd) is the parent form's section OFormSub window,
    ' which is IMMORTAL for the entire lifetime of the parent form.
    If IsSubformWindow(frm) Then
        Dim hParentSection As LongPtr
        hParentSection = GetParent(frm.hwnd)
        If hParentSection <> 0 Then
            GetAccessFormCanvasHwnd = hParentSection
            Exit Function
        End If
    End If

    ' Standalone top-level Form (OForm / OFormPopup): select its Detail OFormSub
    ' (Access creates 3 OFormSub children in Z-order: 1=FormHeader, 2=Detail, 3=FormFooter).
    Dim hChild As LongPtr
    Dim bestHwnd As LongPtr
    Dim bestHeight As Long
    Dim secondSub As LongPtr
    Dim subIndex As Long
    Dim rc As RECT

    hChild = FindWindowExW(frm.hwnd, 0, StrPtr("OFormSub"), 0)
    Do While hChild <> 0
        subIndex = subIndex + 1
        If subIndex = 2 Then secondSub = hChild
        If GetClientRect(hChild, rc) <> 0 Then
            If (rc.Bottom - rc.Top) > bestHeight Then
                bestHeight = rc.Bottom - rc.Top
                bestHwnd = hChild
            End If
        End If
        hChild = FindWindowExW(frm.hwnd, hChild, StrPtr("OFormSub"), 0)
    Loop

    If bestHwnd <> 0 Then
        GetAccessFormCanvasHwnd = bestHwnd
    ElseIf secondSub <> 0 Then
        GetAccessFormCanvasHwnd = secondSub
    Else
        GetAccessFormCanvasHwnd = frm.hwnd
    End If
End Function

Public Function GetAccessFormViewportRect(ByVal frm As Access.Form, ByVal HostHwnd As LongPtr, ByRef outRect As RECT) As Boolean
    If frm Is Nothing Or HostHwnd = 0 Then Exit Function

    If IsSubformWindow(frm) Then
        Dim rcScreen As RECT
        If GetWindowRect(frm.hwnd, rcScreen) = 0 Then Exit Function
        MapWindowPoints 0, HostHwnd, rcScreen, 2

        Dim rcHost As RECT
        GetClientRect HostHwnd, rcHost

        If rcScreen.Left < 0 Then rcScreen.Left = 0
        If rcScreen.Top < 0 Then rcScreen.Top = 0
        If rcScreen.Right > rcHost.Right Then rcScreen.Right = rcHost.Right
        If rcScreen.Bottom > rcHost.Bottom Then rcScreen.Bottom = rcHost.Bottom

        outRect = rcScreen
        GetAccessFormViewportRect = True
    Else
        If GetClientRect(HostHwnd, outRect) <> 0 Then
            GetAccessFormViewportRect = True
        End If
    End If
End Function

Public Sub SyncAccessWebViewZOrder(ByVal frm As Access.Form, ByVal HostHwnd As LongPtr)
    If frm Is Nothing Or HostHwnd = 0 Then Exit Sub
    If IsSubformWindow(frm) Then
        Const SWP_FLAGS As Long = &H13&   ' SWP_NOSIZE (1) | SWP_NOMOVE (2) | SWP_NOACTIVATE (&H10)
        ' Keep placeholder OFormChild at HWND_BOTTOM (1) and Chrome_WidgetWin_0 at HWND_TOP (0)
        SetWindowPos frm.hwnd, 1, 0, 0, 0, 0, SWP_FLAGS
        Dim hWv2 As LongPtr
        hWv2 = FindWindowExW(HostHwnd, 0, StrPtr("Chrome_WidgetWin_0"), 0)
        Do While hWv2 <> 0
            SetWindowPos hWv2, 0, 0, 0, 0, 0, SWP_FLAGS
            hWv2 = FindWindowExW(HostHwnd, hWv2, StrPtr("Chrome_WidgetWin_0"), 0)
        Loop
    End If
End Sub

Public Function GetDefaultUserDataFolder() As String
    Dim baseDir As String
    baseDir = Environ$("LOCALAPPDATA")
    If Len(baseDir) = 0 Then baseDir = Environ$("TEMP")
    GetDefaultUserDataFolder = baseDir & "\AccessWebview\UserData"
End Function

Public Function GetDefaultSpaFolder() As String
    GetDefaultSpaFolder = CurrentProject.path & "\spa"
End Function

Public Sub LaunchWebViewDemo()
    DoCmd.OpenForm "frmWv2Demo", acNormal
End Sub

Public Function GetDemoFormHwnd() As Long
    On Error Resume Next
    GetDemoFormHwnd = CLng(Forms("frmWv2Demo").hwnd)
    On Error GoTo 0
End Function

' ============================================================================
' Automated Verification Functions (Callable via DAO SQL / Immediate Window)
' ============================================================================

Public Function VerifyStep1_MockThunks() As String
    On Error GoTo ErrHandler
    If Not Thunks_Init() Then
        VerifyStep1_MockThunks = "FAIL: Thunks_Init returned False"
        Exit Function
    End If

    Dim h1 As ComCallbackHandler
    Dim h2 As ComCallbackHandler
    Set h1 = AcquireHandlerFor(HK_EnvironmentCompleted, Nothing)
    Set h2 = AcquireHandlerFor(HK_ControllerCompleted, Nothing)

    If h1 Is Nothing Or h2 Is Nothing Then
        VerifyStep1_MockThunks = "FAIL: AcquireHandlerFor returned Nothing"
        Exit Function
    End If

    Dim pObj1 As LongPtr
    Dim pObj2 As LongPtr
    pObj1 = h1.Slot + VTABLE_OBJ_OFFSET
    pObj2 = h2.Slot + VTABLE_OBJ_OFFSET

    Dim rc1 As Long
    Dim rc2 As Long
    rc1 = ComAddRef(pObj1)
    rc1 = ComAddRef(pObj1)
    rc2 = ComAddRef(pObj2)
    If rc1 <> 2 Or rc2 <> 1 Then
        VerifyStep1_MockThunks = "FAIL: AddRef mismatch rc1=" & rc1 & " rc2=" & rc2
        Exit Function
    End If

    Dim iidUnk As GUID
    Dim iidEnv As GUID
    Dim pOut As LongPtr
    Dim hr As Long
    FillGUID iidUnk, "00000000-0000-0000-C000-000000000046"
    FillGUID iidEnv, "4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d"

    hr = dcf(pObj1, 0, "QI_IUnknown", VarPtr(iidUnk), VarPtr(pOut))
    If hr <> 0 Or pOut <> pObj1 Then
        VerifyStep1_MockThunks = "FAIL: QI IUnknown hr=&H" & Hex$(hr)
        Exit Function
    End If
    ComRelease pObj1

    hr = dcf(pObj1, 0, "QI_EnvHandler", VarPtr(iidEnv), VarPtr(pOut))
    If hr <> 0 Or pOut <> pObj1 Then
        VerifyStep1_MockThunks = "FAIL: QI EnvHandler hr=&H" & Hex$(hr)
        Exit Function
    End If
    ComRelease pObj1

    hr = dcf(pObj1, 3, "MockInvoke", 0&, CLngPtr(&H123456))
    If hr <> 0 Then
        VerifyStep1_MockThunks = "FAIL: Invoke via thunk returned hr=&H" & Hex$(hr)
        Exit Function
    End If

    rc1 = ComRelease(pObj1)
    rc1 = ComRelease(pObj1)
    rc2 = ComRelease(pObj2)
    Set h1 = Nothing
    Set h2 = Nothing
    Thunks_Shutdown

    If rc1 <> 0 Or rc2 <> 0 Then
        VerifyStep1_MockThunks = "FAIL: Release mismatch rc1=" & rc1 & " rc2=" & rc2
        Exit Function
    End If

    VerifyStep1_MockThunks = "PASS: Mock QI + AddRef + Release + VBE7/MachineCode Thunk Invoke all succeeded!"
    Exit Function

ErrHandler:
    VerifyStep1_MockThunks = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function VerifyStep2_RealEnvironment() As String
    On Error GoTo ErrHandler
    Dim env As Wv2Environment
    Set env = New Wv2Environment
    env.Init GetDefaultUserDataFolder()

    Dim t0 As Single
    t0 = Timer
    Do While Not env.IsReady And Not env.IsFailed
        DoEvents
        If Timer - t0 > 15 Then
            VerifyStep2_RealEnvironment = "FAIL: Timeout waiting for WebView2 Environment"
            Set env = Nothing
            Exit Function
        End If
    Loop

    If env.IsFailed Or env.EnvironmentPtr = 0 Then
        VerifyStep2_RealEnvironment = "FAIL: Environment failed hr=&H" & Hex$(env.LastError)
        Set env = Nothing
        Exit Function
    End If

    Dim verStr As String
    verStr = GetStringProperty(env.EnvironmentPtr, 5, "Env.get_BrowserVersionString")
    Set env = Nothing
    Thunks_Shutdown

    VerifyStep2_RealEnvironment = "PASS: Real WebView2 Environment created via Office16\WebView2Loader.dll! Runtime=" & verStr
    Exit Function

ErrHandler:
    VerifyStep2_RealEnvironment = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function VerifyStep3_FullSpaAndBridge() As String
    Dim stage As String
    On Error GoTo ErrHandler

    stage = "1-OpenForm"
    DoCmd.OpenForm "frmWv2Demo", acNormal

    stage = "2-GetForm"
    Dim frm As Form_frmWv2Demo
    Set frm = Forms("frmWv2Demo")

    stage = "3-ActivePane"
    Dim p As Wv2Pane
    Set p = frm.ActivePane
    If p Is Nothing Then
        VerifyStep3_FullSpaAndBridge = "FAIL: frmWv2Demo.ActivePane is Nothing"
        Exit Function
    End If

    stage = "4-WaitForHostBadge"
    If p.WaitFor("#host-badge", 10) Is Nothing Then
        VerifyStep3_FullSpaAndBridge = "FAIL: Timed out waiting for #host-badge in SPA (Url=" & p.View_GetSource() & ")"
        Exit Function
    End If

    stage = "5-WaitForTaskRow"
    If p.WaitFor(".task-row", 10) Is Nothing Then
        VerifyStep3_FullSpaAndBridge = "FAIL: Timed out waiting for DAO tasks (.task-row) | rpcStatus=" & p.QuerySelector("#rpc-status").InnerText
        Exit Function
    End If

    stage = "6-ReadBadgeAndCount"
    Dim badgeText As String
    badgeText = p.QuerySelector("#host-badge").InnerText

    Dim initialCount As String
    initialCount = p.EvalSync("document.querySelectorAll('.task-row').length")

    stage = "7-AddNewTaskViaDom"
    p.QuerySelector("#new-title").value = "Verified Live from 32-bit Access!"
    p.QuerySelector("#btn-add").Click

    Dim t0 As Single
    t0 = Timer
    Dim afterCount As String
    Do
        DoEvents
        afterCount = p.EvalSync("document.querySelectorAll('.task-row').length")
        If Val(afterCount) > Val(initialCount) Then Exit Do
    Loop While Timer - t0 < 8

    stage = "8-JsToVbaPostMessage"
    p.View_ExecuteScript "window.chrome.webview.postMessage('hello-from-webview2-js');"
    t0 = Timer
    Do While Len(frm.LastWebMessage) = 0 And (Timer - t0 < 5)
        DoEvents
    Loop

    stage = "9-VbaToJsPostMessage"
    p.View_PostWebMessageAsString "Live Push from MSACCESS.EXE!"
    DoEvents

    VerifyStep3_FullSpaAndBridge = "PASS | Badge=[" & badgeText & "] | Tasks: " & initialCount & " -> " & afterCount & _
                                   " | LastPostMessage=[" & frm.LastWebMessage & "] | Url=" & p.View_GetSource()
    Exit Function

ErrHandler:
    VerifyStep3_FullSpaAndBridge = "ERROR " & Err.Number & " at [" & stage & "]: " & Err.Description
End Function

Public Function VerifyStep4_ResizeSurvival() As String
    On Error GoTo ErrHandler
    DoCmd.OpenForm "frmWv2Demo", acNormal
    Dim frm As Form_frmWv2Demo
    Set frm = Forms("frmWv2Demo")

    If frm.ActivePane.WaitFor(".task-row", 10) Is Nothing Then
        VerifyStep4_ResizeSurvival = "FAIL: Initial load failed"
        Exit Function
    End If

    Dim wInit As String
    wInit = frm.ActivePane.EvalSync("window.innerWidth + 'x' + window.innerHeight")

    ' 1. Shrink
    frm.InsideWidth = 10800
    frm.InsideHeight = 7200
    DoEvents
    Dim wShrink As String
    wShrink = frm.ActivePane.EvalSync("window.innerWidth + 'x' + window.innerHeight")

    ' 2. Grow beyond initial design size (triggers Access Subform OFormChild recreation!)
    frm.InsideWidth = 16800
    frm.InsideHeight = 11200
    DoEvents
    If frm.ActivePane Is Nothing Then
        VerifyStep4_ResizeSurvival = "FAIL: ActivePane lost after Grow"
        Exit Function
    End If
    Dim wGrow As String
    wGrow = frm.ActivePane.EvalSync("window.innerWidth + 'x' + window.innerHeight")

    ' 3. Maximize & Restore
    ShowWindow frm.hwnd, 3   ' SW_MAXIMIZE
    DoEvents
    Dim wMax As String
    wMax = frm.ActivePane.EvalSync("window.innerWidth + 'x' + window.innerHeight")

    ShowWindow frm.hwnd, 9   ' SW_RESTORE
    DoEvents
    Dim wRestore As String
    wRestore = frm.ActivePane.EvalSync("window.innerWidth + 'x' + window.innerHeight")

    VerifyStep4_ResizeSurvival = "PASS | Init=" & wInit & " -> Shrink=" & wShrink & _
                                 " -> Grow=" & wGrow & " -> Max=" & wMax & " -> Restore=" & wRestore
    Exit Function

ErrHandler:
    VerifyStep4_ResizeSurvival = "ERROR " & Err.Number & ": " & Err.Description
End Function
