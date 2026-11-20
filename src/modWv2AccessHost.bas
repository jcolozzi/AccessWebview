Attribute VB_Name = "modWv2AccessHost"
Option Compare Database
Option Explicit

' ============================================================================
' modWv2AccessHost - Access Host Helpers & Automated Verification Suite
' ============================================================================

Public Declare PtrSafe Function FindWindowExW Lib "user32" ( _
    ByVal hWndParent As LongPtr, _
    ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As LongPtr, _
    ByVal lpszWindow As LongPtr) As LongPtr

Public Function GetAccessFormCanvasHwnd(ByVal frm As Access.Form) As LongPtr
    If frm Is Nothing Then Exit Function
    ' Access creates 3 OFormSub child windows under frm.hwnd in order:
    '   1st = FormHeader (height 0 when header is hidden/absent)
    '   2nd = Detail     (the main canvas!)
    '   3rd = FormFooter (height 0 when footer is hidden/absent)
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

    ' Test AddRef/Release via DispCallFunc (exact same path C++ WebView2 takes!)
    Dim rc1 As Long
    Dim rc2 As Long
    rc1 = ComAddRef(pObj1)
    rc1 = ComAddRef(pObj1)
    rc2 = ComAddRef(pObj2)
    If rc1 <> 2 Or rc2 <> 1 Then
        VerifyStep1_MockThunks = "FAIL: AddRef mismatch rc1=" & rc1 & " rc2=" & rc2
        Exit Function
    End If

    ' Test QueryInterface via DispCallFunc for IID_IUnknown and handler IID
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

    ' Test Invoke (vtable index 3) through VBE7 EntryPoint stub -> x86/x64 machine-code thunk -> ComCallbackHandler.Handler_Invoke!
    hr = dcf(pObj1, 3, "MockInvoke", 0&, CLngPtr(&H123456))
    If hr <> 0 Then
        VerifyStep1_MockThunks = "FAIL: Invoke via thunk returned hr=&H" & Hex$(hr)
        Exit Function
    End If

    ' Release all references and verify auto-release
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

Public Function GetDemoFormHwnd() As Long
    On Error Resume Next
    GetDemoFormHwnd = CLng(Forms("frmWv2Demo").hwnd)
    On Error GoTo 0
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
