Attribute VB_Name = "clsWv2Host"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

' ============================================================================
' clsWv2Host - Persistent Access WebView2 Controller Host
'
' Survives Access Subform (OFormChild) window recreation & VBA Form reload
' during grow/maximize resizes by parenting ICoreWebView2Controller to the
' parent form's immortal section OFormSub HWND, with automatic SafeTimer
' cleanup if the parent form closes without explicit ShutdownWebView.
' ============================================================================

Private m_env As Wv2Environment
Private WithEvents m_pane As Wv2Pane
Attribute m_pane.VB_VarHelpID = -1
Private WithEvents m_cleanupTimer As SafeTimer
Attribute m_cleanupTimer.VB_VarHelpID = -1
Private m_bridge As Object
Private m_hostHwnd As LongPtr
Private m_lastWebMessage As String
Private m_lastWebMessageJson As String
Private m_isShutdown As Boolean

Public Property Get HostHwnd() As LongPtr
    HostHwnd = m_hostHwnd
End Property

Public Property Get pane() As Wv2Pane
    Set pane = m_pane
End Property

Public Property Get LastWebMessage() As String
    LastWebMessage = m_lastWebMessage
End Property

Public Property Get LastWebMessageJson() As String
    LastWebMessageJson = m_lastWebMessageJson
End Property

Public Function InitSpaOnHwnd( _
    ByVal targetHostHwnd As LongPtr, _
    ByRef initialBounds As RECT, _
    ByVal virtualHostName As String, _
    ByVal localFolderPath As String, _
    ByVal entryFile As String, _
    Optional ByVal hostObject As Object = Nothing, _
    Optional ByVal hostObjectName As String = "vba", _
    Optional ByVal userDataFolder As String = "") As Boolean

    On Error GoTo ErrHandler

    If targetHostHwnd = 0 Then Exit Function
    m_hostHwnd = targetHostHwnd
    m_isShutdown = False

    Dim uFolder As String
    uFolder = userDataFolder
    If Len(uFolder) = 0 Then uFolder = GetDefaultUserDataFolder()

    Set m_env = New Wv2Environment
    m_env.Init uFolder

    Dim t0 As Single
    t0 = Timer
    Do While Not m_env.IsReady And Not m_env.IsFailed
        DoEvents
        If Timer - t0 > 15 Then Exit Function
    Loop
    If Not m_env.IsReady Then Exit Function

    Set m_pane = New Wv2Pane
    m_pane.Init m_env, m_hostHwnd

    t0 = Timer
    Do While Not m_pane.IsReady And Not m_pane.IsFailed
        DoEvents
        If Timer - t0 > 15 Then Exit Function
    Loop
    If Not m_pane.IsReady Then Exit Function

    If m_pane.Ctrl_GetCoreWebView2() <> 0 Then Exit Function

    m_pane.View_AddNavigationStarting
    m_pane.View_AddNavigationCompleted
    m_pane.View_AddDocumentTitleChanged
    m_pane.View_AddWebMessageReceived
    m_pane.View_AddHistoryChanged
    m_pane.View2_AddDOMContentLoaded

    If Not hostObject Is Nothing Then
        Set m_bridge = hostObject
        m_pane.View_AddHostObjectToScript hostObjectName, m_bridge
    End If

    m_pane.Ctrl_PutBounds initialBounds
    m_pane.Ctrl_PutIsVisible True

    If m_pane.View3_SetVirtualHostNameToFolderMapping(virtualHostName, localFolderPath, 2) <> 0 Then
        Exit Function
    End If

    Dim targetUrl As String
    targetUrl = "https://" & virtualHostName & "/" & entryFile
    m_pane.View_Navigate targetUrl

    InitSpaOnHwnd = True
    Exit Function

ErrHandler:
    Debug.Print "clsWv2Host.InitSpaOnHwnd Error " & Err.Number & ": " & Err.Description
    InitSpaOnHwnd = False
End Function

Public Sub UpdateBounds(ByRef rc As RECT)
    If m_isShutdown Then Exit Sub
    If m_pane Is Nothing Then Exit Sub
    If Not m_pane.IsReady Then Exit Sub
    m_pane.Ctrl_PutBounds rc
    m_pane.Ctrl_NotifyParentWindowPositionChanged
End Sub

Public Sub ScheduleOrphanCheck()
    If m_isShutdown Then Exit Sub
    Set m_cleanupTimer = New SafeTimer
    m_cleanupTimer.StartTimer 80
End Sub

Private Sub m_cleanupTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    On Error Resume Next
    If Not m_cleanupTimer Is Nothing Then
        m_cleanupTimer.StartTimer 0
        Set m_cleanupTimer = Nothing
    End If
    If m_hostHwnd <> 0 Then
        If IsWindow(m_hostHwnd) = 0 Then
            UnregisterAccessHost m_hostHwnd
            Shutdown
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub Shutdown()
    If m_isShutdown Then Exit Sub
    m_isShutdown = True
    On Error Resume Next
    If Not m_cleanupTimer Is Nothing Then
        m_cleanupTimer.StartTimer 0
        Set m_cleanupTimer = Nothing
    End If
    If Not m_pane Is Nothing Then
        m_pane.Ctrl_Close
        Set m_pane = Nothing
    End If
    m_hostHwnd = 0
    Set m_bridge = Nothing
    Set m_env = Nothing
    On Error GoTo 0
End Sub

Private Sub m_pane_WebMessageReceived(ByVal messageText As String, ByVal messageJson As String)
    If Len(messageText) > 0 Then
        m_lastWebMessage = messageText
    Else
        m_lastWebMessage = messageJson
    End If
    m_lastWebMessageJson = messageJson
End Sub

Private Sub Class_Terminate()
    Shutdown
End Sub
