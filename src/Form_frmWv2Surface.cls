Attribute VB_Name = "Form_frmWv2Surface"
Attribute VB_Base = "0{21CA795F-C866-4C36-A3BE-03B6CC8C22B3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit


Private m_host As clsWv2Host
Private m_hostHwnd As LongPtr

Private Function EnsureHost() As clsWv2Host
    If Not (m_host Is Nothing) Then
        Set EnsureHost = m_host
        Exit Function
    End If
    If m_hostHwnd = 0 Then m_hostHwnd = GetAccessFormCanvasHwnd(Me)
    If m_hostHwnd <> 0 Then
        Set m_host = GetRegisteredAccessHost(m_hostHwnd)
    End If
    Set EnsureHost = m_host
End Function

Public Property Get pane() As Wv2Pane
    Dim h As clsWv2Host
    Set h = EnsureHost()
    If Not (h Is Nothing) Then Set pane = h.pane
End Property

Public Property Get LastWebMessage() As String
    Dim h As clsWv2Host
    Set h = EnsureHost()
    If Not (h Is Nothing) Then LastWebMessage = h.LastWebMessage
End Property

Public Property Get LastWebMessageJson() As String
    Dim h As clsWv2Host
    Set h = EnsureHost()
    If Not (h Is Nothing) Then LastWebMessageJson = h.LastWebMessageJson
End Property

Public Function InitSpa( _
    ByVal virtualHostName As String, _
    ByVal localFolderPath As String, _
    ByVal entryFile As String, _
    Optional ByVal hostObject As Object = Nothing, _
    Optional ByVal hostObjectName As String = "vba", _
    Optional ByVal userDataFolder As String = "") As Boolean

    On Error GoTo ErrHandler

    m_hostHwnd = GetAccessFormCanvasHwnd(Me)
    If m_hostHwnd = 0 Then Exit Function

    Dim existing As clsWv2Host
    Set existing = GetRegisteredAccessHost(m_hostHwnd)
    If Not (existing Is Nothing) Then
        UnregisterAccessHost m_hostHwnd
        existing.Shutdown
        Set existing = Nothing
    End If

    Dim rc As RECT
    GetAccessFormViewportRect Me, m_hostHwnd, rc

    Set m_host = New clsWv2Host
    RegisterAccessHost m_hostHwnd, m_host

    If Not m_host.InitSpaOnHwnd( _
        m_hostHwnd, rc, virtualHostName, localFolderPath, entryFile, _
        hostObject, hostObjectName, userDataFolder) Then
        UnregisterAccessHost m_hostHwnd
        m_host.Shutdown
        Set m_host = Nothing
        Exit Function
    End If

    SyncAccessWebViewZOrder Me, m_hostHwnd
    InitSpa = True
    Exit Function

ErrHandler:
    Debug.Print "frmWv2Surface.InitSpa Error " & Err.Number & ": " & Err.Description
    InitSpa = False
End Function

Public Sub ResizeViewport()
    Dim h As clsWv2Host
    Set h = EnsureHost()
    If h Is Nothing Then Exit Sub

    Dim rc As RECT
    If GetAccessFormViewportRect(Me, h.HostHwnd, rc) Then
        h.UpdateBounds rc
        SyncAccessWebViewZOrder Me, h.HostHwnd
    End If
End Sub

Public Sub ShutdownWebView()
    On Error Resume Next
    Dim h As clsWv2Host
    Set h = EnsureHost()
    If Not (h Is Nothing) Then
        UnregisterAccessHost h.HostHwnd
        h.Shutdown
    End If
    Set m_host = Nothing
    m_hostHwnd = 0
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    ResizeViewport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsSubformWindow(Me) Then
        Dim h As clsWv2Host
        Set h = EnsureHost()
        If Not (h Is Nothing) Then h.ScheduleOrphanCheck
        Set m_host = Nothing
    Else
        ShutdownWebView
    End If
End Sub
