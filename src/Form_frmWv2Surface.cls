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


Private m_env As Wv2Environment
Private WithEvents m_pane As Wv2Pane
Attribute m_pane.VB_VarHelpID = -1
Private m_bridge As Object
Private m_lastWebMessage As String
Private m_lastWebMessageJson As String

Public Property Get pane() As Wv2Pane
    Set pane = m_pane
End Property

Public Property Get LastWebMessage() As String
    LastWebMessage = m_lastWebMessage
End Property

Public Property Get LastWebMessageJson() As String
    LastWebMessageJson = m_lastWebMessageJson
End Property

Public Function InitSpa( _
    ByVal virtualHostName As String, _
    ByVal localFolderPath As String, _
    ByVal entryFile As String, _
    Optional ByVal hostObject As Object = Nothing, _
    Optional ByVal hostObjectName As String = "vba", _
    Optional ByVal userDataFolder As String = "") As Boolean

    On Error GoTo ErrHandler

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

    Dim targetHwnd As LongPtr
    targetHwnd = GetAccessFormCanvasHwnd(Me)

    Set m_pane = New Wv2Pane
    m_pane.Init m_env, targetHwnd

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

    ResizeViewport
    m_pane.Ctrl_PutIsVisible True

    If m_pane.View3_SetVirtualHostNameToFolderMapping(virtualHostName, localFolderPath, 2) <> 0 Then
        Exit Function
    End If

    Dim targetUrl As String
    targetUrl = "https://" & virtualHostName & "/" & entryFile
    m_pane.View_Navigate targetUrl

    InitSpa = True
    Exit Function

ErrHandler:
    Debug.Print "frmWv2Surface.InitSpa Error " & Err.Number & ": " & Err.Description
    InitSpa = False
End Function

Public Sub ResizeViewport()
    If m_pane Is Nothing Then Exit Sub
    If Not m_pane.IsReady Then Exit Sub

    Dim rc As RECT
    Dim targetHwnd As LongPtr
    targetHwnd = GetAccessFormCanvasHwnd(Me)
    If GetClientRect(targetHwnd, rc) <> 0 Then
        m_pane.Ctrl_PutBounds rc
    End If
End Sub

Public Sub ShutdownWebView()
    On Error Resume Next
    If Not m_pane Is Nothing Then
        m_pane.Ctrl_Close
        Set m_pane = Nothing
    End If
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

Private Sub Form_Resize()
    ResizeViewport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShutdownWebView
End Sub


