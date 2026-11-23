Attribute VB_Name = "Form_frmWv2Demo"
Attribute VB_Base = "0{A81D0D81-FC73-4BDA-8D1B-9DB7B29AA75F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit


Private m_bridge As clsWv2AccessBridge

Private Function GetSurface() As Form_frmWv2Surface
    On Error Resume Next
    Set GetSurface = Me.subWebView.Form
    On Error GoTo 0
End Function

Public Property Get ActivePane() As Wv2Pane
    Dim surf As Form_frmWv2Surface
    Set surf = GetSurface()
    If Not (surf Is Nothing) Then Set ActivePane = surf.pane
End Property

Public Property Get LastWebMessage() As String
    Dim surf As Form_frmWv2Surface
    Set surf = GetSurface()
    If Not (surf Is Nothing) Then LastWebMessage = surf.LastWebMessage
End Property

Public Sub RefreshHeaderAndViewport()
    On Error Resume Next
    If Me.InsideWidth > 4200 Then
        Me.cmdOpenDevTools.Left = Me.InsideWidth - Me.cmdOpenDevTools.Width - 180
        Me.cmdPushToJs.Left = Me.cmdOpenDevTools.Left - Me.cmdPushToJs.Width - 120
        Me.lblBanner.Width = Me.cmdPushToJs.Left - 360
        If Me.Width < Me.InsideWidth Then Me.Width = Me.InsideWidth
        Me.subWebView.Left = 0
        Me.subWebView.Top = 0
        Me.subWebView.Width = Me.InsideWidth
    End If
    If Me.InsideHeight > 1200 Then
        Dim targetH As Long
        targetH = Me.InsideHeight - Me.FormHeader.Height
        If targetH > 0 Then
            If targetH > Me.Detail.Height Then
                Me.Detail.Height = targetH
                Me.subWebView.Height = targetH
            Else
                Me.subWebView.Height = targetH
                Me.Detail.Height = targetH
            End If
        End If
    End If
    Dim surf As Form_frmWv2Surface
    Set surf = GetSurface()
    If Not (surf Is Nothing) Then surf.ResizeViewport
    On Error GoTo 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.InsideWidth = 14400
    Me.InsideHeight = 9600
    On Error GoTo 0
    Set m_bridge = New clsWv2AccessBridge
    Dim surf As Form_frmWv2Surface
    Set surf = GetSurface()
    If surf Is Nothing Then
        Me.lblBanner.Caption = "WebView2 Subform Not Loaded"
        Exit Sub
    End If

    Dim ok As Boolean
    ok = surf.InitSpa( _
        "appassets.access.local", _
        GetDefaultSpaFolder(), _
        "index.html", _
        m_bridge, _
        "vba")
    If Not ok Then
        Me.lblBanner.Caption = "WebView2 Initialization Failed - Check Immediate Window"
    Else
        RefreshHeaderAndViewport
    End If
End Sub

Private Sub Form_Resize()
    RefreshHeaderAndViewport
End Sub

Private Sub cmdPushToJs_Click()
    Dim p As Wv2Pane
    Set p = Me.ActivePane
    If Not p Is Nothing Then
        p.View_PostWebMessageAsString "Hello from Native Access CommandButton @ " & Format$(Now, "hh:nn:ss")
    End If
End Sub

Private Sub cmdOpenDevTools_Click()
    Dim p As Wv2Pane
    Set p = Me.ActivePane
    If Not p Is Nothing Then
        p.Settings_PutAreDevToolsEnabled True
        p.View_OpenDevToolsWindow
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim surf As Form_frmWv2Surface
    Set surf = GetSurface()
    If Not (surf Is Nothing) Then surf.ShutdownWebView
    ShutdownAccessFormWebViews Me
    Set m_bridge = Nothing
    On Error GoTo 0
End Sub
