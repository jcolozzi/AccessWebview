Attribute VB_Name = "Wv2SettingsBridge"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
''''''''''''''''''''''''''''''''''
' --- Wv2SettingsBridge.cls  _9.29 __ (_______________ + custom ____) ---
'
'   _9.29 _____ (___________ URL ___________):
'     *___________ (__1 _a)*
'       __________________________________.
'       %s _________ (Chrome/Firefox _) _ NormalizeUrl / PreviewUrlFor /
'       _________ / ______________________,
'       ____________.
'
'     *______ (__2 _a)*
'       ____ + http:// ___ https:// ____, __. .
'       ?q= ____*____ URL*___, ___ URL .
'       _________. ___ NormalizeCustomTemplate ______.
'
'     *_____ 2 ___template __ custom _____ (__3 _a-1)*
'         engine=custom
'         template=https://example.com/search?q=
'       _________ engine=bing _ 1 ___ (_9.20 ___).
'       *__ 2 _______* ________ URL ___________,
'         ______ URL ___________. ___________,
'         URL _____________.
'       ____: custom -> ____________ template _______,
'         _____________ URL ____ (__).
'
'     - __ Public (JS __, __5 _a):
'         SetCustomEngine(template) As String
'           __ -> Browser ___ -> __. _____________.
'           ____ "" (Browser __________).
'           *___________ URL ______ "custom" .
'             ___________, ____________*
'             (TemplateToName __________. _______________)
'         GetCustomTemplate() As String
'           _______. __ custom ____ URL, .
'           custom, ____ "".
'         PreviewUrlForTemplate(template, terms) As String
'           *___*__________. Browser __________.
'           ______ JS ______ VBA ____________.
'
'     - __ Public (VBA _):
'         LoadCustomTemplate() As String  (____ custom ______. ____)
'         Debug_SaveSettings(engineName, template)  (______)
'
'     - _____:
'         LoadEngineName  -> ___ LoadValue(keyName) _____, ______.
'                           (engine / template _ 2 _____________)
'         LoadEngineTemplate -> engine=custom __ template ____ (____.
'                           ____ "" _, Wv2Browser ___ Google ______)
'         SaveEngineName  -> SaveSettingsFile(engineName, template) _______.
'                           *______ SaveSettingsFile _____*
'         _Wv2Browser ____________ or .
'           (_9.20 __Browser ____________________).
'
' --- Wv2SettingsBridge.cls  _9.20 __ (____________) ---
'
'   _9.20 _____ (__________________):
'     *__________ (APPDATA) _ __________*
'       __ Excel ________________ (_D) ______,
'       ____ Sheets(...) ___ Excel _____________, Word .
'       _________ (____________) ___. .
'       (%APPDATA%\Wv2Browser\settings.txt) _______. I/O _ VBA .
'       Open / Print / Line Input / Dir / MkDir ___, Win32 _ ADODB .
'       (______ ASCII ___ UTF-8 _____). _________.
'
'     *____________ (___________)*
'       ______________________(_______ SE_*.
'       ____________SetEngine/GetEngine) ______. .
'       ____________________, .
'       ___________. Wv2Store.bas .
'       (____________).
'
'     *_______ 1 _ key=value ____*
'       _: engine=bing . ______ 1 ____, .
'       __________________ 2 _______. .
'       '=' ____ engine ________. _______.
'
'     *__/________*
'       ___: SetEngine ___________ SaveEngineName _____.
'         ______ (_______) ______, _________.
'       ___: Wv2Browser.Class_Initialize ________ New .
'         LoadEngineTemplate() ___ (__). _____ m_searchTemplate .
'         ______, ______ (Google) ___. Browser .
'         __, UserForm ____________________.
'
'     *_____*
'       LoadEngineName()     As String  _________ ("bing" _). ____ "".
'       LoadEngineTemplate() As String  __________________.
'                                        __/____ "" (__________).
'       Debug_SaveEngineName(name)       ___: SaveEngineName ________.
'
' --- Wv2SettingsBridge.cls  _9.18 __ (__) ---
'
'   *________*
'     ____ (_____ HTML/JS) _____________________.
'     ICoreWebView2.AddHostObjectToScript __________ JS ____,
'     JS __ window.chrome.webview.hostObjects.sync.Settings.SetEngine("bing")
'     ____ VBA _________________ (postMessage .
'     _____________ RPC __).
'
'   *__ Wv2Browser ________________ (__2-b)*
'     Wv2Browser ___ AddHostObjectToScript ___, _____ JS .
'     CloseTab / MoveTab ______________ (JS _________).
'     _____________________ API ______, .
'     Wv2Browser _____ 1 ___. _JS .
'     _____. ___________________, ____________.
'
'   *__________ (DRY)*
'     "google"/"bing"/"duckduckgo"(__ "ddg")/"yahoo" _ _________ .
'     _______________. 9.17 __ UserForm1 .
'     _______ (SE_*) ________________, UserForm1 .
'     UseSearchEngine _____________. _____ 1 _____,
'     ________________________.
'
'   *Instancing ____*
'     _ WebView2 __ (______) __ Instancing = 1 - Private .
'     AddHostObjectToScript _____ (VBA _____ IDispatch _ OS .
'     _______). ___________ Private ______.
'     PublicNotCreatable ___ (__________________).
'
'   *______ (JS ______)*
'     GetEngine() As String          ___________ ("google" _).
'                                     ________________ "custom".
'     SetEngine(name) As String      ___________ Browser ___.
'                                     ________________.
'     PreviewUrl(terms) As String    ___________________ URL.
'                                     ______________.
'     EngineList() As String         ________ JSON _____.
'                                     ________________.
'
'   *VBA->JS ____*
'     ____________, ____________ JS .
'     _______ (hostObjects.sync __________).
' ============================================================
Option Explicit

' --- ____________ (_9.17 _ UserForm1 _________) ---
'   __________________. __________________.
'   DuckDuckGo _ /?q= __Yahoo _ ?p= .
'   1 _______ (_9.17 _____).
Private Const SE_GOOGLE     As String = "https://www.google.com/search?q="
Private Const SE_BING       As String = "https://www.bing.com/search?q="
Private Const SE_DUCKDUCKGO As String = "https://duckduckgo.com/?q="
Private Const SE_YAHOO      As String = "https://search.yahoo.com/search?p="

' --- _9.20 __: ___ (____________) ---
'   %APPDATA%\Wv2Browser\settings.txt _ 1 _ key=value ___.
'   I/O _ VBA ______ (Win32/ADODB ___). Word _________.
Private Const STORE_SUBFOLDER As String = "Wv2Browser"
Private Const STORE_FILENAME  As String = "settings.txt"
Private Const STORE_KEY_ENGINE As String = "engine"
' _9.29: custom _______ 2 ___ (______ URL ____)
Private Const STORE_KEY_TEMPLATE As String = "template"

' --- _____ Browser __ (________) ---
Private m_browser As Wv2Browser


' ============================================================
' BindBrowser (VBA _____. JS ________)
'
'   __________ Wv2Browser ______. .
'   AddHostObjectToScript _____ Browser ___.
'   *______* ____->Browser ________. Browser .
'     WithEvents ____, ____ (Pane) ___ hostObject ___ WebView2 .
'     ____. _________ WebView2 _______, __________.
' ============================================================
Public Sub BindBrowser(ByVal br As Wv2Browser)
    Set m_browser = br
End Sub


' ============================================================
' ResolveTemplate (______: __ -> ______)
'
'   _______________________. .
'   __. ____________ (__________________).
' ============================================================
Private Function ResolveTemplate(ByVal engineName As String) As String
    Dim key As String
    key = LCase$(Trim$(engineName))
    Select Case key
        Case "google":            ResolveTemplate = SE_GOOGLE
        Case "bing":              ResolveTemplate = SE_BING
        Case "duckduckgo", "ddg": ResolveTemplate = SE_DUCKDUCKGO
        Case "yahoo":             ResolveTemplate = SE_YAHOO
        Case Else:                ResolveTemplate = ""
    End Select
End Function


' ============================================================
' TemplateToName (______: ______ -> __)
'
'   ________________________________.
'   ___________ "custom" ___ (.
'   ______).
' ============================================================
Private Function TemplateToName(ByVal template As String) As String
    Select Case template
        Case SE_GOOGLE:     TemplateToName = "google"
        Case SE_BING:       TemplateToName = "bing"
        Case SE_DUCKDUCKGO: TemplateToName = "duckduckgo"
        Case SE_YAHOO:      TemplateToName = "yahoo"
        Case Else:          TemplateToName = "custom"
    End Select
End Function


' ============================================================
' GetEngine (JS __: ________)
'
'   __ Browser ___________________________.
'   ____________________________________.
'   Browser ________.
' ============================================================
Public Function GetEngine() As String
    If m_browser Is Nothing Then
        GetEngine = ""
        Exit Function
    End If
    GetEngine = TemplateToName(m_browser.Debug_SearchEngine())
End Function


' ============================================================
' SetEngine (JS __: _______)
'
'   ___________ Browser.SearchEngine _____. .
'   __ (Google) ____. ___/______________ Browser .
'   Property Let _______, _____________ (_________).
'   ___________________(__________________).
' ============================================================
Public Function SetEngine(ByVal engineName As String) As String
    If m_browser Is Nothing Then
        Debug.Print "Wv2SettingsBridge.SetEngine: Browser ________ (" & engineName & ")"
        SetEngine = ""
        Exit Function
    End If

    Dim template As String
    template = ResolveTemplate(engineName)

    If Len(template) = 0 Then
        Debug.Print "Wv2SettingsBridge.SetEngine: ________ '" & engineName & _
                    "'. __ (Google) ____. "
        template = SE_GOOGLE
    End If

    m_browser.SearchEngine = template
    SetEngine = TemplateToName(m_browser.Debug_SearchEngine())
    Debug.Print "Wv2SettingsBridge.SetEngine: '" & engineName & "' -> '" & SetEngine & "'"

    ' _9.20: __________________ (_________).
    '   __________ (_________). custom .
    '   (_________ UI ________. _____________).
    If SetEngine <> "custom" And Len(SetEngine) > 0 Then
        SaveEngineName SetEngine
    End If
End Function


' ============================================================
' NormalizeCustomTemplate (______: _____________, _9.29)
'
'   ________ 2 ___ (__2 _a):
'     _____________ 0 .
'     _http:// ___ https:// ____ (_________)
'   ___ Trim _______, ______ "" ___.
'   *__________* _________ ?q= ________ URL ___,
'     ____ URL ____________________.
' ============================================================
Private Function NormalizeCustomTemplate(ByVal template As String) As String
    NormalizeCustomTemplate = ""

    Dim t As String
    t = Trim$(template)
    If Len(t) = 0 Then Exit Function

    Dim lo As String
    lo = LCase$(t)
    If Left$(lo, 7) = "http://" Or Left$(lo, 8) = "https://" Then
        NormalizeCustomTemplate = t
    End If
End Function


' ============================================================
' SetCustomEngine (JS __: __________________, _9.29)
'
'   _______________________.
'     ___ "custom"      : .
'     ___ "bing" _     : ___________ URL _____, .
'                            _______ (_______________)
'     ___ ""            : _____. *Browser ________________*
'
'   *___________* VBA _______________, .
'     __________ (nm) _________________.
' ============================================================
Public Function SetCustomEngine(ByVal template As String) As String
    SetCustomEngine = ""

    If m_browser Is Nothing Then
        Debug.Print "Wv2SettingsBridge.SetCustomEngine: Browser ________"
        Exit Function
    End If

    Dim t As String
    t = NormalizeCustomTemplate(template)
    If Len(t) = 0 Then
        Debug.Print "Wv2SettingsBridge.SetCustomEngine: ______________ ('" & _
                    template & "')"
        Exit Function
    End If

    m_browser.SearchEngine = t

    Dim nm As String
    nm = TemplateToName(m_browser.Debug_SearchEngine())

    If nm = "custom" Then
        SaveSettingsFile "custom", t
    Else
        ' ___________ URL _____. custom ________,
        ' ____________ (template ______ = __3 _a-1).
        SaveSettingsFile nm, ""
        Debug.Print "Wv2SettingsBridge.SetCustomEngine: ________________ '" & _
                    nm & "' _____"
    End If

    SetCustomEngine = nm
    Debug.Print "Wv2SettingsBridge.SetCustomEngine: __ '" & t & "' -> '" & nm & "'"
End Function


' ============================================================
' GetCustomTemplate (JS __: ___________, _9.29)
'
'   ____:
'     1. __ Browser _____________ custom .
'     2. ____________ custom .
'     3. ________ ""
'   Browser _____ 2 _____, ________________.
' ============================================================
Public Function GetCustomTemplate() As String
    Dim t As String
    t = ""

    If Not m_browser Is Nothing Then
        Dim cur As String
        cur = m_browser.Debug_SearchEngine()
        If TemplateToName(cur) = "custom" Then t = cur
    End If

    If Len(t) = 0 Then t = LoadCustomTemplate()

    GetCustomTemplate = t
End Function


' ============================================================
' PreviewUrl (JS __: __ URL ______)
'
'   ______________________________ URL ___.
'   _______________________ URL ________,
'   _________ (hostObjects.sync __________________).
'   Browser _ NormalizeUrl _____, _________________.
' ============================================================
Public Function PreviewUrl(ByVal terms As String) As String
    If m_browser Is Nothing Then
        PreviewUrl = ""
        Exit Function
    End If
    PreviewUrl = m_browser.Debug_NormalizeUrl(terms)
End Function


' ============================================================
' PreviewUrlFor (JS __: __________ URL _____)
'
'   _____________, ___________________ URL .
'   ______ (_______________________). .
'   __________ URL ___________.
'   Browser __________, ______________ (.
'   Browser ___________).
' ============================================================
Public Function PreviewUrlFor(ByVal engineName As String, ByVal terms As String) As String
    If m_browser Is Nothing Then
        PreviewUrlFor = ""
        Exit Function
    End If

    Dim template As String
    template = ResolveTemplate(engineName)
    If Len(template) = 0 Then template = SE_GOOGLE

    PreviewUrlFor = template & Wv2Url.UrlEncodeComponent(terms)
End Function


' ============================================================
' PreviewUrlForTemplate (JS __: _______________, _9.29)
'
'   _____________*__*_____, _____ URL ______.
'   ___________ Browser ___________ (_______).
'   ____________ "" ___ (_________________).
'   *______ Wv2Url.UrlEncodeComponent* JS __ encodeURIComponent .
'     _______ (BuildSearchUrl) __________, __ VBA ___.
' ============================================================
Public Function PreviewUrlForTemplate(ByVal template As String, ByVal terms As String) As String
    PreviewUrlForTemplate = ""

    Dim t As String
    t = NormalizeCustomTemplate(template)
    If Len(t) = 0 Then Exit Function

    PreviewUrlForTemplate = t & Wv2Url.UrlEncodeComponent(terms)
End Function


' ============================================================
' EngineList (JS __: ________ JSON ___)
'
'   ____________________ JSON _____.
'   ____ {"id":__, "label":___, "template":______}.
'   ______________ 1 _______ 2 .
'   ________ (HTML ______________).
'
'   *JSON ___________* _______ URL _______,
'     _________________, Wv2Json ___________.
' ============================================================
Public Function EngineList() As String
    Dim s As String
    s = "["
    s = s & "{""id"":""google"",""label"":""Google"",""template"":""" & SE_GOOGLE & """},"
    s = s & "{""id"":""bing"",""label"":""Bing"",""template"":""" & SE_BING & """},"
    s = s & "{""id"":""duckduckgo"",""label"":""DuckDuckGo"",""template"":""" & SE_DUCKDUCKGO & """},"
    s = s & "{""id"":""yahoo"",""label"":""Yahoo"",""template"":""" & SE_YAHOO & """}"
    s = s & "]"
    EngineList = s
End Function


' ============================================================
' SettingsPath (__: _________________)
'
'   %APPDATA%\Wv2Browser\settings.txt ___. APPDATA .
'   ______, ______/_______________.
' ============================================================
Private Function SettingsPath() As String
    Dim base As String
    base = Environ$("APPDATA")
    If Len(base) = 0 Then
        SettingsPath = ""
        Exit Function
    End If
    SettingsPath = base & "\" & STORE_SUBFOLDER & "\" & STORE_FILENAME
End Function


' ============================================================
' EnsureFolder (__: _____________)
'
'   %APPDATA%\Wv2Browser\ _____. _________. .
'   ________, ____ On Error Resume Next ______.
' ============================================================
Private Sub EnsureFolder()
    Dim base As String
    base = Environ$("APPDATA")
    If Len(base) = 0 Then Exit Sub

    Dim folder As String
    folder = base & "\" & STORE_SUBFOLDER

    On Error Resume Next
    If Len(Dir$(folder, vbDirectory)) = 0 Then
        MkDir folder
    End If
    On Error GoTo 0
End Sub


' ============================================================
' LoadEngineName / LoadCustomTemplate / LoadValue (__: ______)
'
'   *_9.29 _ LoadValue(keyName) ______* engine _ template _ 2 .
'   __________. LoadEngineName __________.
'
'   settings.txt _ 1 _____, <key>=<value> _ <value> ___.
'   *__ = _______* ___ = ___________,
'   template=https://example.com/search?q= _____________.
'   ____/_____________________ "" .
'   (________________). ______ (________) .
'   _________ (LoadEngineTemplate ___).
' ============================================================
Public Function LoadEngineName() As String
    LoadEngineName = LoadValue(STORE_KEY_ENGINE)
End Function


' --- _9.29: _______________ (_________ "" ____) ---
Public Function LoadCustomTemplate() As String
    LoadCustomTemplate = NormalizeCustomTemplate(LoadValue(STORE_KEY_TEMPLATE))
End Function


' --- _9.29: settings.txt __ 1 ____________ ---
Private Function LoadValue(ByVal keyName As String) As String
    LoadValue = ""

    Dim path As String
    path = SettingsPath()
    If Len(path) = 0 Then Exit Function

    On Error GoTo done
    If Len(Dir$(path)) = 0 Then Exit Function   ' .

    Dim fno As Integer
    fno = FreeFile
    Open path For Input As #fno

    Dim ln As String, eqPos As Long, k As String, v As String
    Do Until EOF(fno)
        Line Input #fno, ln
        eqPos = InStr(ln, "=")
        If eqPos > 0 Then
            k = LCase$(Trim$(Left$(ln, eqPos - 1)))
            v = Trim$(Mid$(ln, eqPos + 1))
            If k = LCase$(keyName) Then
                LoadValue = v
                ' ___________, ____________ break ___.
            End If
        End If
    Loop

    Close #fno
    Exit Function

done:
    ' ______: ________, "" ___.
    On Error Resume Next
    Close #fno
    On Error GoTo 0
    LoadValue = ""
End Function


' ============================================================
' LoadEngineTemplate (__: _____________________)
'
'   LoadEngineName ______, ResolveTemplate ______________.
'   _____/_____ (custom _____) ____ "" ___, .
'   (Wv2Browser.Class_Initialize) ___ (Google) ____________.
'   *_______________________, Browser .
'     ___ or _______* Browser ________________.
' ============================================================
Public Function LoadEngineTemplate() As String
    Dim name As String
    name = LoadEngineName()
    If Len(name) = 0 Then
        LoadEngineTemplate = ""
        Exit Function
    End If

    ' _9.29: custom __ template ____ (NormalizeCustomTemplate __).
    '   ___ template _____________ "" ___, .
    '   (Wv2Browser.Class_Initialize) ___ (Google) ______.
    If LCase$(Trim$(name)) = "custom" Then
        LoadEngineTemplate = LoadCustomTemplate()
        Exit Function
    End If

    LoadEngineTemplate = ResolveTemplate(name)   ' ____ "" .
End Function


' ============================================================
' SaveEngineName (__: __________)
'
'   settings.txt _ engine=<name> _ 1 ____ (__ 1 ___________).
'   _________ EnsureFolder ___. __________ (.
'   __________). SetEngine _________.
'   *__ 2 ___________, .
'     ____. ______ 1 ______________.
' ============================================================
Private Sub SaveEngineName(ByVal engineName As String)
    SaveSettingsFile engineName, ""
End Sub


' ============================================================
' SaveSettingsFile (__: settings.txt _____, _9.29)
'
'   engine=<n> _____, template ________ template=<URL> .
'   (__3 _a-1). _________ template _ "" _____ 1 ______,
'   *custom __ ____________ template ________*.
' ============================================================
Private Sub SaveSettingsFile(ByVal engineName As String, ByVal template As String)
    Dim path As String
    path = SettingsPath()
    If Len(path) = 0 Then Exit Sub

    EnsureFolder

    On Error GoTo fail
    Dim fno As Integer
    fno = FreeFile
    Open path For Output As #fno
    Print #fno, STORE_KEY_ENGINE & "=" & engineName
    If Len(template) > 0 Then
        Print #fno, STORE_KEY_TEMPLATE & "=" & template
    End If
    Close #fno
    Debug.Print "Wv2SettingsBridge.SaveSettingsFile: __ '" & engineName & "'" & _
                IIf(Len(template) > 0, " / template=" & template, "") & " -> " & path
    Exit Sub

fail:
    On Error Resume Next
    Close #fno
    On Error GoTo 0
    Debug.Print "Wv2SettingsBridge.SaveSettingsFile: ____ (_____) '" & engineName & "'"
End Sub


' ============================================================
' Debug_SaveEngineName (___: SaveEngineName ______)
'
'   SaveEngineName _ Private ___, _______ (Test_9_20_Persistence) .
'   __->________________________. __________.
' ============================================================
Public Sub Debug_SaveEngineName(ByVal engineName As String)
    SaveEngineName engineName
End Sub


' ============================================================
' Debug_SaveSettings (___: SaveSettingsFile ______, _9.29)
'
'   Test_9_29_Custom / Test_9_20_Persistence __engine + template .
'   ______________. __________.
' ============================================================
Public Sub Debug_SaveSettings(ByVal engineName As String, ByVal template As String)
    SaveSettingsFile engineName, template
End Sub


