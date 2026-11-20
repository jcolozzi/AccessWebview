Attribute VB_Name = "Wv2Browser"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
''''''''''''''''''''''''''''''''''
' --- Wv2Browser.cls  K-4c __ (____________) ---
'
'   *__ (K-4a/b)* ____________ COM __________.
'
'     _____                   ... __ 5.05 _ (COM __ 1 __ 300_500ms)
'     about:blank ____ (____) ... __ 0.38 _ (__ 8_16ms)
'     about:blank ____ (3 ___) ... __ 0.37 _  <- *_________*
'
'   *_______* _________. ____ 5.05 _ -> 1.98 _.
'   _____ 1.6 ______, ________ COM _____.
'
'   *TabBar / NavBar _ Pane ____* __ HTML ___, ____ 23_31ms.
'
'   *QuietOnShutdown = False ____* _______ (Test_K4_Step) .
'   ________________, _________.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Browser.cls  K-4a __ (________ Wv2Log ___) ---
'
'   *_____* ________________ (1 ____ 0.5 _____)
'   __________, _____ Debug.Print ____________,
'   ____________________.
'
'   Wv2Log _________________, *______________*.
'   ___ Wv2Log.LogEcho = False / True _ 2 ____,
'   *_COM ___________________________* _____.
'
'   *__________. ____. * (D-7 ____ 3 _________)
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Browser.cls  K-2 __ (________________) ---
'
'   *__* _____________________________,
'     ________________________. .
'     _____________________________.
'
'   *__* IsSettings __________, _____ False _______.
'     ________________, OpenSettingsTab .
'     ______________.
'     *_______________* ____________JS .
'     ________. ___ NavigateActive _____________.
'
'   *__ (A-2)* ________ *________* ______.
'     _____ HTML _ window.__wv2Settings = true ___, OpenSettingsTab .
'     ____ EvalSync (D-1) .
'     _____. ___________________, .
'     _______________.
'     *______, _______________ (____91)*
'
'   *__ (B-1)* _________, _____ OpenSettingsTab .
'     ______ 2 __________. _______________,
'     *_________________________________*
'     (____60 _________).
'     *____27 _____________ SafeTimer _____*
'     SetEngine _ hostObjects __ (JS ______) _______,
'     ____ EvalSync _________.
'
'   *IsSettings __________, ______________*
'     ______________________________.
''''''''''''''''''''''''''''''''''
' --- Wv2Browser.cls  _9.29 __ (___________ URL ___ + _______) ---
'
'   _9.29 _____ (___ Wv2SettingsBridge _. ___ UI ____):
'
'   *(A) BuildSettingsHtml ___________ (__4 _a)*
'     ________*_________*_____. 5 .
'     __, EngineList() .
'     (_9.18 __HTML ______________________).
'       ____ + [__] ___. custom _____________________.
'       _____ Settings.PreviewUrlForTemplate _*___*_________.
'         ______ JS _____ VBA ____________.
'       ______ (http:// / https:// ______) __________,
'         _____ Browser_______________.
'     *<form> _____* Enter _ submit _________________.
'       <input> + <button> _ onclick / onkeydown _____.
'     *JS _________________* VBA _____________+.
'       (VBA _ " __________). ___ HTML ___ VBA _ "" ___.
'
'   *(B) EnsurePrewarmScheduled __________ (_9.28b ___ 1)*
'     *__________ (_____m_prewarmClaimed __________)*
'       m_adoptTimer_TimerCall _ AdoptPrewarmedNow -> CreatePrewarmPaneNow .
'       __. ___ CreatePaneCore _ WaitPaneReady _ DoEvents _____,
'       _________________ AddTab _*____*____. .
'       AddTab ___ EnsurePrewarmScheduled _____ m_prewarmPane .
'       Nothing (______) _________, 1.5 .
'       _______________ (_______)______.
'       -> m_prewarmClaimed ________ False __________.
'         ______*_______*= m_creatingPane > 0 ___.
'     *______* ______________, __ AddTab .
'       ________ AddTab / ______________. .
'       (____ RequestPrewarmAsync ___) _, .
'       1.5 _____________________.
'
' --- Wv2Browser.cls  _9.28 __ (____________ + _________ 2 _) ---
'
'   _9.28 _____ (9.27 _____ / ________ 3 _______. ____):
'
'   *(A) NewWindowMode _____ 0 -> 2 ___*
'     9.27 ___ (put_NewWindow ___________) _________, .
'     0 (____) ______. __________ POST ____ 405 .
'     9.27 ____3 ________, *POST ________________*.
'       ____ PREWARM_DELAY_MS (1500ms) ______________________.
'         ________. .
'         _____, _________*___________.
'       __________. 1500ms _ 9.27 __1 __Init ___ Controller .
'         __________ ____ 20 ______________________,
'         __________________.
'       ____: __ WebView2 _ 1 ________ (msedgewebview2.exe _ 1 ___).
'         ______________ NewWindowMode = 0 ____________.
'
'   *(B) m_creatingPane _ Boolean -> Long (____) ___*
'     9.27 ____3 _ CreatePaneCore ____________ (.
'     WaitPaneReady _ DoEvents _________ m_newTabTimer_TimerCall .
'     AddTab -> CreatePaneCore ____). _________, Boolean .
'     *_________________* __, ________________.
'     __/_______1 ________ > 0______.
'
'   *(C) RequestPrewarmAsync ___ return __________*
'     ___________Claim __ (____)_______________,
'     9.27 ______________________________. .
'     ________. *_______* (__________).
'
'   *______* put_NewWindow _______ / .
'     ____ / Deferral (_________) / _____.
'     Wv2Pane.cls _ 9.27 ______ (______ Browser _________,
'     _________________).
'
' --- Wv2Browser.cls  _9.27 __ (__: put_NewWindow ___________ / ________) ---
'
'   _9.27 _____:
'     ____ 23 (9.26b) __________________ POST .
'     ____. _____ UX __ __________________, *_____* .
'     __________. args.put_NewWindow (vtable index 4) .
'     CoreWebView2_____, ________ WebView _______________.
'
'   *______ = ______ (_B / __ 1_7 ____)*
'     __ Pane ___ 1 ______, NewWindowRequested .
'     put_NewWindow(___ ViewPtr) _____ Exit __.
'     -> *Deferral (GetDeferral / Complete) _______* _______.
'       Deferral ___________ args / deferral _ COM .
'       ___ AddRef/Release _____, ____ 10 (Sentinel MEM_FREE) _____.
'       __________________ vtable __ 1_2 ____, _____.
'       ____ 14 (SafeTimer ______) __ 20 (EBMode __) ______.
'
'   *______*
'     [___]
'       _m_newWindowMode           : 0=__(GET___) / 1=________ / 2=.
'       _m_setHandledWithNewWindow : put_NewWindow _ put_Handled ______ (__ True)
'       _m_prewarmPane / m_prewarmClaimed : .
'       _m_prewarmTimer / m_prewarmPending : _____________ SafeTimer
'       _m_adoptTimer / m_adoptPending / m_adoptUrl : ___________ SafeTimer
'       _m_creatingPane            : Pane ______ (______ TimerCall ______)
'     [____ = __2 _a]
'       _CreatePaneCore / CreatePaneCoreBody : _ AddTab ___ (Pane _________).
'         AddTab ________________. AddTab __CreatePaneCore + __ 3 .
'         (m_tabs.Add / TabAdded / ActivateTab)____________.
'     [______]
'       _NewWindowMode (Get/Let)         : _____ (3 ___ = __4 _b)
'       _PopupDelegateMode (Get/Let)     : *_____* ___ NewWindowMode ___.
'                                           Test_9_26_Popup_On/_Off _______.
'       _SetHandledWithNewWindow (Get/Let): __5 .
'       _ClaimPrewarmViewPtr             : ___ ViewPtr .
'       _ReleasePrewarmClaim             : put_NewWindow .
'       _RequestPrewarmAsync             : _________ (______ Timer)
'       _RequestAdoptPrewarmedAsync      : _______ (______ Timer)
'       _Debug_PrewarmStatus             : .
'
'   *_____*
'     [__] Init __ -> RequestPrewarmAsync (1500ms _) -> CreatePrewarmPaneNow
'            -> __ Pane ____ Ctrl_PutIsVisible False ______ (__3)
'     [__] Wv2Pane.View_OnNewWindowRequested (___ 2)
'            -> ClaimPrewarmViewPtr ______ -> put_NewWindow -> put_Handled
'            -> RequestAdoptPrewarmedAsync -> _ Exit Sub (*____________*)
'     [__] m_adoptTimer_TimerCall (_______)
'            -> AdoptPrewarmedNow : m_tabs.Add -> RaiseEvent TabAdded -> ActivateTab
'            -> CreatePrewarmPaneNow : _______ (__6: ____, ____)
'
'   *_________* ___ / _ Ready / ____ / put_NewWindow __ .
'     Wv2Pane _____ (put_Handled + RequestNewTabAsync = GET ___) .
'     __. POST _____ 405 ____, *________*. .
'     __ *_________* __, _________.
'
'   *________* _ AddTab ____ CreatePaneCore ______,
'     "Wv2Browser.AddTab: Controller __ OK" .
'     "Wv2Browser.CreatePaneCore: ..." ____. _______________.
'
'   *______* RequestNewTabAsync / m_newTabTimer / _____ / ____ 5 _ /
'     _______ / ActivateTab / CloseTab / MoveTab.
'
' --- Wv2Browser.cls  _9.26 __ (NewWindowRequested __________) ---
'
'   _9.26 _____ (POST __________ 405 _________________):
'     *__* _________ (Apache Tomcat/7.0.76) _, .
'       _____ HTTP 405 (Request method 'GET' not supported) ___. .
'       ___________, ___ Wv2Pane.View_OnNewWindowRequested .
'         (1) args.get_Uri _ URL .
'         (2) args.put_Handled(TRUE) .
'         (3) ____ Navigate (= __ GET) .
'       ________, __ POST _____/___/Referer _ (2) ________.
'
'     *________ = _______* put_Handled _____________,
'       WebView2 ________________________ (_9.6a ___).
'       __________________________ POST __________,
'       __ (args.put_NewWindow _____ CoreWebView2 _____) .
'       ___. __ 405 ________ POST _____ (Referer/Cookie/_____) .
'       _______, ________________.
'       *________ Deferral ____ (________) _____________*
'
'     *_____________* ____________ UI ____________,
'       ____ Controller _______________________. .
'       __ (_, ________________) _. ___ OFF.
'
'   *______ (Browser ________ = ________)*
'     _Private m_popupDelegateMode As Boolean        (___________)
'     _Public Property Get/Let PopupDelegateMode     (__ False = _______)
'     ____ Pane ______ Browser ______, ____ Wv2Pane .
'     m_browser.PopupDelegateMode ____________. ____________ 1 _.
'
'   *_____________* ___ UserForm1.CurrentBrowser (_9.17 ___) .
'     ____ UserForm1.frm _______. ______ Wv2Tests.bas .
'     Test_9_26_Popup_On / _Off / _Help ______.
'
'   *______* AddTab / RequestNewTabAsync / SafeTimer ____ (____ 63) /
'     ____ 5 _ / _______. _______________ (________).
'
' --- Wv2Browser.cls  _9.20 __ (________________) ---
'
'   _9.20 _____ (________________):
'     *Class_Initialize ______________ (__)*
'       Browser ________, ____ Wv2SettingsBridge _ New .
'       LoadEngineTemplate() ___, .
'       _____________ m_searchTemplate _________.
'       _________ (Google) ___. UserForm .
'       Browser __________, _________________.
'
'     *____________ Bridge ______*
'       Browser ______ ("bing" _) ________. Bridge .
'       LoadEngineTemplate() ___________ or _______,
'       Browser ____ m_searchTemplate ______. ___ (SetEngine .
'       ___ SaveEngineName) _ Bridge ______ (_9.20 ________).
'
'     *_____* Class_Initialize _ New Wv2SettingsBridge ___, .
'       OpenSettingsTab _______________ (.
'       __________). LoadEngineTemplate .
'       BindBrowser ___ (Browser _____________).
'
' --- Wv2Browser.cls  _9.19 __ ---
'
'   _9.19 _____ (_________ + ______________):
'     *____ hover _______________________*
'       BuildSettingsHtml _ JS _ showCurrentPreview() ____, (a) .
'       _____ (_ hover) _ (b) .
'       _, ______________________. ______ hover
'       ___________________.
'       *___________ mouseenter/mouseleave + hoverId ______*
'       ___________ mouseleave __, .
'       ________________, ______ (.
'       _________). _____ hoverId ___, mouseleave .
'       setTimeout(0) ______________ hover .
'       ______, .
'       __ (_____ mouseenter _________ hoverId .
'       ____________).
'
'     *OpenSettingsTab ______________*
'       9.18 __ [__] _________________, .
'       ___________. 9.19 __ OpenSettingsTab ____ m_tabs .
'       ______, Pane.IsSettings = True _________, .
'       ActivateTab _____ Exit Sub __ (_______). .
'       ______, ____ Pane _ IsSettings = True ____.
'
'     *___ Wv2Pane.IsSettings ___ (_A, ___ 9.19 ______)*
'       _________ = _______ / _____, .
'       ________. Browser ____ Pane .
'       ________. m_settingsBridge .
'       _____ 1 _____ (_____ attach __________).
'
' --- Wv2Browser.cls  _9.18 __ ---
'
'   _9.18 _____ (____ + AddHostObjectToScript _________ UI):
'     *____________* (__3)
'       TabBar _ [__] ___ -> OnPaneWebMessage "settings" -> RequestSettingsTabAsync.
'       __________ AddTab -> WaitPaneReady _____, newtab .
'       WebMessage ______________________ (____ 14).
'       _________ SafeTimer (m_settingsTimer) _______.
'       _RequestSettingsTabAsync(): ________ New SafeTimer .
'       _m_settingsTimer_TimerCall: ________ OpenSettingsTab .
'       _OpenSettingsTab(): AddTab -> ____ attach -> NavigateToString -> ActivateTab
'       _BuildSettingsHtml(): ___________ UI _ HTML .
'
'     *AddHostObjectToScript _ Wv2SettingsBridge _ JS ___* (__2-b)
'       _____ Pane _ Wv2SettingsBridge _ attach _, _____ JS .
'       hostObjects.sync.Settings.SetEngine("bing") ____ VBA _______.
'       Wv2Browser ___________________JS .
'       ___________ (______ CloseTab ___________).
'       __________________________ (9.17 __ UserForm1 .
'       _____ SE_* __________. DRY).
'
'     *attach _ NavigateToString __ (__A) __________*
'       ____________ JS _____ hostObjects ____, _____.
'       ____ ready _______ (__B) __________.
'
'     *______* attach _______ m_settingsBridge _ 1 ____,
'       ___________ WebView2 _______. Shutdown _____.
'
'   _9.16 _____ (________):
'     *__ URL _____________ (_1-A + _2-A)*
'       9.14 _ BuildSearchUrl ___________ Google _____ URL _,
'       _______ m_searchTemplate ______________. .
'       Google ___ (_____ 9.14 ________).
'
'     *___________________* (_1-A)
'       "https://www.google.com/search?q=" ____, ______ (______) .
'       _______________. Google / Bing / DuckDuckGo / Yahoo .
'       _________________. Enum + ___ (_1-B) .
'       _____, ______________ (YAGNI). __ "{q}" .
'       ______________, _____________________.
'
'     *_______ Property __, UI _____ (_2-A)*
'       Public Property Get/Let SearchEngine _______. .
'       ___ Google __________________________ (URL .
'       _________. _______________________). UI
'       _____ (________) ___________ (.
'       ______________________).
'
'     *Debug_SearchEngine ___ (_______)*
'       ___ m_searchTemplate __________. Wv2Thunks.Test_9_16_SearchEngine
'       _____________ BuildSearchUrl ___________ WebView2
'       ______________ (Debug_NormalizeUrl _______).
'
'   _9.16 _____ (________):
'     *URL ____ Wv2Url.bas ______ (_2-A / _3-A)*
'       IsAlphaOnly / UrlEncodeComponent / Utf8Bytes _ 3 _ (.
'       ___________) _ Wv2Url.bas _____. Wv2Json.bas .
'       ___________________. NormalizeUrl / BuildSearchUrl .
'       ________ (m_searchTemplate) .
'       __________ (____)____ Wv2Browser ___, .
'       Wv2Url.IsAlphaOnly / Wv2Url.UrlEncodeComponent ___ (.
'       ________). ________, _______________.
'
'     *________ Sub _ Wv2Tests.bas _____ (_1-A)*
'       Test_9_14/15/16 _ (Wv2Browser _ Debug_* _____________) .
'       Wv2Thunks.bas __ Wv2Tests.bas ____. Wv2Thunks _________.
'       Wv2Browser _____ (Debug_*) ___________ (____).
'
' --- Wv2Browser.cls  _9.15 __ ---
'
'   _9.15 _____ (___________):
'     *MoveTab(fromIndex, toIndex) ___* (_A-1)
'       m_tabs (1-based Collection) ____ fromIndex __ toIndex _____.
'       Collection ______ API _____ Remove -> _____ Add (before/
'       after __) ______. ____ (Pane) _____, ________.
'
'     *___________________ (_A-1)*
'       _____________ index _____, ___________ Pane
'       ______, ______ m_tabs __________ index .
'       (NotifyTitleChanged ___________). .
'       ___ ActivateTab ______ (CloseTab _____).
'
'     *__________, __________ (_A-2b)*
'       ________ ActivateTab(_active) ____ RaiseEvent ActiveChanged
'       ___, TabBar/NavBar _ PushSyncToJs (_________ 60) ___.
'       _________________________________,
'       __________________________ JS ______.
'       -> Public Event _ 5 _________ (TabsReordered _____).
'
'     *Debug_TabOrderTitles ___ (_______)*
'       ___________ (____ (t<index>)) _ "|" .
'       ___. Wv2Thunks.Test_9_15_Reorder _ MoveTab .
'       WebView2 ______________ (Debug_NormalizeUrl _______).
'
' --- Wv2Browser.cls  _9.14 __ ---
'
'   _9.14 _____ (__________, __ Google):
'     *____________________ Google _____*
'       9.12a _____________________ (NormalizeUrl _____)
'       _____. _____ / ________ URL ___________,
'       https://www.google.com/search?q=... _____ Navigate __.
'
'     *NormalizeUrl _______ (_A)*
'       (1) ______ ("://" ___ / "about:" _)      -> .
'       (2) ______ __ _______               -> https:// ___ (URL)
'       (3) ____ (_____ / _____)             -> ___ -> BuildSearchUrl
'       Chrome/Edge ______________. "localhost:8080" .
'       __ (IsAlphaOnly) _ (1) _______________.
'
'     *URL ________ COM __________ (_E-1)*
'       percent-encoding __ UTF-8 ________ (______ ->
'       %E6%97%A5%E6%9C%AC _). ________ WebView2 _____ COM .
'       ________, UTF-8 ______________vtable_SAFEARRAY.
'       ________________________, __ HTML .
'       ADODB.Stream _____ (___________). ADODB.Stream .
'       Charset="utf-8" ___, __________ RFC3986 .
'       %XX ________.
'
'     *____ = Wv2Browser _ Private ____ (_P-1)*
'       __________ Browser ___ (NormalizeUrl _____) ___,
'       UrlEncodeComponent / Utf8Bytes / BuildSearchUrl _ Wv2Browser __ Private .
'       __. __ URL ____________ Wv2Url.bas _____ (9.13 _ JSON .
'       ______________). Wv2Json.bas __JSON ____________.
'
' --- Wv2Browser.cls  _9.12a __ ---
'
'   _9.12a _____ (______/________):
'     *______ TabUrlChanged(index, url, canGoBack, canGoForward)*
'       9.11a _ TabTitleChanged ____________. Browser .
'       ______, _________. _____ WithEvents _____.
'     - NotifyUrlChanged(senderPane, url, canBack, canFwd): Pane _____,
'       ____________ index ______ RaiseEvent TabUrlChanged.
'       (NotifyTitleChanged ___)
'
'     *HistoryChanged __________ (_2-A-1)*
'       __ HistoryChanged __SPA _____ EnableSpaHandlers .
'       __ (9.10d-2b, 9.10c-2 ___). _____/___________ SPA .
'       __________________, RegisterStandardHandlers .
'       __________. -> _______ 4 _ -> 5 __.
'       EnableSpaHandlers _ DOMContentLoaded ___________ (SPA .
'       ___________). __________ EnableSpaHandlers .
'       HistoryChanged _______.
'
'     *_____________ (________ API)*
'       _______________________, Pane .
'       _______. Browser ________________.
'         NavigateActive(url) / GoBackActive / GoForwardActive /
'         ReloadActive / StopActive
'         GetActiveUrl / GetActiveCanGoBack / GetActiveCanGoForward
'       *____ 63* ____________________, .
'         _______________ (9.11d _ newtab ___________).
'
'     *URL _______ (_3-A)* NavigateActive __________ https:// .
'       ____ ("example.com" -> "https://example.com"). .
'       _____ (___).
'
' --- Wv2Browser.cls  _9.11c __ ---
'
'   _9.11c _____ (_______ Bounds _____________, _A):
'     *__* AddTab _______ Bounds _ GetClientRect(m_hwndHost)=.
'       _______. _____ (__ 36px) _________, .
'       _______________ (UserForm_Resize _________).
'     *_____* Browser _____________ (___) __, .
'       _______ (___________) ___ (UserForm) ______.
'     *__ (_A)* ___________________, Browser .
'       _________________________________.
'       - _______: m_lastContentRect As RECT / m_hasContentRect As Boolean
'       - ResizeAll: ____ rc _ m_lastContentRect ____ m_hasContentRect=True.
'       - AddTab: m_hasContentRect _ True _________ Bounds ___.
'         ____ (ResizeAll ___) _____ GetClientRect(__) ________.
'       *______* Browser .
'       _____. UserForm ____ (___ UserForm_Resize ___ ResizeAll .
'       ____, ___ AddTab _______________).
'
' --- Wv2Browser.cls  _9.11b __ ---
'
'   _9.11b _____ (Wv2TabBar _ Environment ____):
'     - Property Get Environment() As Wv2Environment ___. Wv2TabBar .
'       ________________, Browser ___ Environment .
'       ______ (Environment _ 1 _________).
'       *___* Browser __Environment _______, .
'       _____ (_________) _____.
'
'   _9.11a _____ (_____________ + CloseTab, _A):
'     *___* Browser ________ (WebView2 _/MsForms _/__) .
'       ______. _____ Public Event _________ (RaiseEvent .
'       _______________). ____ WithEvents m_browser ___.
'     - ______: TabAdded(index) / TabClosed(index) / ActiveChanged(index) /
'       TabTitleChanged(index, title)
'     - AddTab: Collection _____ RaiseEvent TabAdded (ActivateTab __).
'       ___ ActivateTab _____, ______ TabAdded -> ActiveChanged .
'       __ 2 ___.
'     - ActivateTab: ____ RaiseEvent ActiveChanged.
'     - CloseTab(index): __. ClearBrowserRef (____ 58) -> Ctrl_Close (vtable 24,
'       __ Close __) -> m_tabs.Remove -> _______ -> RaiseEvent TabClosed.
'       ________ 4 __ (0 _=__ 0 __ / _____________ /
'       _______=__________ / ________). .
'       ActivateTab ___ (____ 57). ________________.
'     - NotifyTitleChanged(senderPane, title): __. Pane .
'       __, ____________ index _____ RaiseEvent TabTitleChanged.
'
'   _9.10e _____ (NewWindowRequested _________, _Y + __ Timer):
'     *________ (__)* Wv2Pane __ NewWindowRequested .
'       AddTabWithUrl ______________ (____ 14) __,
'       RequestNewTabAsync + SafeTimer (______) _____________.
'     - _______: m_newTabTimer (WithEvents SafeTimer) / m_pendingUrls (_____)
'     - RequestNewTabAsync(url): ________, New SafeTimer .
'       (SafeTimer.StartTimer ________________ m_uElapse _____,
'        _____ New __________)
'     - m_newTabTimer_TimerCall: ____________ AddTabWithUrl .
'     - AddTab _ pane.SetBrowser Me ____, _ Pane ___ (Browser) .
'       _______ (Wv2Pane _ NewWindowRequested __ AddTabWithUrl .
'       ________).
'     - Shutdown: ____ (Browser.m_tabs _ Pane.m_browser) ___. _ Pane .
'       ClearBrowserRef __________. UserForm_Terminate .
'       __ (____ 58). UserForm ________ Shutdown ____ 1 ___.
'
'   _9.10d-2b _____ (SPA _____, __-c + _F-1):
'     - EnableSpaHandlers(pane): SPA _____ (HistoryChanged + DOMContentLoaded)
'       _ pane ___________. ___ pane.View_AddHistoryChanged +
'       pane.View2_AddDOMContentLoaded _____. __ 4 .
'       (RegisterStandardHandlers) ______ SPA .
'       (9.10c-2 ___: SPA ______). ____________.
'     - AddTabWithUrlForSpa(hostName, folderPath, fileName): SPA _______.
'       AddTab -> View3_SetVirtualHostNameToFolderMapping (Navigate __View ___) ->
'       Navigate(https://<host>/<file>) -> EnableSpaHandlers __.
'       *____ 13* Navigate _____________ (___ "/" .
'       ___________ ERR_ACCESS_DENIED).
'
'   _9.10c-2 _____:
'     - RegisterStandardHandlers(pane): __ 4 ____ (NavigationStarting/
'       Completed/DocumentTitleChanged/NewWindowRequested) _ AddTab .
'       (__-A).
'
'   _9.10c-1 _____:
'     - ActivateTab(index): _______ (___________) .
'       pane.Ctrl_PutIsVisible ___. _______ 1 _____ (____ 57).
'       AddTab ________.
'
'   _____________. Wv2Environment _ 1 ______,
'   ____ N __ Wv2Pane (= __) ________.
'   UserForm _________ m_env + m_pane __ m_browser 1 .
'   _____ (UserForm _______ 9.10c).
'
'   ____ (_9.10 __ ____):
'     - ____ Wv2Browser (Wv2 _____, ____ 38)
'     - Wv2Pane ___ (_ D-1). _____ Collection _______,
'       ____ View_/Ctrl_/Settings_ __ ActivePane _______.
'     - Environment _ 1 _______ (_ E-1).
'     - ________ AddTab / Init _________ (___ a _).
'       -> UserForm ___ Do While ... DoEvents ___________.
'
'   9.10b _____ (____):
'     __  : Init / AddTab / AddTabWithUrl / AddTabWithHtml / TabCount /
'             ActivePane / Class_Terminate
'     ____: ActivateTab (____) / RemoveTab / UserForm __ /
'               _________ (___ 9.10c __)
'     -> ______ m_hwndHost __________ (.
'       1 ____). 9.10b ___ Pane _ Controller/View .
'       _______ (___).
'
'   Release __ (____ 43: _____):
'     Class_Terminate ___ Pane _______ Environment _____.
'     _ Wv2Pane _ Class_Terminate _ Settings->View2->View->Ctrl .
'     ______, ______Pane _ -> Environment__________.
''''''''''''''''''''''''''''''''''

Option Explicit

' ============================================================
' Browser .
' ============================================================
Public Enum BrowserState
    Bs_Created = 0     ' New __, Init .
    Bs_EnvReady = 1    ' Environment ____, .
    Bs_Failed = 2      ' Environment .
End Enum

' --- _______________ (__ StartWebView2 ___) ---
Private Const READY_TIMEOUT As Single = 10#

' --- _9.16 _____: _____ URL ______ (___, _________) ---
Private Const DEFAULT_SEARCH_TEMPLATE As String = "https://www.google.com/search?q="

' ============================================================
' .
' ============================================================
Private m_env      As Wv2Environment   ' 1 _, .
Private m_tabs     As Collection       ' Wv2Pane ___ (1-based)
Private m_dbgTagMap As Collection            ' _9.15 ___: ObjPtr(key) -> __ ("A","B",..). _______ Nothing
Private m_hwndHost As LongPtr          ' _____ HWND (_____)
Private m_state    As BrowserState

' --- _9.16 _____: _________ URL ______ (_1-A) ---
'   BuildSearchUrl _____________________________.
'   ___ Google. Class_Initialize _____, Property SearchEngine ______.
'   _: "https://www.google.com/search?q="  /  "https://www.bing.com/search?q="
Private m_searchTemplate As String

' _9.10c-1 _____: ________ index (1-based, 0 = _______).
'   ActivateTab __________. ActivePane ________.
Private m_activeIndex As Long

' --- _9.11c _____: ____________ (_A) ---
'   __ (UserForm) _ ResizeAll ________________, AddTab .
'   __ Bounds _____. .
'   __________. m_hasContentRect _ False __ (ResizeAll ___) .
'   AddTab _ GetClientRect(__) __________.
'   *Browser __________* .
'   ___, ________ Browser _____________ (_____).
Private m_lastContentRect As RECT
Private m_hasContentRect  As Boolean

' _9.10e __ (Timer ____) ___:
'   NewWindowRequested _______, .
'   ______ SafeTimer ___ URL ___. WithEvents _ TimerCall ____.
Private WithEvents m_newTabTimer As SafeTimer
Attribute m_newTabTimer.VB_VarHelpID = -1
Private m_pendingUrls As Collection

' _9.18 __ (_________) ___:
'   TabBar _ [__] ___ -> OnPaneWebMessage _ "settings" .
'   RequestSettingsTabAsync _____. .
'   AddTab -> WaitPaneReady (Controller ______) _____, WebMessage
'   _____________ ____ 14 ___________ (newtab .
'   ____). ___ newtab __________ SafeTimer _______.
'   *newtab ____ (m_pendingUrls) ________* ____ URL .
'     __. .
'     (_____________, ____63 ________).
'   *m_settingsBridge* _____ attach __ Wv2SettingsBridge _____.
'     ___________ WebView2 _________ Browser _ 1 ___.
' K-4c: _________ about:blank _____ (__ True)
Private m_quietOnShutdown As Boolean

Private WithEvents m_settingsTimer As SafeTimer
Attribute m_settingsTimer.VB_VarHelpID = -1

' --- K-2: __________ (____60) ---
'   ____________, _____________________.
'   *SafeTimer _________27 _________*
Private WithEvents m_syncTimer As SafeTimer
Attribute m_syncTimer.VB_VarHelpID = -1
Private m_syncPending As Boolean
Private m_settingsPending As Boolean
Private m_settingsBridge As Wv2SettingsBridge

' --- _9.26 _____ -> _9.27 ___: m_popupDelegateMode (Boolean) ---
'   9.26 _______ Boolean 1 _____, 9.27 _____ 3 .
'   m_newWindowMode (Long) _____ (__4 _b). Boolean _ 2 .
'   ________ 4 __________.
'   *PopupDelegateMode _____________________* __,
'     Test_9_26_Popup_On / _Off _ 9.27 ________.

' --- _9.27 _____: NewWindowRequested _ 3 _____________ ---
'   *____ (Wv2Pane __ Private Const ___)*
'     0 = NWM_LEGACY  : ____. put_Handled(TRUE) + RequestNewTabAsync _______.
'                       URL __________ POST _ GET ____ (= 405 ___).
'     1 = NWM_POPUP   : 9.26 _____. put_Handled _____________.
'                       POST ________________________.
'     2 = NWM_PREWARM : 9.27 ___. _____ CoreWebView2 _ put_NewWindow ___.
'                       POST ____, ___________.
'   *m_setHandledWithNewWindow (__5 ___)*
'     put_NewWindow ________ put_Handled(TRUE) ___________.
'     __________________ True. .
'     ____________ False ___________.
Private m_newWindowMode As Long
Private m_setHandledWithNewWindow As Boolean

'   *______ (____)*
'     m_prewarmPane    : _______ Pane (m_tabs ______ = ________)
'     m_prewarmClaimed : put_NewWindow _____ (= ____) ___. .
'                        __ WebView _ 2 _________, Claim _ 1 ______.
Private m_prewarmPane As Wv2Pane
Private m_prewarmClaimed As Boolean

'   *_____________ Timer (__1 _b)*
'     Init ________, _______ (Env -> _____ -> TabBar/NavBar) .
'     ____________. ____________ 4 ___ Controller .
'     ___________ (____ 20 ______).
Private WithEvents m_prewarmTimer As SafeTimer
Attribute m_prewarmTimer.VB_VarHelpID = -1
Private m_prewarmPending As Boolean

'   *___________ Timer*
'     NewWindowRequested _______ m_tabs.Add / ActivateTab _____,
'     __________________. RequestNewTabAsync ____,
'     ________ TimerCall ___ -> _________.
Private WithEvents m_adoptTimer As SafeTimer
Attribute m_adoptTimer.VB_VarHelpID = -1
Private m_adoptPending As Boolean
Private m_adoptUrl As String

'   *Pane _______ (_____)*
'     CreatePaneCore ____ WaitPaneReady _ DoEvents _____, .
'     m_prewarmTimer_TimerCall _______. __________________.
'     *_9.28 _ Boolean -> Long ___* 9.27 ____3 ____________.
'     Boolean _________________________, .
'     _________. __/_____1 ________ > 0______.
Private m_creatingPane As Long

'   ___________________________ (ms).
'   ______________________, _________________,
'   ________ POST __________. 1.5 ________.
Private Const PREWARM_DELAY_MS As Long = 1500

' --- _9.11a _____: _____________ (_A) ---
'   Browser __________________. _____ WebView2 .
'   MsForms _______, ______. RaiseEvent .
'   _____. ____ Private WithEvents m_browser As Wv2Browser _____.
'   _TabAdded(index)              : ________ (Collection ____)
'   _TabClosed(index)             : ________ (Collection ____)
'   _ActiveChanged(index)         : ____________ (ActivateTab ___)
'   _TabTitleChanged(index,title) : ____________ (Pane __ NotifyTitleChanged __)
Public Event TabAdded(ByVal index As Long)
Public Event TabClosed(ByVal index As Long)
Public Event ActiveChanged(ByVal index As Long)
Public Event TabTitleChanged(ByVal index As Long, ByVal title As String)

' --- _9.12a: URL ________ (______ / _________) ---
'   Pane _ NavigationCompleted / HistoryChanged __ NotifyUrlChanged _______.
'   *___* Browser _____________. _______ RaiseEvent .
'   ________.
Public Event TabUrlChanged(ByVal index As Long, ByVal url As String, _
                           ByVal canGoBack As Boolean, ByVal canGoForward As Boolean)


' ============================================================
' Class_Initialize
' ============================================================
Private Sub Class_Initialize()
    m_quietOnShutdown = True             ' K-4c (_____ about:blank ____)
    Set m_tabs = New Collection
    Set m_env = Nothing
    m_hwndHost = 0
    m_state = Bs_Created
    m_activeIndex = 0
    Set m_pendingUrls = New Collection   ' _9.10e (NewWindowRequested __ URL ___)
    Set m_newTabTimer = Nothing          ' _9.10e (_____ New __)
    m_hasContentRect = False             ' _9.11c (_______ ______)
    m_searchTemplate = DEFAULT_SEARCH_TEMPLATE  ' _9.16 (__ Google. Property ______)
    Set m_settingsTimer = Nothing        ' _9.18 (_____ New __)
    m_settingsPending = False            ' _9.18 (____________)
    Set m_settingsBridge = Nothing       ' _9.18 (attach ______)

    ' _9.27: NewWindowRequested .
    '   *_9.28 ____ 0 -> 2 ___* ___________, POST .
    '     ______________________. ___ PREWARM_DELAY_MS .
    '     _________________________, .
    '     ________________________.
    '     ________________ NewWindowMode = 0 ___.
    m_newWindowMode = 2                  ' ___________ (NWM_PREWARM)
    m_setHandledWithNewWindow = True     ' put_NewWindow _ put_Handled ___ (__5)
    Set m_prewarmPane = Nothing
    m_prewarmClaimed = False
    Set m_prewarmTimer = Nothing         ' _____ New .
    m_prewarmPending = False
    Set m_adoptTimer = Nothing           ' _____ New .
    m_adoptPending = False
    m_adoptUrl = ""
    m_creatingPane = 0                   ' _9.28: Boolean -> Long (_______)

    ' _9.20: ________________ (__).
    '   _______ New _, LoadEngineTemplate() .
    '   __. ____ m_searchTemplate ____, _____ (Google) ___.
    '   _____ Browser ________________ BindBrowser __.
    Dim initBridge As Wv2SettingsBridge
    Set initBridge = New Wv2SettingsBridge
    Dim savedTemplate As String
    savedTemplate = initBridge.LoadEngineTemplate()
    If Len(savedTemplate) > 0 Then
        m_searchTemplate = savedTemplate
        Debug.Print "Wv2Browser.Class_Initialize: ___________ = " & savedTemplate
    End If
    Set initBridge = Nothing
End Sub


' ============================================================
' Init ? Environment _ 1 ____ Ready .
'   hwndHost   : ____________ HWND
'   userDataPath: WebView2 _ userDataFolder
'   ___      : Environment _ Ready ____ True
' ============================================================
Public Function Init( _
    ByVal hwndHost As LongPtr, _
    ByVal userDataPath As String) As Boolean

    If m_state <> Bs_Created Then
        Debug.Print "Wv2Browser.Init: ______ (state=" & m_state & ")"
        Init = (m_state = Bs_EnvReady)
        Exit Function
    End If

    m_hwndHost = hwndHost

    ' --- Environment __ ---
    Set m_env = New Wv2Environment
    m_env.Init userDataPath

    ' --- Ready __ (_________________ = ___ a) ---
    If Not WaitEnvReady() Then
        Debug.Print "Wv2Browser.Init: Environment ____ (state=" & _
                    m_env.state & ", lastError=&H" & Hex(m_env.LastError) & ")"
        m_state = Bs_Failed
        Exit Function
    End If

    Debug.Print "Wv2Browser.Init: Environment __ OK (pEnv=" & _
                m_env.EnvironmentPtr & ")"
    m_state = Bs_EnvReady

    ' --- _9.27: ____ (______) ________ (__1 _b) ---
    '   ____________, ________________ (.
    '   _____ / TabBar / NavBar _ Controller ________) ______.
    '   ______ Timer _ PREWARM_DELAY_MS ________.
    RequestPrewarmAsync

    Init = True
End Function


' ============================================================
' AddTab ? _______ 1 _____ (Navigate ___)
'   _ Wv2Pane ____, __ Environment __ Controller ___,
'   Ready _____ Bounds ____, Collection _____.
'   ___: _____ Wv2Pane (___ Nothing)
' ============================================================
Public Function AddTab() As Wv2Pane
    ' --- __ (Pane __ _ ________) _ CreatePaneCore ______ (_9.27) ---
    Dim pane As Wv2Pane
    Set pane = CreatePaneCore()
    If pane Is Nothing Then Exit Function

    ' --- 6. Collection ___ ---
    m_tabs.Add pane
    Debug.Print "Wv2Browser.AddTab: ____ OK (______=" & m_tabs.Count & ")"

    ' --- 6.5 _______ (_9.11a, _______) ---
    '   ActivateTab ______. ______________, .
    '   ________________ (AddTab ____ ActivateTab ___).
    RaiseEvent TabAdded(m_tabs.Count)

    ' --- 7. ________________ (___ a) ---
    '   ___________ ActivateTab _____ (____).
    '   ________________ ActivateTab _ 1 ________.
    ActivateTab m_tabs.Count

    ' --- 8. _________________ (_9.27) ---
    '   ________________________________.
    '   AddTab _______________, ______ 1 ____.
    EnsurePrewarmScheduled

    Set AddTab = pane
End Function


' ============================================================
' CreatePaneCore / CreatePaneCoreBody (_9.27, _ AddTab _______ = __2 _a)
'
'   _Pane _ 1 ____________________. m_tabs .
'   TabAdded ____ActivateTab _______. ______ 2 _:
'     _AddTab            -> ___ Pane .
'     _CreatePrewarmPaneNow -> ___ Pane .
'
'   *2 ___ (Core / CoreBody) ___*
'     CoreBody ______ Exit Function _____. _____ (m_creatingPane) .
'     CoreBody __________, __ Exit ___________________.
'     _____ (Core) ____, Body ____________ 1 ______.
'
'   *__________* CoreBody __ WaitPaneReady _ DoEvents ___. .
'     m_prewarmTimer_TimerCall ______ Pane _________. .
'     _____ (____ 14 / 20 ___________).
' ============================================================
Private Function CreatePaneCore() As Wv2Pane
    ' *_9.28: ___ -> ____* _________, ________ 0 _____.
    m_creatingPane = m_creatingPane + 1
    Set CreatePaneCore = CreatePaneCoreBody()
    m_creatingPane = m_creatingPane - 1
    If m_creatingPane < 0 Then m_creatingPane = 0   ' .
End Function

Private Function CreatePaneCoreBody() As Wv2Pane
    If m_state <> Bs_EnvReady Then
        Debug.Print "Wv2Browser.CreatePaneCore: Environment __ Ready (state=" & m_state & ")"
        Exit Function
    End If

    ' --- 1. Pane ____, __ Environment __ Controller ___ ---
    Dim pane As Wv2Pane
    Set pane = New Wv2Pane
    pane.Init m_env, m_hwndHost

    ' --- 2. Controller Ready __ (_______) ---
    If Not WaitPaneReady(pane) Then
        If pane.IsFailed Then
            Debug.Print "Wv2Browser.CreatePaneCore: Controller ____ (state=" & _
                        pane.state & ", lastError=&H" & Hex(pane.LastError) & ")"
        Else
            Debug.Print "Wv2Browser.CreatePaneCore: Controller ________ (state=" & _
                        pane.state & ")"
        End If
        Exit Function
    End If
    Debug.Print "Wv2Browser.CreatePaneCore: Controller __ OK (pCtrl=" & pane.CtrlPtr & ")"

    ' --- 3. Bounds _______ (_9.11c: ______________, _A) ---
    '   __ (UserForm) _ ResizeAll _____________, .
    '   (______________). ______________.
    '   ________________ UserForm_Resize _________.
    Dim rc As RECT
    If m_hasContentRect Then
        rc = m_lastContentRect
        Debug.Print "Wv2Browser.CreatePaneCore: __ Bounds __________________"
    Else
        GetClientRect m_hwndHost, rc
        Debug.Print "Wv2Browser.CreatePaneCore: __ Bounds __________ (ResizeAll ___)"
    End If
    Dim hr As Long
    hr = pane.Ctrl_PutBounds(rc)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.CreatePaneCore: Ctrl_PutBounds __ hr=&H" & Hex(hr)
        ' Bounds _____________ (____________)
    End If

    ' --- 4. ICoreWebView2 (View) ___ ---
    '   *__* Controller ______ m_pView = 0 ___. View_Navigate /
    '   View_NavigateToString _ m_pView ______, _________.
    '   *_9.27 ______* put_NewWindow _______ ViewPtr ____.
    hr = pane.Ctrl_GetCoreWebView2()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.CreatePaneCore: Ctrl_GetCoreWebView2 __ hr=&H" & Hex(hr)
        Exit Function
    End If
    If pane.ViewPtr = 0 Then
        Debug.Print "Wv2Browser.CreatePaneCore: ViewPtr _ 0 (_________)"
        Exit Function
    End If
    Debug.Print "Wv2Browser.CreatePaneCore: View __ OK (pView=" & pane.ViewPtr & ")"

    ' --- 5. __________________ (__-A, _9.10c-2) ---
    '   ________ (NavigationStarting/Completed/DocumentTitleChanged/
    '   NewWindowRequested/HistoryChanged) ___________.
    '   *__________* ______________URL _________.
    RegisterStandardHandlers pane

    ' --- 5.5 Wv2Pane _________ (_9.10e, _Y) ---
    '   NewWindowRequested _ Pane _ m_browser ____________.
    '   *_________, ____ Shutdown -> ClearBrowserRef ___ (____ 58).
    pane.SetBrowser Me

    Set CreatePaneCoreBody = pane
End Function


' ============================================================
' RegisterStandardHandlers ? __ 1 .
'   _9.10c-2 _____ (__-A). __ StartWebView2 _____ 10?13 .
'   ______. ____________ 1 _______ (____ 57 ___).
'
'   *____ 45/52 ___* ____ View_AddXxx ____ COM .
'   WebView2 ____________. Wv2Pane _____ VTABLE_OBJ_OFFSET .
'   ________________, ____ Wv2Pane .
'   _______ (____________).
' ============================================================
Private Sub RegisterStandardHandlers(ByVal pane As Wv2Pane)
    Dim hr As Long

    hr = pane.View_AddNavigationStarting()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: NavigationStarting __ hr=&H" & Hex(hr)

    hr = pane.View_AddNavigationCompleted()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: NavigationCompleted __ hr=&H" & Hex(hr)

    hr = pane.View_AddDocumentTitleChanged()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: DocumentTitleChanged __ hr=&H" & Hex(hr)

    hr = pane.View_AddNewWindowRequested()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: NewWindowRequested __ hr=&H" & Hex(hr)

    ' --- _9.12a: HistoryChanged __________ (_2-A-1) ---
    '   __/___________ SPA ___________________,
    '   SPA ____ (EnableSpaHandlers) ______________.
    '   pushState/replaceState/popstate ___, _____/________.
    hr = pane.View_AddHistoryChanged()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: HistoryChanged __ hr=&H" & Hex(hr)

    Debug.Print "Wv2Browser.RegisterStandardHandlers: ______ 5 _ ____"
End Sub


' ============================================================
' EnableSpaHandlers ? SPA __________ pane _____ (_9.10d-2b, __-c)
'
'   SPA (history __ + DOM __) ____________ 2 ______:
'     - HistoryChanged   (__ ICoreWebView2, vtable 13) : pushState/replaceState/
'                          popstate ______ (____ 12)
'     - DOMContentLoaded (ICoreWebView2_2, vtable 64)     : .
'                          (pushState _______ = HistoryChanged ______)
'
'   __ 4 ____ (RegisterStandardHandlers) _____, SPA .
'   ______ (9.10c-2 ___). AddTabWithUrlForSpa _______,
'   Public _____________________ (__-c).
'
'   *____ 45* View_AddHistoryChanged / View2_AddDOMContentLoaded .
'   Wv2Pane __ VTABLE_OBJ_OFFSET _____________. .
'   _____________ (____________).
' ============================================================
Public Sub EnableSpaHandlers(ByVal pane As Wv2Pane)
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.EnableSpaHandlers: pane _ Nothing"
        Exit Sub
    End If

    ' --- _9.12a: HistoryChanged ________ ---
    '   RegisterStandardHandlers (_______) _______, .
    '   _______. SPA _______ (DOMContentLoaded) _____.
    Dim hr As Long
    hr = pane.View2_AddDOMContentLoaded()
    If hr <> 0 Then Debug.Print "Wv2Browser.EnableSpaHandlers: DOMContentLoaded __ hr=&H" & Hex(hr)

    Debug.Print "Wv2Browser.EnableSpaHandlers: SPA ____ 1 _ " & _
                "(DOMContentLoaded) ____ _HistoryChanged __________ (9.12a)"
End Sub


' ============================================================
' AddTabWithUrl ? ______, URL _ Navigate .
' ============================================================
Public Function AddTabWithUrl(ByVal url As String) As Wv2Pane
    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_Navigate(url)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithUrl: View_Navigate __ hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.AddTabWithUrl: Navigate OK (" & url & ")"
    End If

    Set AddTabWithUrl = pane
End Function


' ============================================================
' AddTabWithHtml ? ______, HTML .
'
'   *____ 6* NavigateToString _________ origin _ about:blank
'   (= "null" origin) ___. __ origin __ history.pushState /
'   replaceState _ SecurityError _____ (Web ____).
'   pushState/replaceState _______ SPA ____, 9.10d _ http(s)
'   origin ___ (AddTabWithUrl) ______.
' ============================================================
Public Function AddTabWithHtml(ByVal htmlContent As String) As Wv2Pane
    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_NavigateToString(htmlContent)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithHtml: View_NavigateToString __ hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.AddTabWithHtml: NavigateToString OK (" & _
                    Len(htmlContent) & " __)"
    End If

    Set AddTabWithHtml = pane
End Function


' ============================================================
' AddTabWithUrlForSpa ? SPA _________ (_9.10d-2b, _F-1)
'
'   _____________________, ___ HTML _ https origin .
'   ___ SPA ___________, SPA ____________.
'
'   __ (_____):
'     1. AddTab            : Controller/View _____ 4 .
'     2. SetVirtualHostNameToFolderMapping : *Navigate ___View ______*
'                            (View ___ AddTab ____________ OK)
'     3. View_Navigate     : https://<hostName>/<fileName>
'                            *____ 13* ___ "/" _ ERR_ACCESS_DENIED .
'                            .
'     4. EnableSpaHandlers : HistoryChanged + DOMContentLoaded .
'
'   __:
'     hostName   : ______ (_ "appassets.example")
'     folderPath : .
'     fileName   : _____________ (_ "index.html")
'
'   accessKind _ ALLOW (=1) __. ___________.
'
'   ___: ____ Wv2Pane (____ Nothing, .
'           _____/Navigate __________ = ________)
' ============================================================
Public Function AddTabWithUrlForSpa( _
    ByVal hostName As String, _
    ByVal folderPath As String, _
    ByVal fileName As String) As Wv2Pane

    ' --- 1. _______ (View ____________________) ---
    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then Exit Function

    ' --- 2. _____________ (*Navigate _____) ---
    Dim hr As Long
    hr = pane.View3_SetVirtualHostNameToFolderMapping(hostName, folderPath, 1)   ' 1 = ALLOW
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithUrlForSpa: _______ hr=&H" & Hex(hr)
        Set AddTabWithUrlForSpa = pane   ' .
        Exit Function
    End If

    ' --- 3. https://<host>/<file> _ Navigate (____ 13: _________) ---
    Dim url As String
    url = "https://" & hostName & "/" & fileName
    hr = pane.View_Navigate(url)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithUrlForSpa: Navigate __ hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.AddTabWithUrlForSpa: Navigate OK (" & url & ")"
    End If

    ' --- 4. SPA _____ (HistoryChanged + DOMContentLoaded) ___ ---
    EnableSpaHandlers pane

    Set AddTabWithUrlForSpa = pane
End Function


' ============================================================
' TabCount ? .
' ============================================================
Public Property Get TabCount() As Long
    TabCount = m_tabs.Count
End Property


' ============================================================
' ActivePane ? ________ Wv2Pane .
'   _9.10c-1 _ activeIndex _______ (_______).
'   m_activeIndex (1-based, 0 = __) ________.
' ============================================================
Public Property Get ActivePane() As Wv2Pane
    If m_activeIndex < 1 Or m_activeIndex > m_tabs.Count Then
        Debug.Print "Wv2Browser.ActivePane: _________ (activeIndex=" & _
                    m_activeIndex & ", tabCount=" & m_tabs.Count & ")"
        Set ActivePane = Nothing
        Exit Property
    End If
    Set ActivePane = m_tabs.Item(m_activeIndex)
End Property

' ============================================================
' Environment ? __ Environment _____ (_9.11b, Wv2TabBar _)
'   Wv2TabBar _____________________, .
'   Browser _____ Environment (______) ___. Init .
'   (Bs_EnvReady) ______. _ Ready ___ Nothing ___.
' ============================================================
Public Property Get Environment() As Wv2Environment
    Set Environment = m_env
End Property


' ============================================================
' ActiveIndex ? ________ index (1-based, 0 = __)
' ============================================================
Public Property Get ActiveIndex() As Long
    ActiveIndex = m_activeIndex
End Property


' ============================================================
' ActivateTab ? __ index __________, .
'   _9.10c-1 _____. __________ (single source of truth).
'
'   __: index (1-based)
'   ___: ___ True
'
'   ____ (___ A _): ________, index .
'   Ctrl_PutIsVisible(True), __ False ___. _______,
'   ______________. ___________________.
'
'   *__* ____________ + activeIndex ________.
'   ____________. AddTab ___________ (__) ___,
'   _________ 1 _______.
' ============================================================
Public Function ActivateTab(ByVal index As Long) As Boolean
    If index < 1 Or index > m_tabs.Count Then
        Debug.Print "Wv2Browser.ActivateTab: index ___ (index=" & index & _
                    ", tabCount=" & m_tabs.Count & ")"
        Exit Function
    End If

    Dim i As Long
    Dim pane As Wv2Pane
    Dim hr As Long
    For i = 1 To m_tabs.Count
        Set pane = m_tabs.Item(i)
        hr = pane.Ctrl_PutIsVisible(i = index)
        If hr <> 0 Then
            Debug.Print "Wv2Browser.ActivateTab: __" & i & " _ PutIsVisible __ hr=&H" & Hex(hr)
        End If
    Next i

    m_activeIndex = index
    Debug.Print "Wv2Browser.ActivateTab: __" & index & " _______ (__)"

    ' --- __________ (_9.11a, _______) ---
    RaiseEvent ActiveChanged(index)
    ActivateTab = True
End Function


' ============================================================
' CloseTab ? __ index _______ (_9.11a _____)
'
'   __: index (1-based)
'   ___: ___ True
'
'   __:
'     1. index .
'     2. ___ Pane ________ (ClearBrowserRef, ____ 58) ->
'        Ctrl_Close _ Controller _______ (vtable 24, __ Close ___)
'     3. m_tabs.Remove index (______ 0 _ Pane.Class_Terminate)
'     4. m_activeIndex ___ + _______ (ActivateTab ___, ____ 57):
'        ___________ : ____, ______ (__) .
'        ______________ : m_activeIndex _ 1 .
'        ____ 0 _____ : m_activeIndex = 0 (_______, __ 0 __)
'     5. RaiseEvent TabClosed(index) (__________ ActivateTab ___ ActiveChanged ___)
'
'   *__ 0 __* ___________ Browser _____. __ (.
'   ___) _______, ________________________.
' ============================================================
Public Function CloseTab(ByVal index As Long) As Boolean
    If index < 1 Or index > m_tabs.Count Then
        Debug.Print "Wv2Browser.CloseTab: index ___ (index=" & index & _
                    ", tabCount=" & m_tabs.Count & ")"
        Exit Function
    End If

    Dim pane As Wv2Pane
    Set pane = m_tabs.Item(index)

    ' --- 1. _________ (____ 58) ---
    '   ____ Browser_Pane _______. Shutdown _____.
    pane.ClearBrowserRef

    ' --- 2. Controller _______ ---
    '   Ctrl_Close _ View3/View2/View/Ctrl ___ Release _, .
    '   ____ (__ Close _ S_OK ___).
    Dim hr As Long
    hr = pane.Ctrl_Close()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.CloseTab: Ctrl_Close __ hr=&H" & Hex(hr) & " (__)"
    End If

    ' --- 3. Collection ____ (______ 0 _ Pane.Class_Terminate) ---
    Set pane = Nothing
    m_tabs.Remove index
    Debug.Print "Wv2Browser.CloseTab: __" & index & " ____ (_____=" & m_tabs.Count & ")"

    ' --- 4. _______ + _______ ---
    Dim newActive As Long
    If m_tabs.Count = 0 Then
        ' __ 0 _: _______ (__ 0 __)
        m_activeIndex = 0
        Debug.Print "Wv2Browser.CloseTab: _________ (activeIndex=0)"
    ElseIf index < m_activeIndex Then
        ' _____________: index _ 1 .
        ' ___________. ___________ ActivateTab ___.
        newActive = m_activeIndex - 1
        ActivateTab newActive
    ElseIf index = m_activeIndex Then
        ' __________: ____, ______ (__).
        '   Remove _, _________ index ___. index .
        '   ____ (________) ____________.
        If index <= m_tabs.Count Then
            newActive = index
        Else
            newActive = m_tabs.Count
        End If
        ActivateTab newActive
    Else
        ' _____________: activeIndex ______.
        '   ____________ ActivateTab _____, _________.
        ActivateTab m_activeIndex
    End If

    ' --- 5. _______ (_9.11a, _______) ---
    RaiseEvent TabClosed(index)
    CloseTab = True
End Function


' ============================================================
' MoveTab ? ___ fromIndex __ toIndex ______ (_9.15 _____)
'
'   _____: Wv2TabBar.OnPaneWebMessage _ {"cmd":"reorder","from":N,"to":M}
'   _______ 1-based. from == to _________ False ___.
'
'   *__*
'     1. _____ Pane ______ (index __________).
'     2. ____ Pane _ Remove __, _____ Add (before/after) _____.
'        _to == 1            : ___ (Add pane, , 1)
'        _____            : (to-1) ______ (Add pane, , , to-1)
'          _ Remove __ Collection _____ index _________.
'     3. ___________ Pane __________ m_activeIndex ___.
'     4. ActivateTab ________ -> RaiseEvent ActiveChanged -> .
'        (____ 60, _A-2b). __________.
'
'   *______* ____ (Remove/Add) _ PutIsVisible ___ Controller
'     _______________, WebMessage .
'     (activate/close ___. newtab ____ SafeTimer _____).
' ============================================================
Public Function MoveTab(ByVal fromIndex As Long, ByVal toIndex As Long) As Boolean
    Dim n As Long
    n = m_tabs.Count

    ' --- ______ / _______ ---
    If fromIndex < 1 Or fromIndex > n Then
        Debug.Print "Wv2Browser.MoveTab: from ___ (from=" & fromIndex & ", tabCount=" & n & ")"
        Exit Function
    End If
    If toIndex < 1 Or toIndex > n Then
        Debug.Print "Wv2Browser.MoveTab: to ___ (to=" & toIndex & ", tabCount=" & n & ")"
        Exit Function
    End If
    If fromIndex = toIndex Then
        ' _______________ (____).
        MoveTab = True
        Exit Function
    End If

    ' --- 1. _____ Pane ______ (index __________) ---
    Dim activePaneRef As Wv2Pane
    If m_activeIndex >= 1 And m_activeIndex <= n Then
        Set activePaneRef = m_tabs.Item(m_activeIndex)
    End If

    ' --- 2. __ Pane ______ Remove ---
    Dim moving As Wv2Pane
    Set moving = m_tabs.Item(fromIndex)
    m_tabs.Remove fromIndex

    ' Remove __ Collection ____ n-1. ___ moving _ toIndex ______.
    '   Collection.Add _ before/after __Remove __ Collection ___ index____.
    '   __________ moving _ toIndex ________.
    If toIndex = 1 Then
        ' ___: _____ (before:=1) ______.
        m_tabs.Add moving, , 1
    Else
        ' toIndex ______ = (toIndex-1) ____________.
        '   Remove __ Collection __ 1..(toIndex-1) .
        '   __ (fromIndex _______ toIndex __________, .
        '   fromIndex < toIndex ___ (toIndex-1) _____ toIndex _______).
        '   _______ after:=(toIndex-1) ________ toIndex ______.
        m_tabs.Add moving, , , toIndex - 1
    End If

    Debug.Print "Wv2Browser.MoveTab: ___ " & fromIndex & " -> " & toIndex & " ___ (tabCount=" & m_tabs.Count & ")"

    ' --- 3. _____ Pane ____ index ____ ---
    If Not activePaneRef Is Nothing Then
        Dim i As Long
        For i = 1 To m_tabs.Count
            If m_tabs.Item(i) Is activePaneRef Then
                m_activeIndex = i
                Exit For
            End If
        Next i
    End If

    ' --- 4. _______ -> ActiveChanged -> ____ (_A-2b) ---
    '   _____________, ActivateTab ______ PushSyncToJs .
    '   __, _______ TabBar/NavBar ______.
    If m_activeIndex >= 1 And m_activeIndex <= m_tabs.Count Then
        ActivateTab m_activeIndex
    End If

    MoveTab = True
End Function


' ============================================================
' Debug_TabOrderTitles ? _______ "|" __________ (_9.15, ___)
'
'   MoveTab ________ WebView2 ___________________.
'   _________ (_____ "(t<index>)") ____, .
'   Debug_SeedDummyTabs _____ Pane _ ObjPtr -> ____ ("A","B",...) .
'   ___ m_dbgTagMap _________, _____ index .
'   _________________________.
'   Debug_NormalizeUrl _________________.
' ============================================================
Public Function Debug_TabOrderTitles() As String
    Dim i As Long
    Dim parts As String
    Dim tag As String
    For i = 1 To m_tabs.Count
        tag = ""
        ' ____________ ObjPtr ___ (___________)
        If Not m_dbgTagMap Is Nothing Then
            Dim key As String
            key = CStr(ObjPtr(m_tabs.Item(i)))
            On Error Resume Next
            tag = m_dbgTagMap.Item(key)
            On Error GoTo 0
        End If
        ' ___________________ (___________)
        If Len(tag) = 0 Then
            On Error Resume Next
            tag = m_tabs.Item(i).DocumentTitle
            On Error GoTo 0
            If Len(tag) = 0 Then tag = "(t" & i & ")"
        End If
        If i > 1 Then parts = parts & "|"
        parts = parts & tag
    Next i
    Debug_TabOrderTitles = parts
End Function


' ============================================================
' Debug_SeedDummyTabs ? ______________ n ____ (_9.15, ___)
'
'   Test_9_15_Reorder _ WebView2 ______ MoveTab .
'   __________. __ Wv2Pane _ n _ m_tabs ___, _ Pane _ ObjPtr .
'   ____ ("A","B","C",...) ______ m_dbgTagMap ___. activeIndex _ 1.
'   __ MoveTab _______________________.
'   ____ (ActivateTab _ PutIsVisible) ____ Pane _ View _ Nothing .
'   _____________ (___________).
' ============================================================
Public Sub Debug_SeedDummyTabs(ByVal n As Long)
    Set m_tabs = New Collection
    Set m_dbgTagMap = New Collection
    Dim i As Long
    Dim p As Wv2Pane
    Dim tag As String
    For i = 1 To n
        Set p = New Wv2Pane
        m_tabs.Add p
        tag = Chr$(64 + i)   ' 1->"A", 2->"B", ...
        m_dbgTagMap.Add tag, CStr(ObjPtr(p))
    Next i
    If n >= 1 Then
        m_activeIndex = 1
    Else
        m_activeIndex = 0
    End If
End Sub


' ============================================================
' Debug_ActiveIndex ? ___ m_activeIndex ___ (_9.15, ___)
'   MoveTab ___________________.
' ============================================================
Public Function Debug_ActiveIndex() As Long
    Debug_ActiveIndex = m_activeIndex
End Function


' ============================================================
' NotifyTitleChanged ? Pane _____________, index .
'   RaiseEvent TabTitleChanged _____ (_9.11a _____)
'
'   _____: Wv2Pane.View_OnDocumentTitleChanged (____ m_browser __)
'   __: senderPane = _________ Pane (____ Me), title = .
'
'   *___* Collection _____________ index ___. .
'   ________________. ________ (.
'   __________) _____.
' ============================================================
Public Sub NotifyTitleChanged(ByVal senderPane As Wv2Pane, ByVal title As String)
    Dim i As Long
    For i = 1 To m_tabs.Count
        If m_tabs.Item(i) Is senderPane Then
            RaiseEvent TabTitleChanged(i, title)
            Exit Sub
        End If
    Next i
    Debug.Print "Wv2Browser.NotifyTitleChanged: senderPane _ m_tabs _______ (__)"
End Sub


' ============================================================
' NotifyUrlChanged (_9.12a, Pane ___ URL ________ index ___)
'
'   Pane _ NavigationCompleted / HistoryChanged ______. .
'   ___ index _____, RaiseEvent TabUrlChanged __.
'   (9.11a _ NotifyTitleChanged _______)
'
'   *________________* ____/_________ m_tabs .
'   ________, ____ m_browser _________ (SetBrowser ____)
'   __, _________.
' ============================================================
Public Sub NotifyUrlChanged(ByVal senderPane As Wv2Pane, ByVal url As String, _
                            ByVal canGoBack As Boolean, ByVal canGoForward As Boolean)
    Dim i As Long
    For i = 1 To m_tabs.Count
        If m_tabs.Item(i) Is senderPane Then
            RaiseEvent TabUrlChanged(i, url, canGoBack, canGoForward)
            Exit Sub
        End If
    Next i
    Debug.Print "Wv2Browser.NotifyUrlChanged: senderPane _ m_tabs _______ (__)"
End Sub


' ============================================================
' _____________ (_9.12a, ________ API)
'
'   _______________________, Pane .
'   ___ (_____ Browser _____________). Browser .
'   __________, _____ Browser __________.
'
'   *____ 63* _________________ (Navigate/GoBack/GoForward/
'     Reload/Stop __________) ___, WebMessage .
'     ________. 9.11d _ newtab (Controller ____) ___________.
'
'   __ 0 (_______) ________ False / ______.
' ============================================================
Public Function NavigateActive(ByVal url As String) As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.NavigateActive: _________ (__)"
        Exit Function
    End If

    Dim target As String
    target = NormalizeUrl(url)
    If Len(target) = 0 Then
        Debug.Print "Wv2Browser.NavigateActive: URL __ (__)"
        Exit Function
    End If

    Dim hr As Long
    hr = pane.View_Navigate(target)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.NavigateActive: Navigate __ hr=&H" & Hex(hr) & " (" & target & ")"
        Exit Function
    End If

    Debug.Print "Wv2Browser.NavigateActive: Navigate OK (" & target & ")"
    NavigateActive = True
End Function

Public Function GoBackActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    If Not pane.View_GetCanGoBack() Then
        Debug.Print "Wv2Browser.GoBackActive: ____ (____)"
        Exit Function
    End If

    Dim hr As Long
    hr = pane.View_GoBack()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.GoBackActive: GoBack __ hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.GoBackActive: GoBack OK"
    GoBackActive = True
End Function

Public Function GoForwardActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    If Not pane.View_GetCanGoForward() Then
        Debug.Print "Wv2Browser.GoForwardActive: ____ (____)"
        Exit Function
    End If

    Dim hr As Long
    hr = pane.View_GoForward()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.GoForwardActive: GoForward __ hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.GoForwardActive: GoForward OK"
    GoForwardActive = True
End Function

Public Function ReloadActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_Reload()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.ReloadActive: Reload __ hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.ReloadActive: Reload OK"
    ReloadActive = True
End Function

Public Function StopActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_Stop()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.StopActive: Stop __ hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.StopActive: Stop OK"
    StopActive = True
End Function


' --- ____________ (__________) ---
'   ____ (TabUrlChanged) _____, .
'   _____________________.
Public Function GetActiveUrl() As String
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function
    If Not pane.IsReady Then Exit Function
    GetActiveUrl = pane.View_GetSource()
End Function

Public Function GetActiveCanGoBack() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function
    If Not pane.IsReady Then Exit Function
    GetActiveCanGoBack = pane.View_GetCanGoBack()
End Function

Public Function GetActiveCanGoForward() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function
    If Not pane.IsReady Then Exit Function
    GetActiveCanGoForward = pane.View_GetCanGoForward()
End Function


' ============================================================
' NormalizeUrl (_9.12a, _______________, _3-A)
'
'   ___________________ Navigate ___ URL ____.
'     "example.com"        -> "https://example.com"
'     "https://example.com" -> .
'     "about:blank"        -> ____ (__________)
'
'   *________________ (_3-A)* .
'     ________________ URL ___________________.
'     ___________ Navigate _____, .
'     _________.
' ============================================================
Private Function NormalizeUrl(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function

    ' __________ ("://" ___, ___ "about:" ________)
    If InStr(t, "://") > 0 Then
        NormalizeUrl = t
        Exit Function
    End If

    ' about: / data: / file: ________ (____________)
    Dim colonPos As Long
    colonPos = InStr(t, ":")
    If colonPos > 1 Then
        Dim head As String
        head = Left$(t, colonPos - 1)
        ' ______________ ("example.com:8080" .
        ' __________)
        If Wv2Url.IsAlphaOnly(head) Then
            NormalizeUrl = t
            Exit Function
        End If
    End If

    ' __________________. ___ URL / ___ ______ (_A, 9.14).
    '   _______ __ _______ -> _________ https:// ___ (URL)
    '   _____ (_____ / _____) -> ________ Google .
    ' Chrome/Edge ______________. "localhost" ______________.
    If InStr(t, ".") > 0 And InStr(t, " ") = 0 Then
        NormalizeUrl = "https://" & t
    Else
        NormalizeUrl = BuildSearchUrl(t)
    End If
End Function


' ============================================================
' Debug_NormalizeUrl (_9.14, _______)
'
'   NormalizeUrl _ Private ___, ______________________.
'   NavigateActive ___________ NormalizeUrl _________,
'   _______________________.
'   (Wv2Thunks.bas _ Test_9_14_SearchFallback ______. 9.13 .
'    ________________________. )
' ============================================================
Public Function Debug_NormalizeUrl(ByVal s As String) As String
    Debug_NormalizeUrl = NormalizeUrl(s)
End Function


' ============================================================
' BuildSearchUrl (_9.14 -> _9.16, __________, _A/_E-1/_P-1)
'
'   ___________ URL _____.
'     "__ __"  -> m_searchTemplate & "%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97"
'
'   *_9.16: _____ m_searchTemplate _______* (_1-A)
'     9.14 __ "https://www.google.com/search?q=" ____________,
'     _______ m_searchTemplate ___________________.
'     ___ Google (Class_Initialize ____). Property SearchEngine _____.
' ============================================================
Private Function BuildSearchUrl(ByVal terms As String) As String
    BuildSearchUrl = m_searchTemplate & Wv2Url.UrlEncodeComponent(terms)
End Function


' ============================================================
' SearchEngine (_9.16, _________ URL ______, _1-A + _2-A)
'
'   __________ (NormalizeUrl __3__) _____ URL .
'   _______. ____________________.
'     _: Get -> "https://www.google.com/search?q="
'         Let "https://www.bing.com/search?q=" _ Bing .
'
'   *____* ___ (_______) _________ Google .
'     _________. URL _____________ (.
'     ____________, ___________).
' ============================================================
' ============================================================
' ============================================================
' NewWindowMode (_9.27, NewWindowRequested _______ 3 _____)
'
'   0 = ____ (LEGACY)   : args.get_Uri _ URL ____, put_Handled(TRUE) .
'                             _______ Navigate ___. Navigate _*__ GET*
'                             ___, POST ____ 405 ___ (___ 9.26 ____).
'   1 = ________     : put_Handled _______. .
'                             _______________ (____ 23 _____).
'                             POST ______________, _______.
'   2 = ________     : *9.27 ___* __________ CoreWebView2 .
'                             args.put_NewWindow (vtable 4) ___. .
'                             __ WebView ______________ POST ____,
'                             ___________.
'
'   *______ 0 ____* _____________________,
'     _________________________.
'   *_____* Wv2Pane ____ m_browser.NewWindowMode _____, .
'     __________ (____________).
'   *___ = 2 (_9.28 _ 0 ____)* Class_Initialize ___. .
'     PREWARM_DELAY_MS ___________, .
'     ______________ (*_________).
' ============================================================
Public Property Get NewWindowMode() As Long
    NewWindowMode = m_newWindowMode
End Property

Public Property Let NewWindowMode(ByVal value As Long)
    Dim v As Long
    v = value
    If v < 0 Or v > 2 Then
        Debug.Print "Wv2Browser.NewWindowMode: _____ " & value & " _ 0 ____"
        v = 0
    End If
    m_newWindowMode = v

    Select Case v
        Case 0
            Debug.Print "Wv2Browser.NewWindowMode = 0 (____: put_Handled + ______. POST _ GET ____)"
        Case 1
            Debug.Print "Wv2Browser.NewWindowMode = 1 (________: _______. POST _____________)"
        Case 2
            Debug.Print "Wv2Browser.NewWindowMode = 2 (________: put_NewWindow _________)"
            Debug.Print "    _______: " & Debug_PrewarmStatus()
            If Not PrewarmAvailable Then
                Debug.Print "    *_________________ (__________________)*"
                RequestPrewarmAsync
            End If
    End Select
End Property


' ============================================================
' SetHandledWithNewWindow (_9.27, __5 _______)
'
'   put_NewWindow ______ CoreWebView2 ________, .
'   put_Handled(TRUE) ________. __________________ True.
'
'   *________* _NewWindow _______ Handled ___ / .
'     ________________________. __ .
'     ________________, ___ False .
'     *________* _______________.
' ============================================================
Public Property Get SetHandledWithNewWindow() As Boolean
    SetHandledWithNewWindow = m_setHandledWithNewWindow
End Property

Public Property Let SetHandledWithNewWindow(ByVal value As Boolean)
    m_setHandledWithNewWindow = value
    If value Then
        Debug.Print "Wv2Browser.SetHandledWithNewWindow = True (put_NewWindow ___ put_Handled ____)"
    Else
        Debug.Print "Wv2Browser.SetHandledWithNewWindow = False (put_NewWindow ____)"
    End If
End Property


' ============================================================
' PopupDelegateMode (_9.26 -> _9.27 _______)
'
'   9.26 ____ Boolean ____. 9.27 _ NewWindowMode (3 _) _______,
'   *___ NewWindowMode _____________* ____.
'   _____ Test_9_26_Popup_On / _Off _ 9.27 ________.
'
'   Get : NewWindowMode = 1 (________) ___ True
'   Let : True -> NewWindowMode = 1 / False -> NewWindowMode = 0
'         _ False _____ *________ (2) ______* _____.
'            9.27 _________ NewWindowMode = 2 _______.
' ============================================================
Public Property Get PopupDelegateMode() As Boolean
    PopupDelegateMode = (m_newWindowMode = 1)
End Property

Public Property Let PopupDelegateMode(ByVal value As Boolean)
    If value Then
        NewWindowMode = 1
    Else
        NewWindowMode = 0
    End If
End Property


Public Property Get SearchEngine() As String
    SearchEngine = m_searchTemplate
End Property

Public Property Let SearchEngine(ByVal template As String)
    Dim t As String
    t = Trim$(template)
    If Len(t) = 0 Then
        m_searchTemplate = DEFAULT_SEARCH_TEMPLATE
        Debug.Print "Wv2Browser.SearchEngine: ____________ (Google) ________"
    Else
        m_searchTemplate = t
        Wv2Log.LogD "Wv2Browser.SearchEngine: ___________ = " & t
    End If

    ' --- K-2: ____________________ ---
    RequestSettingsSyncAsync
End Property


' ============================================================
' Debug_SearchEngine (_9.16, _______)
'
'   ___ m_searchTemplate __________. BuildSearchUrl _ Private ___,
'   ___ (Test_9_16_SearchEngine) ____________________.
'   __ BuildSearchUrl ________ m_searchTemplate _________,
'   ________________ (Debug_NormalizeUrl _______).
' ============================================================
Public Function Debug_SearchEngine() As String
    Debug_SearchEngine = m_searchTemplate
End Function



' ============================================================
' ResizeAll ? ____ Bounds .
'   _9.10c-1 _____. UserForm_Resize __ 1 ____ (___ a).
'
'   _______ Bounds __________: .
'   _______________. __/_______________.
' ============================================================
Public Sub ResizeAll(ByRef rc As RECT)
    ' --- _9.11c: ____________________ (_A) ---
    '   ___ AddTab ________ Bounds ___, __________.
    m_lastContentRect = rc
    m_hasContentRect = True

    Dim i As Long
    Dim pane As Wv2Pane
    Dim hr As Long
    For i = 1 To m_tabs.Count
        Set pane = m_tabs.Item(i)
        If pane.IsReady Then
            hr = pane.Ctrl_PutBounds(rc)
            If hr <> 0 Then
                Debug.Print "Wv2Browser.ResizeAll: __" & i & " _ PutBounds __ hr=&H" & Hex(hr)
            End If
        End If
    Next i
End Sub


' ============================================================
' PaneAt ? index (1-based) ________ (___)
' ============================================================
Public Property Get PaneAt(ByVal index As Long) As Wv2Pane
    If index < 1 Or index > m_tabs.Count Then
        Set PaneAt = Nothing
        Exit Property
    End If
    Set PaneAt = m_tabs.Item(index)
End Property


' ============================================================
' state ? ___ Browser .
' ============================================================
Public Property Get state() As BrowserState
    state = m_state
End Property


' ============================================================
' ___________ (private, _________________)
' ============================================================
Private Function WaitEnvReady() As Boolean
    Dim t As Single
    t = Timer
    Do While (Not m_env.IsReady) And (Not m_env.IsFailed) And ((Timer - t) < READY_TIMEOUT)
        DoEvents
    Loop
    WaitEnvReady = m_env.IsReady
End Function

Private Function WaitPaneReady(ByVal pane As Wv2Pane) As Boolean
    Dim t As Single
    t = Timer
    Do While (Not pane.IsReady) And (Not pane.IsFailed) And ((Timer - t) < READY_TIMEOUT)
        DoEvents
    Loop
    WaitPaneReady = pane.IsReady
End Function


' ============================================================
' RequestNewTabAsync (_9.10e, __: NewWindowRequested ________)
'
'   Wv2Pane.View_OnNewWindowRequested ______. _____ AddTabWithUrl .
'   ________ (____ 14) __, URL .
'   SafeTimer ____, ________ m_newTabTimer_TimerCall ______.
'
'   *_____ New SafeTimer ____* SafeTimer.StartTimer .
'   _uElapse = m_uElapse ______, __ (StartTimer 0) __ m_uElapse ___.
'   ________________________________, .
'   _______ (m_uElapse=0) ____________. .
'   ___________________ URL ____________.
' ============================================================
Public Sub RequestNewTabAsync(ByVal url As String)
    If m_pendingUrls Is Nothing Then Set m_pendingUrls = New Collection
    m_pendingUrls.Add url

    Set m_newTabTimer = New SafeTimer      ' ____________ (_____)
    m_newTabTimer.StartTimer 1             ' 1ms _ (_______) _ TimerCall

    Debug.Print "Wv2Browser.RequestNewTabAsync: __ URL __ (" & url & _
                "). ____=" & m_pendingUrls.Count & ", ______ Timer __"
End Sub


' ============================================================
' RequestSettingsTabAsync (_9.18, ___________)
'
'   TabBar _ [__] ___ -> OnPaneWebMessage "settings" ______. .
'   ________ AddTab -> WaitPaneReady _____, WebMessage .
'   ________ ____ 14 ___________. ___ newtab ____,
'   _______________ SafeTimer _________. .
'   m_settingsTimer_TimerCall _____________.
'
'   *_____ New SafeTimer _____ RequestNewTabAsync ___*
'     StartTimer __________________ m_uElapse _____,
'     ______________. ____________________.
'   *_______* ___________________. TimerCall .
'     ______ 1 _____ (____________________).
' ============================================================
Public Sub RequestSettingsTabAsync()
    m_settingsPending = True

    Set m_settingsTimer = New SafeTimer    ' ____________ (_____)
    m_settingsTimer.StartTimer 1           ' 1ms _ (_______) _ TimerCall

    Debug.Print "Wv2Browser.RequestSettingsTabAsync: _________ (______ Timer __)"
End Sub


' ============================================================
' m_settingsTimer_TimerCall (_9.18, __________________)
'
'   WebMessage _______ return _____________. .
'   OpenSettingsTab ____ AddTab _ WaitPaneReady .
'   (newtab _ TimerCall _____).
' ============================================================
Private Sub m_settingsTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' _______: .
    If Not m_settingsTimer Is Nothing Then m_settingsTimer.StartTimer 0

    If Not m_settingsPending Then Exit Sub
    m_settingsPending = False

    Debug.Print "Wv2Browser.TimerCall(settings): _______"
    OpenSettingsTab
End Sub


' ============================================================
' IsSettingsView (K-2)
'
'   __ Pane _ *_* _____________, _________.
'   _____ HTML __ window.__wv2Settings = true ______.
'
'   *________ Elem_CheckStale _____* .
'     InStr ___. EvalSync _______ JSON _______,
'     ________________________.
'
'   *_________ False (___)* ______Close ___in-callback .
'     ________________. ____________________,
'     ____________________ (____93 _ LastEvalOk ___).
' ============================================================
Private Function IsSettingsView(ByVal pane As Wv2Pane) As Boolean
    If pane Is Nothing Then Exit Function

    Dim js As String
    Dim res As String
    js = "(function(){try{return {s:(window.__wv2Settings===true)?'yes':'no'};}" & _
         "catch(e){return {s:'no'};}})()"

    res = pane.EvalSync(js, 2)

    If Not pane.LastEvalOk Then
        Wv2Log.LogD "Wv2Browser.IsSettingsView: ___________________ (" & _
                    pane.LastEvalError & ")"
        Exit Function
    End If

    IsSettingsView = (InStr(1, res, """s"":""yes""") > 0)
End Function


' ============================================================
' RequestSettingsSyncAsync / m_syncTimer_TimerCall (K-2)
'
'   *________* SetEngine / SetCustomEngine _ hostObjects .
'   (JS->VBA) _____. ____27 ___, ___ *_________ JS .
'   ____________*. ____ EvalSync .
'   ___________. SafeTimer .
'   (RequestSettingsTabAsync / RequestNewTabAsync ____).
' ============================================================
Public Sub RequestSettingsSyncAsync()
    m_syncPending = True

    Set m_syncTimer = New SafeTimer
    m_syncTimer.StartTimer 1
End Sub


Private Sub m_syncTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' _______: .
    If Not m_syncTimer Is Nothing Then m_syncTimer.StartTimer 0

    If Not m_syncPending Then Exit Sub
    m_syncPending = False

    SyncSettingsViewsNow
End Sub


' ============================================================
' SyncSettingsViewsNow (K-2, __________)
'
'   _____________, _____________.
'   *____60 (VBA->JS _____)* ______, __________.
'
'   _______, _______________ 2 _______.
'   _____ OpenSettingsTab _______, _____________.
'   ________________________.
'
'   *_________________* _________________.
'     _________________.
' ============================================================
Private Sub SyncSettingsViewsNow()
    If m_tabs Is Nothing Then Exit Sub
    If m_tabs.Count = 0 Then Exit Sub

    ' .
    If m_settingsBridge Is Nothing Then Exit Sub

    Dim syncName As String
    syncName = m_settingsBridge.GetEngine()
    If Len(syncName) = 0 Then Exit Sub

    Dim syncIdx As Long
    Dim syncPane As Wv2Pane
    Dim syncJs As String
    Dim res As String
    Dim hit As Long

    For syncIdx = 1 To m_tabs.Count
        Set syncPane = m_tabs.Item(syncIdx)
        If IsSettingsView(syncPane) Then
            syncJs = "(function(){try{return {r:applyFromHost(" & _
                     syncPane.JsQuote(syncName) & ")};}catch(e){return {r:'err'};}})()"
            res = syncPane.EvalSync(syncJs, 2)
            If syncPane.LastEvalOk Then
                If InStr(1, res, """r"":""ok""") > 0 Then
                    hit = hit + 1
                Else
                    Wv2Log.LogW "Wv2Browser.SyncSettingsViewsNow: index " & syncIdx & _
                                " ______ (" & res & ")"
                End If
            Else
                Wv2Log.LogW "Wv2Browser.SyncSettingsViewsNow: index " & syncIdx & _
                            " _______ (" & syncPane.LastEvalError & ")"
            End If
        End If
    Next syncIdx

    If hit > 0 Then
        Wv2Log.LogD "Wv2Browser.SyncSettingsViewsNow: _____ " & hit & _
                    " __ " & syncName & " _____"
    End If
End Sub

' ============================================================
' OpenSettingsTab (_9.18, _____________ attach __)
'
'   __ (__A: attach ____ Navigate):
'     1. AddTab ____ (Pane + Controller) _____.
'     2. Wv2SettingsBridge ____, __ (Browser) _ BindBrowser ______.
'     3. __ Pane _ View_AddHostObjectToScript("Settings", bridge) _ attach.
'        *NavigateToString ___ attach __* ______ JS .
'        hostObjects.Settings ____, _____ (__A).
'     4. BuildSettingsHtml _____ HTML ___, View_NavigateToString ___.
'     5. _______________ (ActivateTab).
'
'   *_______* attach __________________ WebView2 .
'     _____. Browser _ m_settingsBridge _ 1 ___, .
'     (___________ Sub _____________). .
'     _____________________, ____ WebView2 .
'     attach ____________, ______________.
' ============================================================
Public Sub OpenSettingsTab()
    ' --- _9.19: __________, _____________ ---
    '   m_tabs ________ IsSettings = True _ Pane ___. .
    '   __ index _________ Exit Sub (_______). .
    '   _____ m_tabs _________, ______________.
    ' --- K-2: _______ *________* ---
    '   IsSettings ______ False ____, ____________.
    '   EvalSync _ window.__wv2Settings _____, .
    '   _________________ (_____________).
    Dim j As Long
    For j = 1 To m_tabs.Count
        If IsSettingsView(m_tabs.Item(j)) Then
            Wv2Log.LogD "Wv2Browser.OpenSettingsTab: ___________(index " & j & _
                        ")_____________"
            ActivateTab j
            Exit Sub
        End If
    Next j

    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.OpenSettingsTab: AddTab __"
        Exit Sub
    End If

    ' --- _9.19: ____ Pane _____________ ---
    '   ___ [__] __________________.
    pane.IsSettings = True

    ' _________ Browser .
    Set m_settingsBridge = New Wv2SettingsBridge
    m_settingsBridge.BindBrowser Me

    ' *__A: Navigate ___ attach*
    Dim hr As Long
    hr = pane.View_AddHostObjectToScript("Settings", m_settingsBridge)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.OpenSettingsTab: AddHostObjectToScript __ hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.OpenSettingsTab: Settings _____ attach OK"
    End If

    ' ____ HTML .
    hr = pane.View_NavigateToString(BuildSettingsHtml())
    If hr <> 0 Then
        Debug.Print "Wv2Browser.OpenSettingsTab: NavigateToString __ hr=&H" & Hex(hr)
    End If

    ' _______________ (_ index ____)
    Dim i As Long
    For i = 1 To m_tabs.Count
        If m_tabs.Item(i) Is pane Then
            ActivateTab i
            Exit For
        End If
    Next i
End Sub


' ============================================================
' BuildSettingsHtml (_9.18, _____ HTML ___)
'
'   ____________ UI. hostObjects.sync.Settings ___ VBA .
'   __ (postMessage _________):
'     ____          : Settings.EngineList() _____________.
'                         Settings.GetEngine() _______________.
'     ________  : Settings.SetEngine(id) ___. ________.
'     ____ hover    : Settings.PreviewUrlFor(id, _____) ___ URL .
'                         _____ (_________).
'
'   *__________ (CRLF/UTF-8/BOM_______ Python ___)*
'   *hostObjects.sync _ WebView2 ________________*
'     ________ JS _______ (________).
' ============================================================
Private Function BuildSettingsHtml() As String
    Dim s As String
    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<title>__</title>" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  *{box-sizing:border-box;margin:0;padding:0;}" & vbLf
    s = s & "  html,body{height:100%;font-family:'Segoe UI','Meiryo',sans-serif;" & _
            "color:#e8eaed;background:radial-gradient(1200px 600px at 20% -10%,#2b3550 0%,#151a26 55%,#0e121b 100%);}" & vbLf
    s = s & "  .wrap{max-width:840px;margin:0 auto;padding:48px 28px 64px;}" & vbLf
    s = s & "  .eyebrow{letter-spacing:.22em;font-size:11px;color:#8ea2c8;text-transform:uppercase;}" & vbLf
    s = s & "  h1{font-size:26px;font-weight:650;margin:6px 0 4px;}" & vbLf
    s = s & "  .lead{color:#9aa7bd;font-size:13.5px;margin-bottom:28px;}" & vbLf
    s = s & "  .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px;}" & vbLf
    s = s & "  .card{position:relative;border:1px solid rgba(255,255,255,.08);border-radius:14px;" & _
            "padding:18px 18px 16px;cursor:pointer;background:rgba(255,255,255,.03);" & _
            "transition:transform .12s ease,border-color .12s ease,background .12s ease;overflow:hidden;}" & vbLf
    s = s & "  .card:hover{transform:translateY(-2px);border-color:rgba(120,160,255,.5);background:rgba(255,255,255,.06);}" & vbLf
    s = s & "  .card.sel{border-color:#6ea8fe;background:linear-gradient(180deg,rgba(110,168,254,.16),rgba(110,168,254,.05));" & _
            "box-shadow:0 0 0 1px rgba(110,168,254,.35),0 10px 30px -12px rgba(110,168,254,.5);}" & vbLf
    s = s & "  .card .dot{position:absolute;top:16px;right:16px;width:16px;height:16px;border-radius:50%;" & _
            "border:2px solid rgba(255,255,255,.3);}" & vbLf
    s = s & "  .card.sel .dot{border-color:#6ea8fe;background:#6ea8fe;box-shadow:0 0 10px #6ea8fe;}" & vbLf
    s = s & "  .card .name{font-size:17px;font-weight:620;margin-bottom:2px;}" & vbLf
    s = s & "  .card .host{font-size:11.5px;color:#8ea2c8;word-break:break-all;}" & vbLf
    s = s & "  .badge{display:inline-block;margin-top:10px;font-size:10.5px;letter-spacing:.08em;" & _
            "color:#6ea8fe;border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;opacity:0;" & _
            "transition:opacity .12s ease;}" & vbLf
    s = s & "  .card.sel .badge{opacity:1;}" & vbLf
    s = s & "  .preview{margin-top:26px;border-top:1px solid rgba(255,255,255,.08);padding-top:16px;}" & vbLf
    s = s & "  .preview .plabel{font-size:11px;letter-spacing:.14em;color:#8ea2c8;text-transform:uppercase;margin-bottom:6px;}" & vbLf
    s = s & "  .preview .purl{font-family:'Consolas','Courier New',monospace;font-size:12.5px;color:#cfe0ff;" & _
            "background:rgba(0,0,0,.25);border-radius:8px;padding:10px 12px;word-break:break-all;min-height:38px;}" & vbLf
    s = s & "  .sample{color:#7d8aa0;}" & vbLf
    ' --- _9.29: _______ (_______________) ---
    s = s & "  .custom{margin-top:22px;border:1px solid rgba(255,255,255,.08);border-radius:14px;" & _
            "padding:16px 18px;background:rgba(255,255,255,.03);}" & vbLf
    s = s & "  .custom.sel{border-color:#6ea8fe;" & _
            "background:linear-gradient(180deg,rgba(110,168,254,.16),rgba(110,168,254,.05));" & _
            "box-shadow:0 0 0 1px rgba(110,168,254,.35);}" & vbLf
    s = s & "  .chead{display:flex;align-items:center;gap:10px;}" & vbLf
    s = s & "  .cname{font-size:15px;font-weight:620;}" & vbLf
    s = s & "  .cbadge{font-size:10.5px;letter-spacing:.08em;color:#6ea8fe;" & _
            "border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;" & _
            "opacity:0;transition:opacity .12s ease;}" & vbLf
    s = s & "  .custom.sel .cbadge{opacity:1;}" & vbLf
    s = s & "  .chint{font-size:11.5px;color:#8ea2c8;margin:6px 0 12px;}" & vbLf
    s = s & "  .crow{display:flex;gap:10px;}" & vbLf
    s = s & "  .crow input{flex:1;min-width:0;font-family:'Consolas','Courier New',monospace;" & _
            "font-size:12.5px;color:#e8eaed;background:rgba(0,0,0,.25);" & _
            "border:1px solid rgba(255,255,255,.12);border-radius:8px;padding:9px 11px;outline:none;}" & vbLf
    s = s & "  .crow input:focus{border-color:#6ea8fe;}" & vbLf
    s = s & "  .crow button{font-size:13px;color:#0e121b;background:#6ea8fe;border:0;" & _
            "border-radius:8px;padding:9px 18px;cursor:pointer;font-weight:620;}" & vbLf
    s = s & "  .crow button:hover{background:#8fbcff;}" & vbLf
    s = s & "  .cmsg{margin-top:10px;font-family:'Consolas','Courier New',monospace;font-size:12px;" & _
            "color:#a9c6ff;word-break:break-all;min-height:16px;}" & vbLf
    s = s & "  .cmsg.err{color:#ff9d9d;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "<div class=""wrap"">" & vbLf
    s = s & "  <div class=""eyebrow"">Search Engine</div>" & vbLf
    s = s & "  <h1>_________</h1>" & vbLf
    s = s & "  <div class=""lead"">_______ URL ____________, _________________. " & _
            "____________, _______ URL ___________. </div>" & vbLf
    s = s & "  <div class=""grid"" id=""grid""></div>" & vbLf
    ' --- _9.29: _______. form _____ (Enter _ submit _____) ---
    s = s & "  <div class=""custom"" id=""custom"">" & vbLf
    s = s & "    <div class=""chead""><span class=""cname"">____ (_______)</span>" & _
            "<span class=""cbadge"">___</span></div>" & vbLf
    s = s & "    <div class=""chint"">_____________ URL ______. " & _
            "_: https://example.com/search?q=</div>" & vbLf
    s = s & "    <div class=""crow"">" & vbLf
    s = s & "      <input id=""ctpl"" type=""text"" spellcheck=""false"" " & _
            "placeholder=""https://example.com/search?q="">" & vbLf
    s = s & "      <button id=""capply"">__</button>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "    <div class=""cmsg"" id=""cmsg""></div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  <div class=""preview"">" & vbLf
    s = s & "    <div class=""plabel"">Preview</div>" & vbLf
    s = s & "    <div class=""purl"" id=""purl""><span class=""sample"">____________________</span></div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "</div>" & vbLf
    s = s & "<script>" & vbLf
    ' --- hostObjects.sync.Settings ____. _______________ ---
    s = s & "  var S = (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects" & _
            " && window.chrome.webview.hostObjects.sync) ? window.chrome.webview.hostObjects.sync.Settings : null;" & vbLf
    s = s & "  var SAMPLE = '__ __';" & vbLf
    s = s & "  var current = '';" & vbLf
    ' --- _9.19c: ________________ id (_ hover _ null) ---
    s = s & "  var hoverId = null;" & vbLf
    ' --- K-2: __________________ ---
    '   OpenSettingsTab _ EvalSync ____. *_______*
    '   ___ JS _____, .
    '   (___________________).
    s = s & "  window.__wv2Settings = true;" & vbLf
    ' --- _____________ VBA ______ ---
    s = s & "  function loadState(){" & vbLf
    s = s & "    if(!S){ document.getElementById('grid').textContent='____________'; return null; }" & vbLf
    s = s & "    try{ current = S.GetEngine(); }catch(e){ current=''; }" & vbLf
    s = s & "    try{ return JSON.parse(S.EngineList()); }catch(e){ return []; }" & vbLf
    s = s & "  }" & vbLf
    ' --- _____ URL ____________ (____) ---
    s = s & "  function hostOf(tpl){ try{ return (tpl.split('/')[2]||tpl); }catch(e){ return tpl; } }" & vbLf
    s = s & "  function markSel(){" & vbLf
    s = s & "    var cs=document.querySelectorAll('.card');" & vbLf
    s = s & "    for(var i=0;i<cs.length;i++){ cs[i].classList.toggle('sel', cs[i].getAttribute('data-id')===current); }" & vbLf
    ' --- _9.29: _________________ 1 ________ ---
    s = s & "    markCustom();" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function showPreview(id){" & vbLf
    s = s & "    var el=document.getElementById('purl'); if(!S){ return; }" & vbLf
    s = s & "    var u=''; try{ u=S.PreviewUrlFor(id, SAMPLE); }catch(e){ u=''; }" & vbLf
    s = s & "    el.innerHTML = '<span class=""sample"">' + SAMPLE + '</span> &nbsp;&rarr;&nbsp; ' + (u||'(____)');" & vbLf
    s = s & "  }" & vbLf
    ' --- _9.19b: __________________ (_ hover _/_______) ---
    s = s & "  function showCurrentPreview(){" & vbLf
    s = s & "    if(!S){ return; }" & vbLf
    s = s & "    if(current){ showPreview(current); }" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function build(list){" & vbLf
    s = s & "    var g=document.getElementById('grid'); g.innerHTML='';" & vbLf
    s = s & "    for(var i=0;i<list.length;i++){ (function(item){" & vbLf
    s = s & "      var c=document.createElement('div'); c.className='card'; c.setAttribute('data-id',item.id);" & vbLf
    s = s & "      c.innerHTML='<span class=""dot""></span>' +" & vbLf
    s = s & "        '<div class=""name"">'+item.label+'</div>' +" & vbLf
    s = s & "        '<div class=""host"">'+hostOf(item.template)+'</div>' +" & vbLf
    s = s & "        '<div class=""badge"">___</div>';" & vbLf
    ' --- _9.19c: hover __ hoverId ___, ______________ ---
    s = s & "      c.onmouseenter=function(){ hoverId=item.id; showPreview(item.id); };" & vbLf
    ' --- _9.19c: ____, _______________________ ---
    s = s & "      c.onmouseleave=function(){ hoverId=null; setTimeout(function(){ if(hoverId===null){ showCurrentPreview(); } },0); };" & vbLf
    s = s & "      c.onclick=function(){" & vbLf
    s = s & "        if(!S) return;" & vbLf
    s = s & "        var res=''; try{ res=S.SetEngine(item.id); }catch(e){ res=''; }" & vbLf
    s = s & "        if(res){ current=res; } else { current=item.id; }" & vbLf
    s = s & "        markSel(); showPreview(item.id);" & vbLf
    s = s & "      };" & vbLf
    s = s & "      g.appendChild(c);" & vbLf
    s = s & "    })(list[i]); }" & vbLf
    s = s & "    markSel();" & vbLf
    ' --- _9.19c: _____ (_ hover) _________________ ---
    '   (___________________ mouseleave ___. #grid .
    '    mouseleave _____. ____________________. )
    s = s & "    showCurrentPreview();" & vbLf
    s = s & "  }" & vbLf
    ' --- _9.29: ________ JS (________________) ---
    s = s & "  function markCustom(){" & vbLf
    s = s & "    var box=document.getElementById('custom'); if(!box){ return; }" & vbLf
    s = s & "    box.classList.toggle('sel', current==='custom');" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function setMsg(text, isErr){" & vbLf
    s = s & "    var m=document.getElementById('cmsg'); if(!m){ return; }" & vbLf
    s = s & "    m.textContent = text || '';" & vbLf
    s = s & "    m.className = isErr ? 'cmsg err' : 'cmsg';" & vbLf
    s = s & "  }" & vbLf
    ' --- ____________. ______ VBA (_____) ____ ---
    s = s & "  function customPreview(){" & vbLf
    s = s & "    var ct=document.getElementById('ctpl'); if(!ct||!S){ return; }" & vbLf
    s = s & "    var v=ct.value;" & vbLf
    s = s & "    if(!v){ setMsg('', false); return; }" & vbLf
    s = s & "    var u=''; try{ u=S.PreviewUrlForTemplate(v, SAMPLE); }catch(e){ u=''; }" & vbLf
    s = s & "    if(u){ setMsg(SAMPLE + '  ->  ' + u, false); }" & vbLf
    s = s & "    else { setMsg('http:// ___ https:// ____ URL _________', true); }" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function applyCustom(){" & vbLf
    s = s & "    var ct=document.getElementById('ctpl'); if(!ct||!S){ return; }" & vbLf
    s = s & "    var res=''; try{ res=S.SetCustomEngine(ct.value); }catch(e){ res=''; }" & vbLf
    s = s & "    if(!res){ setMsg('__________. http:// ___ https:// ____ URL _" & _
            "________', true); return; }" & vbLf
    s = s & "    current=res; markSel(); showCurrentPreview();" & vbLf
    s = s & "    if(res==='custom'){ setMsg('______ (____)', false); }" & vbLf
    s = s & "    else { setMsg('________ (' + res + ') ________, " & _
            "__________', false); }" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function initCustom(){" & vbLf
    s = s & "    var ct=document.getElementById('ctpl'); if(!ct){ return; }" & vbLf
    s = s & "    if(S){ try{ ct.value = S.GetCustomTemplate(); }catch(e){ } }" & vbLf
    s = s & "    ct.oninput = function(){ customPreview(); };" & vbLf
    s = s & "    ct.onkeydown = function(e){ if(e.key==='Enter'){ applyCustom(); } };" & vbLf
    s = s & "    var btn=document.getElementById('capply');" & vbLf
    s = s & "    if(btn){ btn.onclick=function(){ applyCustom(); }; }" & vbLf
    s = s & "    markCustom();" & vbLf
    s = s & "  }" & vbLf
    ' --- K-2: ______________ ---
    '   __________________, VBA _____________.
    '   _______ (ctpl) _________________.
    s = s & "  function applyFromHost(name){" & vbLf
    s = s & "    if(!name){ return 'ng'; }" & vbLf
    s = s & "    current = name;" & vbLf
    s = s & "    markSel();" & vbLf
    s = s & "    showCurrentPreview();" & vbLf
    s = s & "    return 'ok';" & vbLf
    s = s & "  }" & vbLf
    s = s & "  var list=loadState(); if(list){ build(list); }" & vbLf
    s = s & "  initCustom();" & vbLf
    s = s & "</script>" & vbLf
    s = s & "</body></html>"
    BuildSettingsHtml = s
End Function


' ============================================================
' ============================================================
' ________ (_9.27) _ _____ 1 .
'
'   *_______* NewWindowRequested ___________ Controller .
'     _____, ____ (ControllerCompleted) ______ return .
'     ________________ (____ 14). Deferral __________,
'     args / deferral _ COM _______________________.
'     *___________, ______ put_NewWindow ________* __ .
'     ___________. _____________ 1 ______.
'
'   *_____ m_tabs _____* .
'     _________ (TabAdded _________)
'     _TabCount / ActivateTab / CloseTab / MoveTab .
'     _NotifyTitleChanged / NotifyUrlChanged ________ (__).
'   __ (AdoptPrewarmedNow) _______ m_tabs ___, ________.
'
'   *_____* m_tabs _________ Class_Terminate .
'     ____. Shutdown _____ ClearBrowserRef -> Set Nothing __ (____ 58).
' ============================================================

' --- _________________ ---
'   ____ / Controller Ready / View ____ / __ Claim ______, _ 4 __.
Public Property Get PrewarmAvailable() As Boolean
    If m_prewarmPane Is Nothing Then Exit Property
    If m_prewarmClaimed Then Exit Property
    If Not m_prewarmPane.IsReady Then Exit Property
    PrewarmAvailable = (m_prewarmPane.ViewPtr <> 0)
End Property


' ============================================================
' ClaimPrewarmViewPtr (_9.27, _____ View ____________)
'
'   Wv2Pane.View_OnNewWindowRequested ______. ___________ ViewPtr .
'   __, ___ m_prewarmClaimed ____*________*.
'
'   *__ Claim ____* put_NewWindow ________ Timer .
'     __ NewWindowRequested ____, __ WebView _ 2 ___________.
'     Claim ____ 0 ___, 2 __________________.
'
'   ___ 0 = ____ (___ / _ Ready / __ Claim __).
' ============================================================
Public Function ClaimPrewarmViewPtr() As LongPtr
    If Not PrewarmAvailable Then
        Debug.Print "Wv2Browser.ClaimPrewarmViewPtr: _________ -> " & Debug_PrewarmStatus()
        Exit Function
    End If
    ' *____________* VBA __ Function .
    '   _________. _______________.
    Dim pv As LongPtr
    pv = m_prewarmPane.ViewPtr
    m_prewarmClaimed = True
    ClaimPrewarmViewPtr = pv
    Debug.Print "Wv2Browser.ClaimPrewarmViewPtr: _______ (ViewPtr=" & pv & ")"
End Function


' --- Claim _____ (put_NewWindow ____________) ---
Public Sub ReleasePrewarmClaim()
    If Not m_prewarmClaimed Then Exit Sub
    m_prewarmClaimed = False
    Debug.Print "Wv2Browser.ReleasePrewarmClaim: ________ (__________)"
End Sub


' ============================================================
' RequestPrewarmAsync / m_prewarmTimer_TimerCall (_9.27, ____________)
'
'   *__________ (__1 _b)* Init ___ Environment -> _____ ->
'     TabBar -> NavBar _ Controller ________________, ____ 20 .
'     _______. ___ 4 ________, _________.
'
'   *_____* CreatePaneCore _ WaitPaneReady _ DoEvents _____, .
'     __ TimerCall __________. _____*______* (.
'     ________ _ ____ SafeTimer ______________).
'     _____ AddTab ___ EnsurePrewarmScheduled _____________,
'     _________________.
' ============================================================
Public Sub RequestPrewarmAsync()
    If Not m_prewarmPane Is Nothing Then
        ' *_9.28: ________* _____________Claim __ (____).
        '   _____________, 9.27 _____________.
        '   *_____________ (___ 9.27 ___)* Claim ____ 1ms .
        '   __ Timer _ AdoptPrewarmedNow -> CreatePrewarmPaneNow _________.
        '   _____ Timer ____, 1.5 _____________ (_______).
        '   _________ 1 _______, ____________.
        If m_prewarmClaimed Then
            Debug.Print "Wv2Browser.RequestPrewarmAsync: _____ Claim __ (____). " & _
                        "_______________, _________"
        Else
            Debug.Print "Wv2Browser.RequestPrewarmAsync: ___________________"
        End If
        Exit Sub
    End If

    m_prewarmPending = True
    Set m_prewarmTimer = New SafeTimer     ' ____________ (_____)
    m_prewarmTimer.StartTimer PREWARM_DELAY_MS

    Debug.Print "Wv2Browser.RequestPrewarmAsync: __________ (" & _
                PREWARM_DELAY_MS & "ms _)"
End Sub


' --- _____, _______________ (AddTab ________) ---
'   *_9.29: ___________* ________ (CreatePrewarmPaneNow ->
'   CreatePaneCore _ WaitPaneReady _ DoEvents ________) _, .
'   ________ AddTab ________________. .
'   m_prewarmPane ___ Nothing (______) _____________,
'   1.5 _____________ (_______)_____________.
'   ___________ 1 ______, __________.
Private Sub EnsurePrewarmScheduled()
    If Not m_prewarmPane Is Nothing Then Exit Sub
    If m_prewarmPending Then Exit Sub

    If m_creatingPane > 0 Then
        Debug.Print "Wv2Browser.EnsurePrewarmScheduled: ________ (___) ___" & _
                    "_____ (______ 1 ____)"
        Exit Sub
    End If

    RequestPrewarmAsync
End Sub


Private Sub m_prewarmTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' _______: .
    If Not m_prewarmTimer Is Nothing Then m_prewarmTimer.StartTimer 0

    If Not m_prewarmPending Then Exit Sub

    ' *_____* Pane ___ (WaitPaneReady _ DoEvents _) .
    If m_creatingPane > 0 Then
        m_prewarmPending = False
        Debug.Print "Wv2Browser.TimerCall(prewarm): Pane ______________ " & _
                    "(AddTab ___________)"
        Exit Sub
    End If

    m_prewarmPending = False
    CreatePrewarmPaneNow
End Sub


' ============================================================
' CreatePrewarmPaneNow (_9.27, ________)
'
'   CreatePaneCore _ Pane ___, *Ctrl_PutIsVisible False ______* (__3).
'   Bounds _ CreatePaneCore __________________, .
'   ActivateTab _________________ (_____________).
'
'   *Navigate ___* ___ about:blank ___. put_NewWindow _____,
'     _____________ (POST _______) ___ WebView ____.
' ============================================================
Private Sub CreatePrewarmPaneNow()
    If Not m_prewarmPane Is Nothing Then
        Debug.Print "Wv2Browser.Prewarm: _________ (_______)"
        Exit Sub
    End If
    If m_state <> Bs_EnvReady Then
        Debug.Print "Wv2Browser.Prewarm: Environment __ Ready ________ (state=" & m_state & ")"
        Exit Sub
    End If

    Debug.Print "Wv2Browser.Prewarm: __________"
    Dim pane As Wv2Pane
    Set pane = CreatePaneCore()
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.Prewarm: *__________* (__ AddTab / __________)"
        Exit Sub
    End If

    Dim hr As Long
    hr = pane.Ctrl_PutIsVisible(False)     ' _____ (____________)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.Prewarm: Ctrl_PutIsVisible(False) __ hr=&H" & Hex(hr)
    End If

    Set m_prewarmPane = pane
    m_prewarmClaimed = False
    Debug.Print "Wv2Browser.Prewarm: ________ (ViewPtr=" & pane.ViewPtr & _
                ", _______=" & m_tabs.Count & " <- _______)"
End Sub


' ============================================================
' RequestAdoptPrewarmedAsync / m_adoptTimer_TimerCall (_9.27, _____)
'
'   Wv2Pane _ put_NewWindow ________________. ______ return _,
'   ____ TimerCall .
'     (1) __ : ___ m_tabs ___, TabAdded -> ActivateTab .
'     (2) __ : .
'   _______ (__6: ____. ___ Controller ______________,
'   ______________).
'
'   *__ TimerCall _________* NewWindowRequested _______ return
'     _____, CreatePaneCore _ WaitPaneReady ________ (____ 14 .
'     ______________). m_newTabTimer_TimerCall _____.
' ============================================================
Public Sub RequestAdoptPrewarmedAsync(ByVal url As String)
    m_adoptPending = True
    m_adoptUrl = url

    Set m_adoptTimer = New SafeTimer      ' ____________ (_____)
    m_adoptTimer.StartTimer 1             ' 1ms _ (_______) _ TimerCall

    Debug.Print "Wv2Browser.RequestAdoptPrewarmedAsync: __________ (" & url & ")"
End Sub


Private Sub m_adoptTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' _______: .
    If Not m_adoptTimer Is Nothing Then m_adoptTimer.StartTimer 0

    If Not m_adoptPending Then Exit Sub
    m_adoptPending = False

    AdoptPrewarmedNow          ' (1) __ _ .
    CreatePrewarmPaneNow       ' (2) __ _ ________ (_____)
End Sub


' ============================================================
' AdoptPrewarmedNow (_9.27, ________________)
'
'   *__ m_prewarmPane ____* m_tabs.Add ____ Nothing ________,
'     _______ (CreatePrewarmPaneNow) _________________.
'   *____ ActivateTab ____* CreatePrewarmPaneNow .
'     ActivateTab ____. __________ 1 ____________.
' ============================================================
Private Sub AdoptPrewarmedNow()
    If m_prewarmPane Is Nothing Then
        Debug.Print "Wv2Browser.Adopt: *_______* (_______. put_NewWindow _" & _
                    "_______________)"
        m_prewarmClaimed = False
        Exit Sub
    End If

    Dim pane As Wv2Pane
    Set pane = m_prewarmPane
    Set m_prewarmPane = Nothing
    m_prewarmClaimed = False

    m_tabs.Add pane
    Debug.Print "Wv2Browser.Adopt: ____________ (______=" & m_tabs.Count & _
                ", __ URL=" & m_adoptUrl & ")"

    RaiseEvent TabAdded(m_tabs.Count)
    ActivateTab m_tabs.Count
End Sub


' --- ___: ________ 1 ________ ---
Public Function Debug_PrewarmStatus() As String
    If m_prewarmPane Is Nothing Then
        Debug_PrewarmStatus = "____ (pending=" & CStr(m_prewarmPending) & ")"
        Exit Function
    End If
    Debug_PrewarmStatus = "____ ViewPtr=" & m_prewarmPane.ViewPtr & _
                          " / IsReady=" & CStr(m_prewarmPane.IsReady) & _
                          " / Claimed=" & CStr(m_prewarmClaimed) & _
                          " / ___=" & CStr(PrewarmAvailable)
End Function


' ============================================================
' m_newTabTimer_TimerCall (_9.10e, SafeTimer ____________)
'
'   NewWindowRequested _______ return _____________.
'   ___ AddTabWithUrl ___, _ Controller _ ControllerCompleted .
'   ________________.
'
'   __ StartTimer 0 __________ (_______), __________.
' ============================================================
Private Sub m_newTabTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' _______: .
    If Not m_newTabTimer Is Nothing Then m_newTabTimer.StartTimer 0

    If m_pendingUrls Is Nothing Then Exit Sub
    Debug.Print "Wv2Browser.TimerCall: __ URL ___ (____=" & m_pendingUrls.Count & ")"

    Do While m_pendingUrls.Count > 0
        Dim url As String
        url = m_pendingUrls(1)
        m_pendingUrls.Remove 1

        Dim pane As Wv2Pane
        Set pane = AddTabWithUrl(url)
        If pane Is Nothing Then
            Debug.Print "  RequestNewTab: AddTabWithUrl __ (" & url & ")"
        Else
            Debug.Print "  RequestNewTab: _____ OK (" & url & ", ViewPtr=" & _
                        pane.ViewPtr & ")"
        End If
    Loop
End Sub


' ============================================================
' Shutdown (_9.10e, ________________)
'
'   _9.10e _ Wv2Pane _ Wv2Browser ______ (m_browser) .
'   _____, Browser.m_tabs _ Pane.m_browser _________.
'   VBA _______ GC ___, UserForm ___ Set m_browser = Nothing .
'   _______________ (Class_Terminate ______).
'
'   ___ UserForm_Terminate ___ Shutdown _______, _ Pane .
'   (ClearBrowserRef) __________. _____ Browser->Pane .
'   _____, __ Set m_browser = Nothing _ Class_Terminate .
'   (_ Pane -> Environment _____).
'
'   *UserForm __________ Shutdown ____ 1 ___* (_Y ___).
'   __ (_________).
' ============================================================
' ============================================================
' QuietOnShutdown (K-4c) - _____________ (__ True)
'
'   *False _____________* __________.
'   ____ False ________.
' ============================================================
Public Property Get QuietOnShutdown() As Boolean
    QuietOnShutdown = m_quietOnShutdown
End Property

Public Property Let QuietOnShutdown(ByVal newValue As Boolean)
    m_quietOnShutdown = newValue
    Wv2Log.LogD "Wv2Browser.QuietOnShutdown = " & newValue
End Property


' ============================================================
' QuietAllPanes (K-4c, Private) - _________ about:blank .
'
'   *____* _________ (0 __ 3 __________).
'   *_______* ____________. View_Navigate .
'   m_pView = 0 _________________.
' ============================================================
Private Sub QuietAllPanes()
    Dim pane As Wv2Pane
    Dim n As Long
    Dim t0 As Single

    t0 = Timer

    ' *_____ m_tabs _______* (_9.27) .
    If Not m_prewarmPane Is Nothing Then
        m_prewarmPane.View_Navigate "about:blank"
        n = n + 1
    End If

    If Not m_tabs Is Nothing Then
        For Each pane In m_tabs
            If Not pane Is Nothing Then
                pane.View_Navigate "about:blank"
                n = n + 1
            End If
        Next pane
    End If

    Wv2Log.LogD "Wv2Browser.QuietAllPanes: " & n & " __ about:blank _ (" & _
                Format$(Timer - t0, "0.00") & " _)"
End Sub


Public Sub Shutdown()
    Wv2Log.LogD "Wv2Browser.Shutdown: *__*"   ' K-4a

    ' --- K-4c: *____________* ---
    If m_quietOnShutdown Then QuietAllPanes

    ' _9.10e (Timer): .
    If Not m_newTabTimer Is Nothing Then
        m_newTabTimer.StartTimer 0
        Set m_newTabTimer = Nothing
    End If

    ' _9.18: .
    If Not m_settingsTimer Is Nothing Then
        m_settingsTimer.StartTimer 0
        Set m_settingsTimer = Nothing
    End If
    m_settingsPending = False
    Set m_settingsBridge = Nothing

    ' _9.27: .
    '   *___ m_tabs _______* __, _______________.
    '   _____________ Nothing ________________ (____ 58).
    If Not m_prewarmTimer Is Nothing Then
        m_prewarmTimer.StartTimer 0
        Set m_prewarmTimer = Nothing
    End If
    If Not m_adoptTimer Is Nothing Then
        m_adoptTimer.StartTimer 0
        Set m_adoptTimer = Nothing
    End If
    m_prewarmPending = False
    m_adoptPending = False
    If Not m_prewarmPane Is Nothing Then
        m_prewarmPane.ClearBrowserRef
        Set m_prewarmPane = Nothing
        m_prewarmClaimed = False
        Wv2Log.LogD "Wv2Browser.Shutdown: ____ (______) ___"
    End If

    If m_tabs Is Nothing Then Exit Sub
    Dim pane As Wv2Pane
    For Each pane In m_tabs
        If Not pane Is Nothing Then pane.ClearBrowserRef
    Next pane
    Wv2Log.LogD "Wv2Browser.Shutdown: _ " & m_tabs.Count & " __________"
End Sub


' ============================================================
' Class_Terminate ? _____ -> Environment __ (____ 43: _____)
' ============================================================
Private Sub Class_Terminate()
    Wv2Log.LogD "Wv2Browser.Terminate __ (___=" & m_tabs.Count & ")"

    ' --- 0. _9.27: _________ ---
    '   ___ Shutdown _________. Shutdown ______________.
    If Not m_prewarmPane Is Nothing Then
        m_prewarmPane.ClearBrowserRef
        Set m_prewarmPane = Nothing
        Wv2Log.LogD "Wv2Browser.Terminate: ____ (______) ___"
    End If

    ' --- 1. _ Pane ___ (_ Pane _ Class_Terminate _ View ______) ---
    '   Collection ____________, ______ 0 .
    '   Wv2Pane.Class_Terminate ___. __________.
    Dim i As Long
    For i = m_tabs.Count To 1 Step -1
        m_tabs.Remove i
    Next i
    Set m_tabs = Nothing

    ' --- 2. Environment ___ (_ Pane ___ = _____) ---
    Set m_env = Nothing

    m_hwndHost = 0
    m_state = Bs_Created
    Wv2Log.LogD "Wv2Browser.Terminate __"
End Sub


