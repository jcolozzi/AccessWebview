Attribute VB_Name = "Wv2Pane"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  N-4 __ (_______) ---
'
'   *N-1_N-3 ______. _________*
'   _____ N-4 __ (_________) _______. ____:
'
'     _`GetContent` ______________. .
'     _*N-2 ____ Seek(0) _____* .
'     _*_________* (`NetRespBodyOn`, __ False)
'     _*____________* encoding ___ 32 ____ hex .
'     _*ResponseView _____ AddRef ______* (____________)
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  N-3 __ (_______________) ---
'
'   ___ 1 __*__*____. ____*_____*____.
'   _____ N-3 __ (_________) _______. ____:
'
'     _*_______________* ___ CSS _____. .
'       *__________________* ____ (___________)
'     _______ URI + ________________________ (FIFO).
'       *__ URL _______________, ___________*
'     _*______________* (0 _______________)
'     _*____ GetIterator _ 7, ____ 8* _____ 1 .
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  N-2 __ (_____________) ---
'
'   N-1 ___ (URL / ____ / __) ___, *1 ____________*___.
'   _____ N-2 __ (_________) _______. ____:
'
'     _*_________ OFF* (NetDetailOn = True _ ON)
'     _*___ IStream ___ Seek(0) ___* .
'     _*Cookie / Authorization _______* (_________)
'     _____________ (__ 50 _). ___ 500 .
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  N-1b __ (____________) ---
'
'   _ *____: _________ WebResourceRequested _____*
'     SetVirtualHostNameToFolderMapping _______ (DOCUMENT / fetch / XHR /
'     __ / CSS ____) __________. .
'     _________. *N-1 _______________* _, .
'     _________ (______ 10 __________).
'     ___________________________________.
'
'   _ *NetLogDrain ________ 0 _____ (__3 _Y)*
'     NetLogTotal / NetLogDropped __*_______* ____________,
'     Drain ___________ 0 ______. .
'       NetCaptureStop: ____ (___ 0 _ ...)
'     ___, *10 ________1 ____________* _____.
'     _NetLogDrain ... ____________ (NetRingEmpty). .
'     _NetLogClear ... ____________ (NetRingAlloc)
'     _NetCaptureStart ... NetLogClear _______ 0 .
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  N-1 __ (WebResourceRequested ___) ---
'
'   *_________________ Excel VBA _____*
'   F12 _ Network _____________. N-1 _ URL / ____ / .
'   (________________).
'
'   _ ____ Public .
'     NetCaptureStart(contexts, uriFilter) As Boolean  .
'     NetCaptureStop()  As Boolean                     .
'     NetCaptureOn      As Boolean                     .
'     NetLogDrain()     As Long     ______ Wv2Log .
'     NetLogLine(i)     As String   *____ 1 _____* (___, 1 __)
'     NetLogCount / NetLogTotal / NetLogDropped As Long
'     NetLogCapacity    As Long     ______ (Get/Let)
'     NetLogClear()                 .
'     View_AddWebResourceRequested / View_RemoveWebResourceRequested
'     View_AddWebResourceRequestedFilter / View_RemoveWebResourceRequestedFilter
'     View_OnWebResourceRequested   (ComCallbackHandler ______)
'
'   _ *__________ (__5)*
'     WebResourceRequested ________ / ______________,
'     *__________________*. .
'     _Wv2Log _ Debug.Print .
'     _EvalSync _______ (____14 / 20)
'     args __ URL / ____ / _________________.
'
'   _ *_______________ (__6)*
'     ___ 1 ____________. ________________,
'     D-4 _ SpaProbeDrainLog _______ (___ <- ________).
'     *____________* NetLogDropped ____ (D-6 __3 _____).
'
'   _ *_____ 1 _____________ 1 _____*
'     AddWebResourceRequestedFilter(uri, ctx) ________. .
'     DOCUMENT(1) / XML_HTTP_REQUEST(7) / FETCH(8) _ 3 _ (__4).
'     *_______________ NetCaptureStop ___* (____106 ___.
'     ____________, __ Start ______________).
'
'   _ *vtable index _ SDK __________ (____115)*
'     ____________ add_WebResourceRequested _ 54 .
'     (__ 55, OpenDevToolsWindow _________). ___ tools/vtable.py.
'       ICoreWebView2::add_WebResourceRequested              55
'       ICoreWebView2::remove_WebResourceRequested           56
'       ICoreWebView2::AddWebResourceRequestedFilter         57
'       ICoreWebView2::RemoveWebResourceRequestedFilter      58
'       ...WebResourceRequestedEventArgs::get_Request         3
'       ...WebResourceRequestedEventArgs::get_ResourceContext 7
'       ICoreWebView2WebResourceRequest::get_Uri              3
'       ICoreWebView2WebResourceRequest::get_Method           5
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  K-4a __ (________ Wv2Log ___) ---
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
' --- Wv2Pane.cls  D-6 __ (QuerySelectorAll _____) ---
'
'   _ ____ Public .
'     QuerySelectorAll(selector, maxCount) As Collection
'     LastAllTruncated  As Boolean  ___ QuerySelectorAll .
'     ElementCount      As Long     JS .
'
'   _ *___ Collection ___ (__2)*
'     For Each el In els ____ VBA _____. 0 __ Count = 0 ___.
'     *__________* ____ UBound __________, .
'     _________. *_____________* For Each .
'     _______ (Attribute Item.VB_UserMemId = 0) ____, import.ps1 .
'     _________ (____47) ___________.
'
'   _ *____________ (__3)*
'     __ 200 _. ____ LastAllTruncated = True __________.
'     ________________________.
'
'   _ *____________. _______ (__4)*
'     QuerySelectorAll _ 1 __ 200 _______, ______ D-2 .
'     _____. ______________*___ Wv2Element .
'     stale ___* (____91 / D-2 __3 __________).
'     ___:
'       _ElementCount .
'       _________ LogW .
'       ______ ClearElementRegistry ________ (D-2 ____)
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  D-4b __ (____) ---
'
'   D-4a ____________*________*.
'
'   _ ____ Public .
'     WaitSettled(timeoutSec, stableMs) As Boolean  .
'     AutoWaitAfterAction  As Boolean  *__ False* __________ (opt-in)
'     AddIgnoreNetwork(pattern)        ________ URL .
'     ClearIgnoreNetwork()             .
'     IgnoreSelectors  As String       ________ DOM .
'     LastSettleInfo   As String       ________ (___)
'
'   _ ______ (__4)
'     ____ fetch + XHR _ 0 __, ___ DOM ____ stableMs ________.
'     *________WaitSettled ___________* .
'     __________. _____*____ stableMs ___*.
'     ___ stableMs = 500 / __ 10 _ / _____ 100ms (___________).
'
'   *** WaitSettled _ True _______________________ ***
'     ________. **________________, .
'     ______. ** _______ D-4c _______ (arm) _______.
'     *___ WaitIdle ___ WaitSettled _________* (__8)
'     _idle = ______________________.
'
'   _ ___ (D-2 / D-3 ______, ____93)
'     True                     ... .
'     False + LastEvalOk=True  ... ___________ (*_____*)
'     False + LastEvalOk=False ... __ (in-callback / Close ___)
'     _____ LastSettleInfo ______ (slack / ______ / __).
'
'   _ _____ (__5)
'     *______ VBA ____, ______________*
'     __________________. .
'     __ 1 __________, _______.
'     *__________________* ______________.
'     __ (__________) _____. ____________.
'
'   _ *______________________* (D-4a ___)
'     ____________ (D-4c _ arm-then-act) _______.
'     ________________________, .
'     NavigationCompleted / DOMContentLoaded _____.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  D-4a __ (____ SPA _________) ---
'
'   D-4 ____*____* (___ DOM ____________).
'   D-3 _ WaitFor / WaitGone __________________,
'   **_____________**______. ________ D-4.
'
'   D-4a _____. *_______ (____) _____________*_,
'   _____ (WaitSettled) _______ (Arm _) _ D-4b / D-4c.
'
'   _ ____ Public .
'     SpaProbeState()  As String   ____________ JSON .
'     SpaProbeReset()  As Boolean  ____________ (___)
'     SpaProbeHealthy() As Boolean __ JSON _ h .
'
'   _ ____________ (window.__wv2p)
'     fetchN / xhrN ... *___* _ fetch / XHR __ (in-flight)
'     lastMut       ... ______ DOM _____. .
'     netN / mutN   ... _____ (___)
'     skipNet / skipMut ... _____ (________________)
'
'   _ *____92 ___: _______________________*
'     fetch _ XMLHttpRequest.prototype.open/send __________,
'     *______________________________*
'     (_______ WebDriver BiDi for SeleniumVBA .
'      = ___________). .
'       .
'       _window.fetch .
'       _XHR _ open/send .
'       _MutationObserver .
'     ____, 1 _______*____*. ________ repairs ____.
'
'     *___________ window.fetch____*
'     _____________________, ___________.
'     _______________________ (____ / _______).
'
'   _ *_______ (__2 _A: ____________)*
'     AddScriptToExecuteOnDocumentCreated _____. D-2 .
'     _________________. .
'     *___________________________*
'     _____. D-4c ________ arm .
'     (arm ______________, __________).
'
'   _ .
'     *___________ Pane ________* ____ / ____ /
'     _____ Pane _________ (D-2 ____________).
'
''''''''''''''''''''''''''''''''''' --- Wv2Pane.cls  D-3 __ (___________) ---
'
'   D-3 _ Wv2Pane ______ *Elem_Exec (Friend) 1 ___*.
'   ___________ Wv2Element _ (Value = / Click / SetAttribute) ___,
'   Wv2Pane _ D-2 _ Elem_Read __________________.
'
'     Elem_Read (D-2) ... __ 1 _*__*   __ e ___ JS *_* .
'     Elem_Exec (D-3) ... ___ 1 _*__* __ e ___ JS *_* .
'
'   _ Elem_Exec _ JS _____ (Elem_Read ________)
'     {s:'gen', g:...}            ____ = stale
'     {s:'nil', g:...}            ________ = stale .
'     {s:'ok',  g:..., v:'<rv>'}  __. v _*___________*
'
'     *JS ___ try-catch ___* View_ExecuteScriptSafe __________,
'     ___ EvalSync _ LastEvalError ("TypeError: ..." _) _________.
'     ____________.
'
'   _ _____ (v) .
'     .Value = ______*_____ setter __ ('setter') .
'     ('direct') _*_, SPA ___________________.
'     Wv2Element.LastInfo ____ (D-3 _____).
'
'   _ DOM ____ (D-3b, __2 _C: *_________, _____ opt-in*)
'
'     WaitFor(selector, timeoutSec)  As Wv2Element  .
'     WaitGone(selector, timeoutSec) As Boolean     .
'     ImplicitWaitSec                As Single      *__ 0 = __*
'
'     *__1 ________ (2026-08-22)*
'       raw ExecuteScript _*________ Promise* _____, .
'       ___ resultJson={} ______. ___ ExecuteScript _ Promise _____.
'       ________ JS ______, *VBA _________* (__3 ___).
'       (____ EvalSync _ JSON.stringify(_) _____, .
'        Promise _____ {} ____. JS ______ D-1 _____________)
'
'     *___________/___________* (____20)
'       EvalSync _ DoEvents ____________. 10 ____ 10 ________.
'
'     *Nothing / False ____ LastEvalOk ____ (____93)*
'       WaitFor  -> Nothing + LastEvalOk=True  ... __________ (__)
'                  Nothing + LastEvalOk=False ... __ (______ / in-callback _)
'       WaitGone -> False   + LastEvalOk=True  ... __________ (__)
'                  False   + LastEvalOk=False ... .
'       *________________* in-callback _ 10 __________.
'
'     ImplicitWaitSec _ QuerySelector / GetElementById .
'     (SeleniumVBA _____). *___ 0 _, __________________*
'     __: ______________________________,
'     D-2 _____Nothing + LastEvalOk=True = _______________.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  K-1 __ (_________________) ---
'
'   K-1 ____ *________* _, _____ 1 ________.
'     EvalSync (9 _) _ ClearElementRegistry (2 _) _ Debug.Print .
'     Wv2Log.LogE / LogW / LogD ______.
'
'   Wv2Log ____ Debug.Print _____, .
'   (Wv2Log.LogEcho = False ____). ___ %APPDATA%\Wv2Browser\logs\ .
'   _____ 1 _________. ___ ?Wv2Log.LogPath ____.
'
'   *_________ Debug.Print _______* .
'   (____ Debug.Print __________, _ K-1 ___).
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  D-2 __ (_______ + Wv2Element _____) ---
'
'   D-2 _____ (JS __________, ___________):
'     GetElementById / QuerySelector ... ___ 1 ____ Wv2Element ___.
'     ClearElementRegistry           ... JS ______________.
'     CurrentDomGen                  ... _____________ (_____).
'     Elem_Read / Elem_CheckStale    ... Wv2Element ___ Friend __.
'     *______* ____ (Value =) / Click / SetAttribute _ D-3.
'                     QuerySelectorAll (____) _ D-3 __ (__7 _a).
'
'   _ ____ Public .
'     GetElementById(elementId) As Wv2Element
'     QuerySelector(selector)   As Wv2Element
'         _____ Wv2Element, ________ *Nothing* ___.
'         *Nothing ____ 2 ______ LastEvalOk ____ (__4)*
'           Nothing + LastEvalOk = True  ... ________ (__)
'           Nothing + LastEvalOk = False ... __. ___ LastEvalError
'                                           (in-callback / timeout / ___________)
'     ClearElementRegistry() As Boolean
'         JS _____________________. ___ Wv2Element .
'         ___ stale ___. _________________________.
'     CurrentDomGen As String
'         ___ JS ______________. ________________.
'
'   _ JS ________ (__1 _c: *_____*)
'     window.__wv2 = { gen: 'g<__>-<__>', els: [__, __, ...] }
'
'     *AddScriptToExecuteOnDocumentCreated _ DOMContentLoaded _________. *
'     ____*___________________*____:
'
'       (function(){ if(typeof window.__wv2==='undefined'||!window.__wv2){
'                      window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]}; }
'                    var R=window.__wv2; ... })()
'
'     __:
'       ____ COM _______ (___________ HandlerKind _ vtable ___)
'       ________ (_________ 1 ______)
'       ______ / SPA _____ / about:blank ________,
'         _________________ (____74 ________)
'       _____ / ____ / _____ Pane .
'     __:
'       ________*_*________________.
'         ________ AddScriptToExecuteOnDocumentCreated __________.
'
'   _ ____ (__2 _b: *JS _____*)
'     ___ JS __ 'g' + Date.now() + '-' + Math.random() ______.
'     *VBA _______________ (_a) _____. *
'       SPA _ pushState __ DOM ____________, ______ JS .
'       ____________________, ________.
'     _________ = ______ = _________, __________,
'     *________ JS _*___.
'
'     m_domGen __JS __________________, ___________:
'       _____ <> m_domGen __*_____ stale ______* (__)
'       _____ =  m_domGen __ JS .
'     ____________________________________.
'
'     *Math.random() ______* Date.now() _________, .
'       ______ 2 _____________. .
'       _______________ (__2 _____).
'
'   _ _______ (__3 _a: *_____*)
'     Wv2Element.Class_Terminate __ JS __________. Terminate .
'     _________________, ____ EvalSync ____ D-1 .
'     in-callback ____ Ctrl_Close ___ Pane ________.
'     ___________________________. .
'     ClearElementRegistry. _________ D-4.
'
'   _ JS _______ (*VBA _____________ + __________*)
'     __________________________, .
'     _____________. JsQuote _ Wv2Json.JsonEscape .
'     Chr$(34) _____, VBA _______________.
'     _: QuerySelector("input[name='q']") _ QuerySelector("a[title=""x""]") ___.
'
'   _ ____ (____79)
'     (a) *JsonPickStr _ Wv2Json.bas _____* (Private -> Public)
'         __6 __JsonUnescape ____ JsonPickStr .
'         ____, Wv2Element __________, Wv2Pane _ Private .
'         ___ 2 _____. _____6 _______________.
'         EvalSync __ 3 ________ Wv2Json.JsonPickStr _______,
'         _____ 1 _________ (____ Wv2Json.JsonUnescapeAt _____).
'     (b) *Wv2Element _ LastOk / LastError ____ (__5 _ 8 ___ + 2)*
'         __________________, ________stale / timeout.
'         ______________. Pane __ LastEvalOk _ JS .
'         ___, stale ___ True _____________.
'     (c) *Elem_Read / Elem_CheckStale _____ Enum ___ Boolean + ______*
'         _________ Public Enum .
'         _________, ____________ 1 _______.
'         D-1 _ LastEvalOk / LastEvalError ________________.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  D-1 __ (ExecuteScript ___________) ---
'
'   *______________*
'     _9.32b ____9.xx___WebView2 .
'     ____, _9.32b ________. D-1 _______________:
'       (__ 9.xx) ... ____________ (_____)
'       D          ... DOM / _______ (MSHTML __)  <- .
'       C          ... __________ (TreeView / ListView __)
'       K          ... _______ (Sentinel _)
'
'   D-1 _____ (JS _________ VBA _________):
'     EvalSync ... _______. JS _ 1 ____, ____ DoEvents .
'                ____ JSON ___ String ___.
'     *______* _______ (__wv2els[]), Wv2Element, DOM API. D-2 __.
'
'   _ ____ Public .
'     EvalSync(jsExpr, [timeoutSec=5]) As String
'         __: JS _____ JSON.stringify ________.
'               1+2       -> 3
'               'abc'     -> _____ abc (*____________*)
'               undefined -> undefined   (JSON.stringify _ undefined _____)
'               null      -> null
'               ({a:1})   -> {a:1} _ JSON .
'         __: ______. ___ LastEvalError ___.
'     LastEvalOk    As Boolean ... ___ EvalSync .
'     LastEvalError As String  ... .
'                                 closed / no-view / in-callback /
'                                 execute-failed(n) / timeout(n s) /
'                                 hresult(&Hxxxx) / ReferenceError: xxx /
'                                 syntax-error-or-null (...)
'     ResetCallbackGuard()     ... in-callback _________ 0 ___ (___)
'     InCallbackDepth As Long  ... ___________ (___)
'     Debug_SetInCallback(n)   ... ________________ (____)
'
'   _ _____ (__________)
'     __1 _____     : ______ EvalSync. View_ / Ctrl_ __COM .
'                            _____________ 1 _ 1 .
'                            ____________ API ______.
'     __2 ______   : callbackId ____________ (m_evalPending).
'                            _____________________.
'                            UserForm1.OnExecuteScriptResult ______ (____).
'     __3 _____     : String (= value _ JSON __). ___ LastEvalOk /
'                            LastEvalError ___. ____ D-2 _____.
'     __4 Safe __      : __ View_ExecuteScriptSafe ___. __ ExecuteScript .
'                            JS ___ S_OK + null ___, ___ null _______.
'     __5 ______   : __ 5 _. ____________*_______*,
'                            __________________ (______).
'     __6 ______   : ____. callbackId ________________.
'                            ______________.
'     __7 in-callback ___:
'                            *_____________* WebView2 .
'                            _______ EvalSync ____, .
'                            _____, ____ JS .
'                            ___ (____ 14). _______ 20 .
'                            _________________. .
'                            m_inCallback > 0 __ JS ___________.
'
'   _ ________ (____ 79)
'     (a) *_____ Core _____* ______ View_On* ____ +1, ___ -1.
'         ______, __________ View_OnNewWindowRequested .
'         Exit Sub _ 5 ____ (__ 1 ____), ___ 2 .
'         ____________. ______ Private ... _Core ______,
'         ___ Public _______ +1 -> Core ____ -> -1 _______.
'         ___ 7 _ (NavigationStarting / NavigationCompleted / DocumentTitleChanged /
'         WebMessageReceived / NewWindowRequested / HistoryChanged / DOMContentLoaded).
'     (b) *____________. _________* .
'         On Error GoTo GuardExit ___________, Err.Raise _*.
'         ____*. ComCallbackHandler __________, .
'         ________, ______________ (.
'         ______). _________________ ResetCallbackGuard ___.
'     (c) *OnExecuteScriptCompleted ____________* .
'         __, ________________. ____ EvalSync _ DoEvents .
'         _______, _________.
'     (d) *JSON ______ Wv2Pane __ Private _____* Wv2Json.JsonGetStr .
'         ____ 64 ___*__________*___, .
'         ______. EvalSync _____ JSON.stringify .
'         _________ + ______________ + uXXXX______.
'         ___ JsonPickStr (_____________) _ Private _____.
'         _____________ Wv2Json ______.
'     (e) *JSON.stringify _______* EvalSync .
'         View_ExecuteScriptSafe("JSON.stringify(" & jsExpr & ")") ___. .
'         Safe ______ JSON _ value _*_____*___, (d) .
'         _____________. undefined ____ value .
'         _______ undefined ______.
'
'   _ ______ (D-2 _______)
'     _jsExpr ____________________ (View_ExecuteScriptSafe .
'       _________). __________________.
'     ____________________. D-2 _ Wv2Element .
'       ________.
'     _*DoEvents _________/___________* (____ 20)
''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  _9.27 __ (__: put_NewWindow ____________) ---
'
'   _9.27 _____ (POST _______________________):
'     View_OnNewWindowRequested _______ 2 _ -> 3 ______.
'       0 = NWM_LEGACY  : ____ (put_Handled + RequestNewTabAsync). Navigate .
'                         *__ GET* ___ POST ____ 405 ___.
'       1 = NWM_POPUP   : 9.26 _____. put_Handled ________, .
'                         ___________. POST _____ (____ 23).
'       2 = NWM_PREWARM : *9.27 ___* Browser ___________ CoreWebView2 .
'                         args.put_NewWindow (vtable index 4) ___. .
'                         __ WebView ______________ POST ____,
'                         ____________.
'
'   *____________* .
'       (1) m_browser.ClaimPrewarmViewPtr ____ ViewPtr .
'       (2) dcf(pArgs, 4, put_NewWindow, pNew)
'       (3) dcf(pArgs, 6, put_Handled, 1&)   _ SetHandledWithNewWindow _ True .
'       (4) m_browser.RequestAdoptPrewarmedAsync ____ Timer .
'     _ 4 ___. Controller ___ Deferral _____, ____ 14 (SafeTimer
'     ______) __ 20 (EBMode __) ______. *______________*
'
'   *_______* _____ / _ Ready / __ Claim __ / put_NewWindow __ .
'     _____, ______ (put_Handled + RequestNewTabAsync) ________.
'     POST _____ 405 __________. ________ *_________.
'
'   *vtable index ___ (WebView2.h _ ICoreWebView2NewWindowRequestedEventArgsVtbl)*
'       3 = get_Uri / 4 = put_NewWindow / 5 = get_NewWindow /
'       6 = put_Handled / 7 = get_Handled / 8 = get_IsUserInitiated / 9 = GetDeferral
'     put_Handled = 6 _ 9.6b ___________, _________.
'
'   *______* get_Uri ______ / _ URL ___ / RequestNewTabAsync (.
'     Timer, ____ 14 ___) / ____________.
'
' --- Wv2Pane.cls  _9.26 __ (NewWindowRequested ______) ---
'
'   _9.26 _____ (POST ____ 405 _________. ________):
'     View_OnNewWindowRequested _______ 1 ______.
'     m_browser.PopupDelegateMode _ True .
'       _args.put_Handled .
'       _m_browser.RequestNewTabAsync .
'     _____, WebView2 .
'     (_9.6a ________: Excel ______ Chrome_WidgetWin_1 ____).
'     ___________________, ________ POST .
'     _____. _______________ POST _____________.
'
'     *_____ View _ View_Navigate __________*
'       9.6b ________ ICoreWebView2::Navigate _____*__ GET*.
'       ____________________, Navigate _______ 405 ___.
'       POST ______________________ (______, .
'       put_NewWindow _) _____________.
'
'     *m_browser _ Nothing ___* __ Pane __ (SetBrowser ___) .
'       ____________, _______ View _ View_Navigate .
'       __. ___ 9.10e ______.
'
'   *______* get_Uri ______ / _ URL ___ / put_Handled _______ /
'     RequestNewTabAsync (__ Timer, ____ 14 ___) / ____________.
'     _______________.
'
' --- Wv2Pane.cls _9.19__ ---
'
'   _9.19 _____ (____________):
'     *m_isSettings ___ (______) ___*
'       __ Pane _____________________. Wv2Browser .
'       OpenSettingsTab _, __________ Pane _ IsSettings = True .
'       ___. [__] _________ Browser _ m_tabs _______,
'       IsSettings = True _ Pane _______ ActivateTab .
'       _____ ([__] _______________).
'
'     *__ Pane __________ (_A)*
'       _____ NavigateToString _ data:text/html;base64,... .
'       ____ URL/_____________. Pane _________,
'       __ m_tabs ______________________ (.
'       _____ m_tabs __________________). Browser .
'       Pane _________________ (___________) .
'       ______.
'
' --- Wv2Pane.cls _9.18__ ---
'
'   _9.18 _____ (AddHostObjectToScript ___ JS __ VBA ________):
'     *View_AddHostObjectToScript(objName, instance) ___* (__1)
'       ICoreWebView2.AddHostObjectToScript (View __ vtable index 49) ___,
'       VBA ___________ JS __ window.chrome.webview.hostObjects.<name>
'       _______. JS ________/.
'       (postMessage _______________ RPC __).
'
'     *____ WebView2.cls (______) ______*
'       _____________ dcf __:
'         Set vObj = instance      (_______ Variant ____)
'         dcf(m_pView, 49, ..., StrPtr(objName), VarPtr(vObj))
'       _1_____________ (LPCWSTR), _2___ VARIANT* (.
'       ____________). ____ HRESULT. Instancing _ 1 - Private .
'       _______ (________. VBA _____ IDispatch _ OS .
'       _________).
'
'     *__ (_9.18 __)*
'       Wv2Browser __________, __ Pane _ Wv2SettingsBridge .
'       attach _, _____ JS __ hostObjects.sync.Settings.SetEngine("bing")
'       __________________. attach _ NavigateToString .
'       __ (__A: ______ Navigate ___, ______ JS .
'       hostObjects ____, ____________. ____ ready
'       _________=__B __________).
'
'     *RemoveHostObjectFromScript (vtable 50) _____* (YAGNI)
'       _________ WebView2 _________, .
'       _9.18 ____. _____________.
'
' --- Wv2Pane.cls _9.12a__ (__: Class9) ---
'
'   _9.12a _____ (______/________: URL _____):
'     *___ 9.11a _ NotifyTitleChanged _____* Pane _____ URL .
'       __________ Browser ___, Browser _ index .
'       RaiseEvent TabUrlChanged __. Pane _____________ (___).
'     - __ Private Sub: NotifyUrlState(pSender)
'         sender ___ get_Source(4) / get_CanGoBack(38) / get_CanGoForward(39) .
'         __, m_browser.NotifyUrlChanged Me, url, canBack, canFwd ___.
'     - View_OnNavigationCompleted: ___ NotifyUrlState ___ (________).
'         *pArgs=0 ___ Exit Sub ________* (________)
'     - View_OnHistoryChanged: __ url/canBack/canFwd ________, .
'         m_browser.NotifyUrlChanged ___ (SPA _ pushState, __/_____).
'     *_______* NavigationCompleted _ HistoryChanged _______,
'       ______ (____ 60) ________ 2 __________.
'
'   _9.11c _____ (VBA->JS _________: _________ + WebMessage ___):
'     *_________ (___3)* _____ Pane .
'       ____. __ index ____________________,
'       _______________. TabBar _ Browser.PaneAt(i).DocumentTitle
'       ___.
'       - _______: m_title As String
'       - __ Public Property Get DocumentTitle() As String
'       - View_OnDocumentTitleChanged ________ m_title = title .
'         (9.11a _ NotifyTitleChanged __________ 1 _____).
'     *__ WebMessage ___ (_P)* JS __ postMessage .
'       __________(TabBar _) ________. m_browser .
'       (NewWindowRequested _) ______ 1 _.
'       - _______: m_webMsgSink As Object (____ OnPaneWebMessage ___)
'       - __ Public Sub: SetWebMsgSink / ClearWebMsgSink
'       - View_OnWebMessageReceived: _______ OnPaneWebMessage(Me, source, msg)
'         __________.
'       - View_AddWebMessageReceived ___ (9.4c) _______, TabBar .
'         ________.
'     *_________* ClearBrowserRef _ m_webMsgSink _ Nothing .
'       (____ 58. _________ Shutdown _____).
'
'   _9.11a _____ (_______________):
'     - View_OnDocumentTitleChanged _ 1 ______. _______,
'       m_browser ____ m_browser.NotifyTitleChanged Me, title ___.
'       ____ m_browser (9.10e _ NewWindowRequested ____) .
'       ____. m_browser ____ (__________) _ Debug.Print .
'       __________, ________. Pane _________/
'       _______ (___ m_browser _____).
'
'   _9.10e _____ (NewWindowRequested _____, _A + _Y + __ Timer):
'     *________ (__)* NewWindowRequested ______ AddTabWithUrl .
'       _____, AddTab _ WaitPaneReady _ ControllerCompleted ____, .
'       _____ return ____ WebView2 .
'       (______) __ (____ 14). ___ m_browser.RequestNewTabAsync .
'       ___, SafeTimer (______) ___________ AddTabWithUrl .
'       _____.
'     - _______: m_browser As Wv2Browser (__ Browser ______)
'     - ______: SetBrowser / ClearBrowserRef (__________)
'     - View_OnNewWindowRequested ____ View Navigate_(9.6b) .
'       _m_browser.AddTabWithUrl ___________. m_browser .
'       __ View Navigate ________. window.open _____ null .
'       (put_NewWindow __ View _________ = _A. __________)
'     *______ (____ 58)* Browser.m_tabs ? Pane.m_browser .
'       __. Wv2Browser.Shutdown __ Pane _ ClearBrowserRef .
'       (UserForm_Terminate __). Class_Terminate __ m_browser .
'       (____ Class_Terminate __________, Shutdown ______)
'
'   _9.10d-2a _____ (______, SPA _ https origin __):
'     - _________:
'         m_pView3 As LongPtr (ICoreWebView2_3 __ IF ______)
'     - Class_Initialize _ m_pView3 = 0 .
'     - Class_Terminate / Ctrl_Close _ Release _____:
'         m_pView3 -> m_pView2 -> m_pView -> m_pCtrl (______, ____ 51)
'     - __ Private Function:
'         EnsureView3() As Boolean
'         (QueryInterface(IID_ICoreWebView2_3) _____, EnsureView2 ___)
'     - __ Public Function:
'         View3_Ensure() As Boolean (______, View2_Ensure ___)
'         View3_SetVirtualHostNameToFolderMapping(hostName, folderPath, accessKind)
'           As Long
'           (ICoreWebView2_3 vtable index 71. .
'            https://<hostName>/ _______, about:blank origin .
'            pushState/replaceState __ (____ 6) _______.
'            ________ View_Navigate _____, ____ 45 ___)
'     - IID_ICoreWebView2_3 = {A0D6DF20-3B92-416D-AA0C-437A9C727857}
'       (EnsureView3 ____________, 9.9a _ a ___)
'
'   _9.10c-1 _____ (____):
'     - __ Public Function:
'         Ctrl_PutIsVisible(visible As Boolean) As Long
'         (ICoreWebView2Controller::put_IsVisible, vtable index 4)
'         Wv2Browser _ ActivateTab _______ (___________) ___.
'         Win32 BOOL (True=1/False=0) ___ dcf __ 1/0 ______ (____ 11).
'         ________ Ctrl_PutBounds ______.
'
'   _9.9c _____ (_______):
'     - __ Public __: View_ExecuteScriptSafe(userCode) As Long
'         ________ try-catch _______ JSON ___ ExecuteScript ___.
'         resultJson ___ {ok:true,value:...} or {ok:false,errorName:...,...}
'         ______ (_ F: a + X + *).
'     - OnExecuteScriptCompleted __ UserForm1.OnExecuteScriptResult .
'         (_ _ ____, _ E1-a-i = Wv2Pane _ UserForm1 ________.
'          VBA ______ UserForm _____, ______).
'         Wv2Pane __ Debug.Print ___ (_ E2-X, _____________).
'     - View_OnHistoryChanged _____:
'         (_) about:blank ___NavigateToString __________.
'         (_) pArgs=0 ___________________.
'     - _______________ (____).
'
'   _9.9b _____ (__, SPA __):
'     - ____________ HK_NewWindowRequested(7) -> HK_DOMContentLoaded(10)
'       .
'     - ______ RemoveInterfaceFor(kind) As LongPtr
'     - HistoryChanged _ (__ ICoreWebView2, vtable 13/14):
'         View_AddHistoryChanged / View_RemoveHistoryChanged / View_OnHistoryChanged
'     - DOMContentLoaded _ (ICoreWebView2_2, vtable 64/65):
'         View2_AddDOMContentLoaded / View2_RemoveDOMContentLoaded /
'         View2_OnDOMContentLoaded
'
'   * _9.9b _____ 45 ____ *
'     _______ 2 ___ = COM _______ WebView2 ______ 2 _.
'     9.8c _ h.Slot ___________________,
'     View_AddHistoryChanged / View2_AddDOMContentLoaded .
'     View_AddNavigationStarting _____________,
'     pVTableObj = h.Slot + VTABLE_OBJ_OFFSET _____.
'     View_PostWebMessageAsString ________________.
'
'   _9.9a _____ (__):
'     - _________:
'         m_pView2 As LongPtr (ICoreWebView2_2 __ IF ______)
'     - Class_Initialize _ m_pView2 = 0 .
'     - Class_Terminate _ Release _____:
'         (Settings -> View2 -> View -> Ctrl, _______, ____ 43)
'     - Ctrl_Close ____ Release _____ (m_pView2 _ View ____ Release)
'     - __ Private Function:
'         EnsureView2() As Boolean
'         (m_pView2 ______ QueryInterface(IID_ICoreWebView2_2) ___,
'          EnsureSettings _______)
'     - __ Public Function:
'         View2_Ensure() As Boolean
'         (______, Immediate ________________)
'
'   _9.9a _____ (__________):
'     __ 1: m_pView2 _____ + EnsureView2 ____ (Settings _____)
'     __ a: IID_ICoreWebView2_2 _ EnsureView2 .
'             (Wv2Thunks _ IID ________, 9.9a _ Wv2Thunks _____)
'     __ X: EnsureView2 / View2_Ensure __ As Boolean (EnsureSettings ___)
'     __ i: ___ Immediate _ View2_Ensure __ + Test_AllEvents_Real .
'             ___________ (___________)
'     __ R1: Release ___ Settings -> View2 -> View -> Ctrl (_______)
'
'   * _9.9a ___ *
'     ICoreWebView2_2 ________________, _9.9b .
'     SPA __ (DOMContentLoaded __, vtable 64/65 _______) .
'     _______________. HistoryChanged ___ IF (vtable 13/14)
'     _____ _2 _____ 9.9b _______, DOMContentLoaded .
'     _2 IF ___. _ 9.6a _________ (View2_xxx) .
'     _____________.
'
'   _9.9a _________ (_________):
'     _____ 31: ____________ (View2_xxx)
'     _____ 42: __ IF _____________ (EnsureSettings ___)
'     _____ 43: Release .
'     _____ 45: COM ______________ VTABLE_OBJ_OFFSET .
'                    (9.9a ___WebView2 __ COM .
'                     ___ VTABLE_OBJ_OFFSET _______, 9.9b .
'                     DOMContentLoaded _________________)
'     _____ 48: Close _____________ + .
'
'   _9.8c _____ (__):
'     - _________:
'         m_executeScriptCounter As Long (View_ExecuteScript _ callback ID __)
'         m_closed As Boolean (Ctrl_Close _____)
'     - Class_Initialize ___ 2 .
'     - __ Public __ (View_ _):
'         View_ExecuteScript (jsCode As String) As Long  ' callback ID .
'     - __ Public Sub (______________):
'         OnExecuteScriptCompleted (callbackId, errorCode, resultJson)
'     - Ctrl_Close ___ (_ E 1+3 ______):
'         Controller::Close ___ + m_pSettings/m_pView/m_pCtrl _ Release +
'         m_persistentTokens / m_persistentHandlers ____ +
'         m_closed = True + m_state = Cs_Failed
'         _____ Close __ Class_Terminate .
'
'   * _9.8c ______ + ____ *
'     View_ExecuteScript _____ h.Slot ___ dcf .
'     (Excel _______). ________ (View_AddNavigationStarting _)
'     ____ h.Slot + VTABLE_OBJ_OFFSET ________.
'     -> _________ WebView2 _ COM _____________,
'        __ VTABLE_OBJ_OFFSET (= 176 byte) .
'        _____________ (____ 45 _____).
'
'   _9.8c _____ (__________):
'     __ a: ____ OK, .
'     __ _: ____________, _____ OnExecuteScriptCompleted .
'             ______ (__ UserForm ______)
'     __ X: __ callback ID .
'     __ E 1+3: ______ + m_closed .
'
'   _9.8b ___ (__):
'     - m_pSettings + EnsureSettings + Settings_PutXxx _ (3 _)
'     - View_Reload / View_Stop / View_GetCanGoBack / View_GetCanGoForward /
'       View_GetBrowserProcessId
'
'   _9.8a ___ (__, __):
'     - ____ Class9 -> Wv2Pane .
'     - Class7 -> ComCallbackHandler / Class8 -> Wv2Environment .
'     - _______.md .

'   __:
'     ICoreWebView2Controller _ ICoreWebView2 ______ (____ = _ a).
'     SDK _______ 1 _ 1 _______, VBA .
'     _1 __ WebView2 __ = 1 _________________.
'     UserForm _ Frame _______________1 .
'     ____________ (= _Pane________).
'
'   __________:
'     HK_NavigationStarting      (View_AddNavigationStarting   / View_OnNavigationStarting)
'     HK_NavigationCompleted     (View_AddNavigationCompleted  / View_OnNavigationCompleted)
'     HK_WebMessageReceived      (View_AddWebMessageReceived   / View_OnWebMessageReceived)
'     HK_DocumentTitleChanged    (View_AddDocumentTitleChanged / View_OnDocumentTitleChanged)
'     HK_NewWindowRequested      (View_AddNewWindowRequested   / View_OnNewWindowRequested)
'     HK_HistoryChanged          (View_AddHistoryChanged       / View_OnHistoryChanged)        ' _9.9b
'     HK_DOMContentLoaded        (View2_AddDOMContentLoaded    / View2_OnDOMContentLoaded)     ' _9.9b
'
'   ____ (____ 31):
'     _ICoreWebView2Controller _____: Ctrl_xxx (__) / Ctrl2_xxx (__, __)
'     _ICoreWebView2 _____:           View_xxx (__) / View2_xxx (__, _9.9a ____)
'     _ICoreWebView2 __________: View_OnXxx (__) / View2_OnXxx (__, 9.9b ______)
'     _ICoreWebView2Environment _____: Env_xxx _ (________)
'     _1 ___________ (OnEnvironmentCompleted / OnControllerCompleted):
'       __________ (.
'       _____________)
'
'   ________ (_9.4c ___, __________):
'     [VBA -> JS]
'       VBA: m_pane.View_PostWebMessageAsString("hello")
'         -> dcf(m_pView, 33, ..., StrPtr("hello"))
'         -> WebView2 ______ JS .
'         -> JS: window.chrome.webview.addEventListener('message', ...) .
'     [JS -> VBA]
'       JS: window.chrome.webview.postMessage("hello")
'         -> WebView2 ______ HK_WebMessageReceived .
'         -> ComCallbackHandler.Invoke[HK_WebMessageReceived]
'         -> Wv2Pane.View_OnWebMessageReceived(pSender, pArgs)
'         -> args.get_Source / args.TryGetWebMessageAsString .
''''''''''''''''''''''''''''''''''
Option Explicit

' Supports both 64-bit (x64) and 32-bit (x86) VBA7 Office hosts
Public Event ExecuteScriptResult(ByVal callbackId As Long, ByVal errorCode As Long, ByVal resultJson As String)
Public Event WebMessageReceived(ByVal messageText As String, ByVal messageJson As String)

' --- N-1: _________ ---
'   *Timer _____* Timer _ Single ____ 0 _________,
'   ____ (5 ___) __ Single ________ ms ________.
'   _____________. GetTickCount ______ ms _____.
'   ____ Long ____ 49.7 ______. __ Double _________,
'   ______ 2^32 ______ (NetElapsedMs).
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long


Public Enum ControllerState
    Cs_Created = 0
    Cs_Waiting = 1
    Cs_Ready = 2
    Cs_Failed = 3
End Enum


' --- ____ ---
' m_handler:     HK_ControllerCompleted _________ (1 ____)
'                (OnControllerCompleted __ Nothing ___)
' m_pCtrl:       WebView2 _______ ICoreWebView2Controller
'                (Class_Terminate _ Release _______)
' m_pView:       Ctrl_GetCoreWebView2 _____ ICoreWebView2
'                (Class_Terminate _ Release _______)
Private m_handler As ComCallbackHandler
Private m_pCtrl   As LongPtr
Private m_pView   As LongPtr
'   _9.8b _____: ICoreWebView2Settings .
'     Wv2Pane ____ Settings .
'     View vtable 3 (get_Settings) ____, _______.
'     m_pView __________ Class_Terminate _ Release __.
Private m_pSettings As LongPtr
'   _9.9a _____: ICoreWebView2_2 __ IF .
'     Wv2Pane ____ View2_xxx .
'     QueryInterface(IID_ICoreWebView2_2) ____, _______.
'     EnsureView2 _______ (= EnsureSettings _____, ____ 42).
'
'     m_pView _ m_pView2 ___ COM ________ IF _____.
'     QueryInterface ___ AddRef _____, m_pView2 _ m_pView .
'     _____________. Class_Terminate / Ctrl_Close .
'     m_pView ____ m_pView2 _ Release __ (_____, ____ 43).
Private m_pView2 As LongPtr
'   _9.10d-2a _____: ICoreWebView2_3 __ IF .
'     SetVirtualHostNameToFolderMapping (vtable 71) ________.
'     EnsureView3 _______ (EnsureView2 ___, ____ 42/50).
'     ICoreWebView2_3 : ICoreWebView2_2 : ICoreWebView2 ________,
'     Release ____ m_pView3 -> m_pView2 -> m_pView (______, ____ 51).
Private m_pView3 As LongPtr
'   _9.10e _____: ____ Wv2Browser .
'     View_OnNewWindowRequested __________ Wv2Browser.AddTabWithUrl .
'     ________. Wv2Browser.AddTab _ SetBrowser _________.
'     *____ (Browser.m_tabs ? Pane.m_browser) _____, .
'      Wv2Browser.Shutdown -> ClearBrowserRef _________ (____ 58).
Private m_browser As Wv2Browser
Private m_state   As ControllerState
Private m_lastError As Long

' --- _____________ (_9.4a _____, _9.6a/9.9b _____) ---
' m_persistentHandlers: HandlerKind ____ index ___ ComCallbackHandler .
' m_persistentTokens:   _ index _ EventRegistrationToken (LongLong, 8 bytes)
'
' _9.6a ______ HK_DocumentTitleChanged -> HK_NewWindowRequested ___.
' _9.9b ______ HK_NewWindowRequested(7) -> HK_DOMContentLoaded(10) ___.
' N-1 ______ HK_DOMContentLoaded(10) -> HK_WebResourceRequested(11) ___.
' N-3 ______ HK_WebResourceRequested(11) -> HK_WebResourceResponseReceived(12) ___.
'   _____ HK_HistoryChanged(9) / HK_DOMContentLoaded(10) .
'   _____.
'
' __: __ 3?10 ___ HK_ExecuteScriptCompleted(8) ______,
'   ExecuteScript _ 1 .
'   (= index 8 ___ Nothing / 0 ______). CleanupAllPersistentHandlers .
'   _If Not (m_persistentHandlers(k) Is Nothing).
'   _______.
'
' Class_Terminate ______ Nothing _________ remove_ .
' (CleanupAllPersistentHandlers _ LBound/UBound ________,
'  ____HandlerKind __________________).
Private m_persistentHandlers(HK_NavigationStarting To HK_WebResourceResponseReceived) As ComCallbackHandler
Private m_persistentTokens(HK_NavigationStarting To HK_WebResourceResponseReceived) As Currency

' --- ____________ (_9.4c _____, ____) ---
'
'   View_OnNavigationCompleted _ True _______.
'   __ (UserForm _) _________________ PostMessage .
'   _________.
'
'   __ (____):
'     ________________________. SPA .
'     history.pushState _ fetch _____________________.
'     SPA __ (DOMContentLoaded, HistoryChanged _) .
'     _________.
Private m_pageLoaded As Boolean


' _9.8c ___:
'   m_executeScriptCounter: View_ExecuteScript .
'     ______ (_ X __). _ ComCallbackHandler.m_callbackId ______.
'     ___ 0, ________ 1 ___.
'   m_closed: Ctrl_Close ______, ___ Class_Terminate ________.
'     True ___ View_xxx / Ctrl_xxx / Settings_xxx ______ return.
'     _____ Close _____________________.
'     (_ E 1+3 ______: ______ + m_closed ___)
Private m_executeScriptCounter As Long
Private m_closed As Boolean

' --- _9.11c _____ ---
'   m_title: .
'     (___3). View_OnDocumentTitleChanged ______. _____ Pane
'     ______________, __ index .
'     ________ (____________). TabBar .
'     Browser.PaneAt(i).DocumentTitle ___.
'   m_webMsgSink: JS __ postMessage ____________ (_P).
'     TabBar ____________. ____ OnPaneWebMessage .
'     __ (_________ As Object). m_browser (NewWindowRequested .
'     ____) _____. ___ ClearBrowserRef ___ (____ 58).
Private m_title As String
Private m_webMsgSink As Object
Private m_isSettings As Boolean   ' _9.19 (__ Pane _________)

' --- _9.27: NewWindowRequested ____ (Wv2Browser.NewWindowMode ___) ---
'   __________ Public Const _______, Browser .
'   Private Const ___. _________________.
Private Const NWM_LEGACY  As Long = 0
Private Const NWM_POPUP   As Long = 1
Private Const NWM_PREWARM As Long = 2


' --- D-1: EvalSync (__________) ____________ ---
'   m_evalPending : callbackId ____________ (__2).
'                   ___ cb & callbackId, __ Array(__, errorCode, resultJson).
'                   *Collection ______________, ___ Remove -> Add*
'   m_lastEvalOk / m_lastEvalError : ___ EvalSync ______ (__3).
'   m_inCallback  : WebView2 _______________ (__7 ____).
'                   Public _____ +1 / -1 __. 0 _____ EvalSync _____.
'   m_evalDepth   : EvalSync ______. ______________ (__6).
Private m_evalPending   As Collection
Private m_lastEvalOk    As Boolean
Private m_lastEvalError As String
Private m_inCallback    As Long
Private m_evalDepth     As Long

' .
Private Const EVAL_WAITING   As Long = 0
Private Const EVAL_DONE      As Long = 1
Private Const EVAL_DISCARDED As Long = 2

' __________ (_). READY_TIMEOUT (Controller ____ 10 _) ______.
' JS _ 1 _________________, 5 ____________.
Private Const EVAL_TIMEOUT_DEFAULT As Single = 5

' --- D-3: DOM _______ (__2 _C / __4 _A) ---
'   m_implicitWaitSec : QuerySelector / GetElementById ________ (_).
'                       *__ 0 = __* ImplicitWaitSec __________.
'   POLL_INTERVAL_SEC : _______ (__4 _A ___ 100ms).
'                       *EvalSync 1 _____ 38ms ___, .
'                       _ 140ms ___* (2026-08-22 _______).
'   WAIT_TIMEOUT_DEFAULT : WaitFor / WaitGone ______ (_).
Private m_implicitWaitSec As Single
Private Const POLL_INTERVAL_SEC    As Single = 0.1
Private Const WAIT_TIMEOUT_DEFAULT As Single = 10


' --- D-4a: ____ SPA ________ ---
'   ____________*_____*. ______________,
'   ___________________ (____________).
Private Const SPA_PROBE_VER As Long = 3


' --- D-6: QuerySelectorAll ---
'   ALL_MAX_DEFAULT    : 1 _________. *_______* (LastAllTruncated)
'   ELEMENT_WARN_COUNT : _________________ (______)
Private Const ALL_MAX_DEFAULT    As Long = 200
Private Const ELEMENT_WARN_COUNT As Long = 5000
Private m_allTruncated As Boolean


' --- D-4b: _______ ---
'   m_ignoreNet / m_ignoreSel : *VBA ____* _____ (_____________).
'   m_autoWait   : ________ WaitSettled ___. *__ False*
'   m_lastSettle : ________ (___).
Private m_ignoreNet  As Collection
Private m_ignoreSel  As String
Private m_autoWait   As Boolean
Private m_lastSettle As String

' ______ (ms). _______ (WebDriver BiDi for SeleniumVBA) ____.
Private Const SETTLE_STABLE_MS_DEFAULT As Long = 500

' ________ (ms). *10 _________________*
Private Const STALE_INFLIGHT_MS_DEFAULT As Long = 10000


' --- D-4c: ______ (arm) ---
'   *VBA __ arm _______* ____________, .
'   _______arm ________________ (.
'   ______). _____, ______ signal-lost ________.
Private m_sigDom As Boolean
Private m_sigNet As Boolean
Private m_sigUrl As Boolean

' --- D-4f: ______________ ---
'   *__________________________*
'   Google Maps ____ /search?tbm=map _________________,
'   __________________ (D-4d ___).
'   ____________________________________.
'   0 ______ (______).
Private m_staleMs As Long


' --- N-2: ________ (_______) ---
'   *_________* (__5 _S). ___________,
'   ____________ 1 ________.
'
'   m_detOn        : ______. *__ False* (__8 __)
'   m_detCap       : ________ (__ 50)
'   m_detHead      : ______ (0 __)
'   m_detCount     : .
'   m_detDropped   : *________* .
'   m_detBodyMax   : ________ (___. __ 65536)
'   m_detRedact    : ___________. *__ True* (__3 _M)
'   m_detRedactList: _________ |___| .
'   m_detUriFilter : _____ URL _____ (_____)
Private m_detOn         As Boolean
Private m_detCap        As Long
Private m_detHead       As Long
Private m_detCount      As Long
Private m_detDropped    As Long
Private m_detBodyMax    As Long
Private m_detRedact     As Boolean
Private m_detRedactList As String
Private m_detUriFilter  As String
Private m_detSeq()     As Long
Private m_detMethod()  As String
Private m_detUri()     As String
Private m_detHeaders() As String
Private m_detCType()   As String
Private m_detBytes()   As Long
Private m_detTrunc()   As Boolean
Private m_detIsText()  As Boolean
Private m_detBody()    As String
Private m_detHex()     As String
Private m_detRespHeaders() As String   ' N-3: .
Private m_detStatus()      As Long     ' N-3: _____ (0 = __)
Private m_detReason()      As String   ' N-3: .
Private m_detRespBody()    As String   ' N-4: _______ (______)
Private m_detRespBytes()   As Long     ' N-4: .
Private m_detRespTrunc()   As Boolean  ' N-4: .
Private m_detRespIsText()  As Boolean  ' N-4: .
Private m_detRespEnc()     As String   ' N-4: content-encoding (*______*)
Private m_detRespHex()     As String   ' N-4: *__ 32 ____ hex (____)*
Private m_detRespGot()     As Boolean  ' N-4: *_______* (0 _________)

' _____________. *_________* 1 __________.
Private Const NET_DET_CAP_DEFAULT As Long = 50
Private Const NET_DET_CAP_MAX     As Long = 2000

' ______ (__ 64KB) _, _________ (8MB).
'   *_____________* _________.
Private Const NET_BODY_MAX_DEFAULT As Long = 65536
Private Const NET_BODY_MAX_LIMIT   As Long = 8388608
Private Const NET_BODY_CHUNK       As Long = 16384

' ___________. 1 ___ COM _____ 3 _ x ______.
Private Const NET_HDR_MAX As Long = 40

' _____________________.
Private Const NET_HEX_HEAD As Long = 32

' IStream::Seek _ dwOrigin (Windows SDK objidl.h _ STREAM_SEEK_SET).
Private Const STREAM_SEEK_SET As Long = 0


' --- N-1: ______________ (WebResourceRequested) ---
'   m_netOn      : ____ (add_WebResourceRequested ___)
'   m_netCap     : ______ (_). NetLogCapacity ______.
'   m_netHead    : ______ (0 __)
'   m_netCount   : ____________ (0 _ m_netCap)
'   m_netTotal   : _______*____* (________)
'   m_netDropped : *________* .
'   m_netT0      : ______ GetTickCount _. _____ ms ___.
'   m_netFilters : ________ (uri & vbTab & ctx). Stop _________.
'   m_netStatus  : _________ (N-3). *0 = __________*
'   m_netReason  : ___ (N-3)
'   m_netRespOn  : _____________ (N-3)
'   m_netUnmatched : *______________* (N-3. _____)
Private m_netOn      As Boolean
Private m_netCap     As Long
Private m_netHead    As Long
Private m_netCount   As Long
Private m_netTotal   As Long
Private m_netDropped As Long
Private m_netT0      As Long
Private m_netRespOn    As Boolean
Private m_netUnmatched As Long

' --- N-4: _______ ---
'   m_respBodyOn     : ______. *__ False* (__1 _A)
'   m_respBodyMax    : _____ (___. __ 65536)
'   m_respPending    : *__________ ResponseView* (__5 __)
'                      ___ "r" & ____. __ AddRef _______.
'   m_respBodyOrphan : ____________ (_______)
Private m_respBodyOn     As Boolean
Private m_respBodyMax    As Long
Private m_respPending    As Collection
Private m_respBodyOrphan As Long

' _____________ (64KB). ________ NET_BODY_MAX_LIMIT ___.
Private Const NET_RESP_BODY_MAX_DEFAULT As Long = 65536

' *204 No Content __ GetContent ______ HRESULT ___* (N-4 ___).
'   0x800700E8 = HRESULT_FROM_WIN32(ERROR_NO_DATA). *.
'   __________* ___, .
'   (____111: _____________).
Private Const E_NO_DATA As Long = &H800700E8
Private m_netFilters As Collection
Private m_netStatus()  As Long
Private m_netReason()  As String
Private m_netSeq()    As Long
Private m_netMs()     As Long
Private m_netCtx()    As Long
Private m_netMethod() As String
Private m_netUri()    As String

' ___________. __ 500 ___1 _________________.
Private Const NET_CAP_DEFAULT As Long = 500
Private Const NET_CAP_MAX     As Long = 50000

' _________ (__4). ___CSS________________.
'   1 = DOCUMENT (____ POST / _____) / 7 = XML_HTTP_REQUEST / 8 = FETCH
Private Const NET_CTX_DEFAULT As String = "DOCUMENT,XHR,FETCH"


' --- D-2: _______________ ---
'   m_domGen : JS __________ window.__wv2.gen. ______.
'              ____________________.
'              *_________* (_____ JS _. ________)
Private m_domGen As String


' ============================================================
' Class_Initialize / Class_Terminate
' ============================================================
Private Sub Class_Initialize()
    m_staleMs = STALE_INFLIGHT_MS_DEFAULT
    m_state = Cs_Created
    m_pCtrl = 0
    m_pView = 0
    m_pView2 = 0    ' _9.9a ___ (__ IF ______, _____)
    m_pView3 = 0    ' _9.10d-2a ___ (ICoreWebView2_3 __ IF, _____)
    Set m_browser = Nothing   ' _9.10e (____, AddTab __ SetBrowser ____)
    m_lastError = 0
    m_pageLoaded = False
    m_executeScriptCounter = 0
    m_closed = False
    m_title = ""                  ' _9.11c (_________)
    Set m_webMsgSink = Nothing    ' _9.11c (WebMessage ___, ___)
    m_isSettings = False          ' _9.19 (_______)

    ' --- D-1: EvalSync ______ ---
    Set m_evalPending = New Collection
    m_lastEvalOk = False
    m_lastEvalError = ""
    m_inCallback = 0
    m_evalDepth = 0

    ' --- D-2: _______________ ---
    m_domGen = ""

    ' --- N-1: ________________ ---
    m_netOn = False
    m_netCap = NET_CAP_DEFAULT
    Set m_netFilters = New Collection
    NetRingAlloc

    ' --- N-2: __ (_______) ---
    m_detOn = False                       ' *__ OFF* (__8 __)
    m_detCap = NET_DET_CAP_DEFAULT
    m_detBodyMax = NET_BODY_MAX_DEFAULT
    m_detRedact = True                    ' *______* (__3 _M)
    m_detRedactList = NetRedactDefaults()
    m_detUriFilter = ""
    NetDetailRingAlloc

    ' --- N-4: _______ ---
    m_respBodyOn = False                  ' *__ OFF* (__1 _A)
    m_respBodyMax = NET_RESP_BODY_MAX_DEFAULT
    Set m_respPending = New Collection
    m_respBodyOrphan = 0
End Sub

Private Sub Class_Terminate()
    ' COM _____: __________ ComAddRef _ VBA .
    ' ________, ___ Release __.
    '
    ' __ (_9.8b _ m_pSettings ___, _9.9a _ m_pView2 ___):
    '   1. _____________ (m_pView _____________)
    '   2. m_pSettings _ Release (Settings _ View ________)
    '   3. m_pView2 _ Release  <- _9.9a .
    '                            (View2 _ View __ QI ________ IF,
    '                             _______________ View ___)
    '   4. m_pView _ Release
    '   5. m_pCtrl _ Release
    '
    ' _______:
    '   - remove_*Event _ m_pView (= ICoreWebView2) .
    '     m_pView .
    '   - Settings _ View ____________. View ____ Release .
    '   - View2 _ View _ QueryInterface ________ IF. __ COM
    '     ________ IF _________, QI _____ AddRef .
    '     ____, View2 _ View __________________.
    '     ____ 43 (_____) ___ View2 ___ Release .
    '   - View _ Ctrl ____________. Controller .
    '     Release ___ WebView2 __________________,
    '     _______ (Ctrl __, View __, Settings __) .
    '     VBA ___________, _________.

    ' --- K-4a: *____________* ---
    '   _______ (Cleanup _ 1 __) ___ remove_*Event ______.
    Wv2Log.LogD "Wv2Pane.Terminate: *__* (pCtrl=" & m_pCtrl & _
                " pView=" & m_pView & ")"

    ' --- 1. _______________ (_9.4a _____) ---
    CleanupAllPersistentHandlers
    Wv2Log.LogD "Wv2Pane.Terminate: ____________"

    ' --- 2. m_pSettings _ Release (_9.8b _____) ---
    If m_pSettings <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pSettings = " & m_pSettings
        ComRelease m_pSettings
        m_pSettings = 0
    End If

    ' --- 3-pre. m_pView3 _ Release (_9.10d-2a _____) ---
    '   View3 _ View _ QI _________ IF (ICoreWebView2_3).
    '   ____________, m_pView2 ____ Release (____ 51).
    If m_pView3 <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pView3 = " & m_pView3
        ComRelease m_pView3
        m_pView3 = 0
    End If

    ' --- 3. m_pView2 _ Release (_9.9a _____) ---
    '   View2 _ View _ QueryInterface ________ IF. View ____ Release.
    If m_pView2 <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pView2 = " & m_pView2
        ComRelease m_pView2
        m_pView2 = 0
    End If

    ' --- 4. m_pView _ Release ---
    If m_pView <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pView = " & m_pView
        ComRelease m_pView
        m_pView = 0
    End If

    ' --- 5. m_pCtrl _ Release ---
    If m_pCtrl <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pCtrl = " & m_pCtrl
        ComRelease m_pCtrl
        m_pCtrl = 0
    End If

    ' m_handler ___ OnControllerCompleted _ Nothing __________,
    ' __ Init ______________________.
    If Not (m_handler Is Nothing) Then
        Wv2Log.LogD "Wv2Pane.Terminate: m_handler __________ (______)"
        Set m_handler = Nothing
    End If

    Wv2Log.LogD "Wv2Pane.Terminate: *__*"   ' K-4a
End Sub


' ============================================================
' CleanupAllPersistentHandlers (_9.4a _____, ______)
'
'   m_persistentHandlers ______________________,
'   WebView2 _________. Class_Terminate ______.
'
'   _ HandlerKind ____, ____ remove_ vtable index ______.
'   ________ WebView2 _ ComCallbackHandler _ Release -> refcount=0 ->
'   _______ ComCallbackHandler.ClearOwner -> m_persistentHandlers(k) .
'   ClearOwner _______ Nothing _________.
'
'   ____m_persistentHandlers(k) ___ Nothing ____,
'   ClearOwner ___ m_owner ___ Nothing .
'   _____ (ClearOwner _ ComCallbackHandler.m_owner _ Nothing ______,
'   Wv2Pane __ m_persistentHandlers ______).
'
'   _____:
'     1. View_RemoveEvent (remove_*Event) _ WebView2 .
'     2. WebView2 _ ComCallbackHandler _ Release -> __________ ->
'        ComCallbackHandler.ClearOwner _ ComCallbackHandler.m_owner = Nothing
'     3. ___ Wv2Pane __ m_persistentHandlers(k) = Nothing .
'
'   _9.6a ______ HK_NewWindowRequested _______,
'   ______ LBound/UBound .
'   (= ____ 15 _HandlerKind ___________________).
' ============================================================
Private Sub CleanupAllPersistentHandlers()
    If m_pView = 0 Then
        ' View ____________, remove_ ________.
        ' _____________________, ______ (_____).
        Dim k1 As Long
        For k1 = LBound(m_persistentHandlers) To UBound(m_persistentHandlers)
            Set m_persistentHandlers(k1) = Nothing
            m_persistentTokens(k1) = 0@
        Next k1
        Exit Sub
    End If

    Dim k As Long
    For k = LBound(m_persistentHandlers) To UBound(m_persistentHandlers)
        If Not (m_persistentHandlers(k) Is Nothing) Then
            Dim removeIndex As Long
            removeIndex = RemoveVtableIndexFor(CLng(k))
            If removeIndex > 0 Then
                ' _9.9b: remove ______________ kind _____.
                '   DOMContentLoaded(_2 IF) _ m_pView2, _____ m_pView.
                '   ____ 45 ___ (__ IF ________ IF ________).
                Dim pIfRemove As LongPtr
                pIfRemove = RemoveInterfaceFor(CLng(k))
                If pIfRemove <> 0 Then
                    Wv2Log.LogD "Wv2Pane.Cleanup: removing persistent handler kind=" & k & _
                                " vtable=" & removeIndex & " token=" & m_persistentTokens(k) & _
                                " iface=" & pIfRemove
                    dcf pIfRemove, removeIndex, "View.remove_Event", m_persistentTokens(k)
                Else
                    ' ______________ (_: m_pView2) ____ / ____.
                    ' remove ________________ (_____).
                    Wv2Log.LogD "Wv2Pane.Cleanup: kind=" & k & " _ remove _ IF _ 0, " & _
                                "remove ____ (______)"
                End If
            End If
            Set m_persistentHandlers(k) = Nothing
            m_persistentTokens(k) = 0@
        End If
    Next k
End Sub


' ============================================================
' RemoveInterfaceFor (_9.9b _____, ______)
'
'   HandlerKind _____ remove_*Event .
'   _______. RemoveVtableIndexFor __vtable index________,
'   _______________ (m_pView / m_pView2 / ...) ________.
'
'   ____ 45 (__ IF ________ IF ________):
'     DOMContentLoaded _ ICoreWebView2_2 _____ (vtable 65) .
'     m_pView2 _ remove __. _______ ICoreWebView2 .
'     m_pView _ remove __.
'
'     9.9a _____m_pView2 == m_pView (___)_________,
'     _______ m_pView _ vtable 65 ________ IF .
'     _______________ (= 9.8c _______________).
'     _________________, ___ SDK _______.
'
'   ___:
'     _________________.
'     0 ________remove _________(_: DOMContentLoaded .
'     m_pView2 ____ / ______) __, ______ remove _______.
'
'   _____:
'     ICoreWebView2_3 ____________, ___ Case _ 1 .
'     ______ (RemoveVtableIndexFor ________).
' ============================================================
Private Function RemoveInterfaceFor(ByVal kind As Long) As LongPtr
    Select Case kind
        Case HK_DOMContentLoaded, HK_WebResourceResponseReceived
            ' _2 IF _____. m_pView2 ___ (_____ 0 ___).
            '   *N-3 _ WebResourceResponseReceived _ _2 IF* _______.
            RemoveInterfaceFor = m_pView2
        Case Else
            ' __ ICoreWebView2 _____. m_pView ___.
            RemoveInterfaceFor = m_pView
    End Select
End Function


' ============================================================
' RemoveVtableIndexFor (_9.4a _____, ______)
'
'   HandlerKind _____ remove_*Event _ vtable index ___.
'   add_*Event _________, SDK _____________.
'
'   ICoreWebView2 vtable layout (__):
'     0  QueryInterface
'     1  AddRef
'     2  Release
'     3  get_Settings
'     ...
'     5  Navigate
'     ...
'     7  add_NavigationStarting       8  remove_NavigationStarting
'     ...
'    13  add_HistoryChanged          14  remove_HistoryChanged       <- _9.9b .
'    15  add_NavigationCompleted     16  remove_NavigationCompleted
'     ...
'    34  add_WebMessageReceived      35  remove_WebMessageReceived
'     ...
'    44  add_NewWindowRequested      45  remove_NewWindowRequested  <- _9.6a .
'    46  add_DocumentTitleChanged    47  remove_DocumentTitleChanged
'    55  add_WebResourceRequested    56  remove_WebResourceRequested  <- N-1 .
'    (_2 IF) 61  add_WebResourceResponseReceived  62  remove_  <- N-3 .
'
'   ICoreWebView2_2 vtable layout (___ 61 ______ IF):
'    64  add_DOMContentLoaded        65  remove_DOMContentLoaded     <- _9.9b .
'       _ ___ add/remove _ vtable index _ m_pView2 (ICoreWebView2_2) .
'         _____. RemoveInterfaceFor _ DOMContentLoaded __ m_pView2 ___.
'
'   _9.6a _ HK_NewWindowRequested = 45 ___.
'   _9.9b _ HK_HistoryChanged = 14, HK_DOMContentLoaded = 65 ___.
' ============================================================
Private Function RemoveVtableIndexFor(ByVal kind As Long) As Long
    Select Case kind
        Case HK_NavigationStarting:   RemoveVtableIndexFor = 8
        Case HK_NavigationCompleted:  RemoveVtableIndexFor = 16
        Case HK_WebMessageReceived:   RemoveVtableIndexFor = 35
        Case HK_DocumentTitleChanged: RemoveVtableIndexFor = 47
        Case HK_NewWindowRequested:   RemoveVtableIndexFor = 45   ' _9.6a .
        Case HK_HistoryChanged:       RemoveVtableIndexFor = 14   ' _9.9b ___ (__ IF)
        Case HK_DOMContentLoaded:     RemoveVtableIndexFor = 65   ' _9.9b ___ (_2 IF)
        Case HK_WebResourceRequested: RemoveVtableIndexFor = 56   ' N-1 ___ (__ IF)
        Case HK_WebResourceResponseReceived: RemoveVtableIndexFor = 62  ' N-3 ___ (_2 IF)
        Case Else:                    RemoveVtableIndexFor = 0   ' __ -> .
    End Select
End Function


' ============================================================
' Init
'
'   ICoreWebView2Controller ________.
'   ___________________________ (___).
'   _____ OnControllerCompleted ____.
'
'   ______ ctrl.IsReady _ True _____ DoEvents ___.
'   ___ WebView2Loader _________________, .
'   Init __________ Ready / Failed ______.
' ============================================================
Public Sub Init(ByVal env As Wv2Environment, ByVal hwnd As LongPtr)

    If m_state <> Cs_Created Then
        MsgBox "Wv2Pane.Init _________", vbCritical
        Exit Sub
    End If

    If env Is Nothing Then
        MsgBox "Wv2Pane.Init: env _ Nothing", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    If Not env.IsReady Then
        MsgBox "Wv2Pane.Init: env ___ Ready ___ (state=" & env.state & ")", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    Dim pEnv As LongPtr
    pEnv = env.EnvironmentPtr
    If pEnv = 0 Then
        MsgBox "Wv2Pane.Init: env.EnvironmentPtr _ 0", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    ' --- 1. ____ (ComCallbackHandler) _ 1 _ Acquire ---
    Set m_handler = AcquireHandlerFor(HK_ControllerCompleted, Me)
    If m_handler Is Nothing Then
        MsgBox "Wv2Pane.Init: _____ Acquire ___", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    ' --- 2. WebView2 ___ obj _______ ---
    Dim pVTableObj As LongPtr
    pVTableObj = m_handler.Slot + VTABLE_OBJ_OFFSET

    ' --- 3. m_state = Cs_Waiting _ API ____________ ---
    '   (_9.3a __________: __________)
    m_state = Cs_Waiting

    ' --- 4. Environment->CreateCoreWebView2Controller(hwnd, handler) ___ ---
    '
    '   ICoreWebView2Environment vtable _____:
    '     0 : QueryInterface
    '     1 : AddRef
    '     2 : Release
    '     3 : CreateCoreWebView2Controller(parentWindow HWND, handler ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*)
    '     4 : ... (___ SDK ________)
    '
    '   __:
    '     hwnd       : Frame _______ HWND
    '     pVTableObj : ComCallbackHandler _____ vtableObj .
    Debug.Print "Wv2Pane.Init: dcf(pEnv=" & pEnv & ", 3, ""Env.CreateController"", hwnd=" & _
                hwnd & ", handler=" & pVTableObj & ")"

    Dim hr As Long
    hr = dcf(pEnv, 3, "Env.CreateCoreWebView2Controller", hwnd, pVTableObj)

    If hr <> 0 Then
        MsgBox "CreateCoreWebView2Controller failed. HRESULT: &H" & Hex(hr), vbCritical
        m_state = Cs_Failed
        m_lastError = hr
        ' ___________9.3a / 9.3b _____ (Wv2Environment _____).
        ' Thunks_Shutdown __________________.
        Exit Sub
    End If
End Sub


' ============================================================
' OnControllerCompleted
'
'   ComCallbackHandler.Handler_Invoke _ HK_ControllerCompleted .
'   m_owner.OnControllerCompleted _______________.
'
'   Wv2Environment.OnEnvironmentCompleted ______. Public Sub .
'   (Friend __ IDispatch _____).
'
'   _ (_9.6a ____): 1 _________________________.
'   __ IF (ICoreWebView2Environment2 _) .
'   _______________ (__________, _9.6a __).
'
'   __:
'     errorCode   : HRESULT (S_OK = 0 ___)
'     pController : ICoreWebView2Controller .
'                   (WebView2 _____ AddRef _________.
'                    _______ VBA __ AddRef __. )
' ============================================================
Public Sub OnControllerCompleted( _
    ByVal errorCode As Long, _
    ByVal pController As LongPtr)

    Debug.Print "Wv2Pane.OnControllerCompleted: errorCode=" & errorCode & _
                ", pController=" & pController

    m_lastError = errorCode

    If errorCode = 0 And pController <> 0 Then
        ' VBA .
        m_pCtrl = pController
        ComAddRef m_pCtrl
        m_state = Cs_Ready
    Else
        m_pCtrl = 0
        m_state = Cs_Failed
    End If

    ' 1 ____________ (_ D1-_): .
    Set m_handler = Nothing
End Sub


' ============================================================
' .
' ============================================================
Public Property Get state() As ControllerState
    state = m_state
End Property

Public Property Get IsReady() As Boolean
    IsReady = (m_state = Cs_Ready)
End Property

Public Property Get IsFailed() As Boolean
    IsFailed = (m_state = Cs_Failed)
End Property

Public Property Get LastError() As Long
    LastError = m_lastError
End Property

' --- ________ (____ / _____) ---
Public Property Get CtrlPtr() As LongPtr
    CtrlPtr = m_pCtrl
End Property

Public Property Get ViewPtr() As LongPtr
    ViewPtr = m_pView
End Property

' --- IsPageLoaded (_9.4c ___) ---
'   View_OnNavigationCompleted _ True _______.
'   __ (UserForm _) ____________________.
'   __ SPA _________________.
Public Property Get IsPageLoaded() As Boolean
    IsPageLoaded = m_pageLoaded
End Property

' --- DocumentTitle (_9.11c ___, _______________) ---
'   View_OnDocumentTitleChanged ______________.
'   TabBar _ PushSyncToJs _ Browser.PaneAt(i).DocumentTitle ___.
'   ____________________.
Public Property Get DocumentTitle() As String
    DocumentTitle = m_title
End Property


' --- IsSettings (_9.19 ___, _________) ---
'   __ Pane __________________. Wv2Browser.OpenSettingsTab
'   _____ True ___, [__] __________________.
Public Property Get IsSettings() As Boolean
    IsSettings = m_isSettings
End Property

Public Property Let IsSettings(ByVal value As Boolean)
    m_isSettings = value
End Property


' ============================================================
' Ctrl_PutBounds (ICoreWebView2Controller::put_Bounds, vtable index 6)
'
'   Frame ________ HWND _____________ RECT ____,
'   WebView2 __________.
'
'   ___ bounds _ GetClientRect ____ RECT. _____ HWND .
'   __________ (0, 0) ___.
'
'   x64 ABI __ 16 bytes _________, .
'   VarPtr(rc) __________ WebView2 .
'   (______). dcf __ LongPtr ________.
'
'   ___: HRESULT (S_OK = 0 ___)
' ============================================================
Public Function Ctrl_PutBounds(ByRef bounds As RECT) As Long
    If m_pCtrl = 0 Then
        Ctrl_PutBounds = &H80004003   ' E_POINTER
        Exit Function
    End If
#If Win64 Then
    Ctrl_PutBounds = dcf(m_pCtrl, 6, "Ctrl.put_Bounds", VarPtr(bounds))
#Else
    Ctrl_PutBounds = dcf(m_pCtrl, 6, "Ctrl.put_Bounds", _
                         bounds.Left, bounds.Top, bounds.Right, bounds.Bottom)
#End If
End Function

Public Function Ctrl_NotifyParentWindowPositionChanged() As Long
    If m_pCtrl = 0 Then
        Ctrl_NotifyParentWindowPositionChanged = &H80004003
        Exit Function
    End If
    Ctrl_NotifyParentWindowPositionChanged = dcf(m_pCtrl, 23, "Ctrl.NotifyParentWindowPositionChanged")
End Function


' ============================================================
' Ctrl_PutIsVisible (ICoreWebView2Controller::put_IsVisible, vtable index 4)
'   _9.10c-1 _____.
'
'   Controller ___/_________. Wv2Browser .
'   (ActivateTab) _____________, ________________.
'
'   __: visible (True = __ / False = ___)
'   ___: HRESULT (S_OK = 0 ___)
'
'   ____:
'     put_IsVisible _ BOOL (__) _ 1 ____________.
'     COM ____________________, ___ Ctrl_PutBounds .
'     ________ (dcf _____) ____.
'     VARIANT_BOOL / Win32 BOOL ___: ICoreWebView2Controller::put_IsVisible .
'     ___ Win32 BOOL (= 4 byte int, True=1/False=0). VBA _ True _ -1 ___,
'     dcf _____ 1 / 0 ______ (____ 11).
' ============================================================
Public Function Ctrl_PutIsVisible(ByVal visible As Boolean) As Long
    If m_pCtrl = 0 Then
        Ctrl_PutIsVisible = &H80004003   ' E_POINTER
        Exit Function
    End If
    Dim b As Long
    If visible Then b = 1 Else b = 0
    Ctrl_PutIsVisible = dcf(m_pCtrl, 4, "Ctrl.put_IsVisible", b)
End Function


' ============================================================
' Ctrl_GetCoreWebView2 (ICoreWebView2Controller::get_CoreWebView2, vtable index 25)
'
'   ICoreWebView2 _____ out ______, m_pView _____.
'
'   COM __: get_xxx _________________ AddRef .
'   _____ (= _______ AddRef ___, Release ____).
'   m_pView _ Class_Terminate _ Release ___.
'
'   ___: HRESULT (S_OK = 0 ___)
'
'   __: _______ Init ______ 1 _________.
'         m_pView _________, .
'         _________ Release __ (__ Init .
'         ____, _____).
' ============================================================
Public Function Ctrl_GetCoreWebView2() As Long
    If m_pCtrl = 0 Then
        Ctrl_GetCoreWebView2 = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' ___ m_pView ______ Release (______)
    If m_pView <> 0 Then
        ComRelease m_pView
        m_pView = 0
    End If

    Dim pv As LongPtr
    Dim hr As Long
    hr = dcf(m_pCtrl, 25, "Ctrl.get_CoreWebView2", VarPtr(pv))
    If hr = 0 Then
        m_pView = pv
    End If
    Ctrl_GetCoreWebView2 = hr
End Function


' ============================================================
' View_Navigate (ICoreWebView2::Navigate, vtable index 5)
'
'   __ URL _____. Navigate _______ true ______,
'   _____________ HK_NavigationCompleted ____.
'
'   ___ URL _ VBA String (= UTF-16 LE) _, StrPtr _ LPCWSTR _____.
'
'   ___: HRESULT (S_OK = 0 ___)
' ============================================================
Public Function View_Navigate(ByVal url As String) As Long
    If m_pView = 0 Then
        View_Navigate = &H80004003   ' E_POINTER
        Exit Function
    End If
    ' .
    m_pageLoaded = False
    View_Navigate = dcf(m_pView, 5, "View.Navigate", StrPtr(url))
End Function


' ============================================================
' View_AddNavigationStarting (_9.4a _____, _9.6a _ View_ ________)
'
'   ICoreWebView2::add_NavigationStarting (vtable index 7) .
'   _______________.
'
'   __:
'     1. ComCallbackHandler (HK_NavigationStarting) _ Acquire (owner = Me)
'     2. dcf(m_pView, 7, handler, &token) _ WebView2 .
'     3. WebView2 _ AddRef ___ (ComCallbackHandler _ refcount ____)
'     4. ____ token _ m_persistentTokens(HK_NavigationStarting) .
'     5. ComCallbackHandler _ m_persistentHandlers(HK_NavigationStarting) .
'
'   ______ Bing .
'   ComCallbackHandler.Invoke[HK_NavigationStarting] -> Wv2Pane.View_OnNavigationStarting
'   ______________.
'
'   _____________ (______).
'
'   ___: HRESULT (S_OK = 0 ___)
' ============================================================
Public Function View_AddNavigationStarting() As Long
    If m_pView = 0 Then
        View_AddNavigationStarting = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' .
    If Not (m_persistentHandlers(HK_NavigationStarting) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddNavigationStarting: _____________"
        View_AddNavigationStarting = 0
        Exit Function
    End If

    ' ComCallbackHandler _ Acquire (owner = Me ___________________)
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_NavigationStarting, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddNavigationStarting: _____ Acquire ___"
        View_AddNavigationStarting = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    ' add_NavigationStarting ___ (vtable index 7)
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As Currency
    Dim hr As Long
    hr = dcf(m_pView, 7, "View.add_NavigationStarting", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddNavigationStarting: add_NavigationStarting __ hr=&H" & Hex(hr)
        ' ComCallbackHandler _______ (refcount=0 _______)
        ' h __________________ Set h = Nothing .
        ' ____ m_handlers(idx) ________.
        ' WebView2 _ AddRef _______________, ComCallbackHandler .
        ' _9.5 ___________________________.
        Set h = Nothing
        View_AddNavigationStarting = hr
        Exit Function
    End If

    ' __: token _ ComCallbackHandler .
    m_persistentTokens(HK_NavigationStarting) = token
    Set m_persistentHandlers(HK_NavigationStarting) = h
    Debug.Print "Wv2Pane.View_AddNavigationStarting OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddNavigationStarting = 0
End Function


' ============================================================
' View_RemoveNavigationStarting (_9.4a _____, _9.6a ___)
'
'   _____ NavigationStarting _________.
'   __________________. Class_Terminate .
'   CleanupAllPersistentHandlers ________, .
'   _______.
'
'   __:
'     1. View_RemoveEvent (vtable index 8) _ WebView2 .
'     2. WebView2 _ ComCallbackHandler _ Release -> .
'     3. Wv2Pane __ m_persistentHandlers(k) = Nothing, token = 0
'
'   ___: HRESULT (S_OK = 0 ___, _____ S_OK ______)
' ============================================================
Public Function View_RemoveNavigationStarting() As Long
    If m_pView = 0 Then
        View_RemoveNavigationStarting = &H80004003   ' E_POINTER
        Exit Function
    End If

    If m_persistentHandlers(HK_NavigationStarting) Is Nothing Then
        View_RemoveNavigationStarting = 0   ' ____ no-op
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 8, "View.remove_NavigationStarting", _
             m_persistentTokens(HK_NavigationStarting))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveNavigationStarting: ____ hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveNavigationStarting OK"
    End If

    ' ________ Wv2Pane .
    Set m_persistentHandlers(HK_NavigationStarting) = Nothing
    m_persistentTokens(HK_NavigationStarting) = 0@

    View_RemoveNavigationStarting = hr
End Function


' ============================================================
' View_OnNavigationStarting (_9.4a _____, _9.6a ___)
'   ComCallbackHandler __________________.
'
'   ComCallbackHandler.Handler_Invoke _ HK_NavigationStarting .
'   m_owner.View_OnNavigationStarting _______________.
'
'   Public Sub __ (Friend __ IDispatch::Invoke _____).
'
'   __:
'     pSender : WebView2 ___ Invoke _ 1 __ = ICoreWebView2 .
'               (= m_pView _______, ___________)
'     pArgs   : WebView2 ___ Invoke _ 2 __ =
'               ICoreWebView2NavigationStartingEventArgs .
'
'   __:
'     args.get_Uri (vtable index 3) ____ URL ________.
'     get_Uri _ LPWSTR* _ out _______, GetStringProperty .
'     PtrToString + CoTaskMemFree __________.
'
'   args _______:
'     WebView2 _ args _____ AddRef _____ Invoke ______.
'     ____________ WebView2 _ Release ____,
'     View_OnNavigationStarting __ args ____________ AddRef __.
'     URL _______________, _____.
' ============================================================
Public Sub View_OnNavigationStarting(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 __7: in-callback .
    '   ___ View_OnNavigationStarting_Core. ____ m_inCallback _ +1 ____ Core ___,
    '   *_________* -1 __. Core ____ Exit Sub _____,
    '   Core __ +1 / -1 ________________ (____ (a)).
    '   Core .
    '   (______ = __________. ____ (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnNavigationStarting_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnNavigationStarting: Core ____ " & errNum & " " & errDsc & _
                    " (in-callback ________, _______)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnNavigationStarting_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnNavigationStarting: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [__] pSender _ m_pView ______ (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs _ 0 ___ URL _______"
        Exit Sub
    End If

    ' args.get_Uri ____ URL .
    Dim url As String
    url = GetStringProperty(pArgs, 3, "NavStartingArgs.get_Uri")
    Debug.Print "  -> ___ URL: " & url
End Sub


' ============================================================
' View_AddNavigationCompleted (_9.4b _____, _9.6a ___)
'
'   ICoreWebView2::add_NavigationCompleted (vtable index 15) .
'   _______________.
'
'   View_AddNavigationStarting _______. ___:
'     - __: HK_NavigationCompleted
'     - vtable index: 15 (add) / 16 (remove)
'     - args _: ICoreWebView2NavigationCompletedEventArgs
'                (get_IsSuccess, get_WebErrorStatus, get_NavigationId ___)
' ============================================================
Public Function View_AddNavigationCompleted() As Long
    If m_pView = 0 Then
        View_AddNavigationCompleted = &H80004003   ' E_POINTER
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_NavigationCompleted) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddNavigationCompleted: _____________"
        View_AddNavigationCompleted = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_NavigationCompleted, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddNavigationCompleted: _____ Acquire ___"
        View_AddNavigationCompleted = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As Currency
    Dim hr As Long
    hr = dcf(m_pView, 15, "View.add_NavigationCompleted", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddNavigationCompleted: add_NavigationCompleted __ hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddNavigationCompleted = hr
        Exit Function
    End If

    m_persistentTokens(HK_NavigationCompleted) = token
    Set m_persistentHandlers(HK_NavigationCompleted) = h
    Debug.Print "Wv2Pane.View_AddNavigationCompleted OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddNavigationCompleted = 0
End Function


' ============================================================
' View_RemoveNavigationCompleted (_9.4b _____, _9.6a ___)
'   View_RemoveEvent (vtable index 16) ___.
'   _____ no-op (S_OK ___).
' ============================================================
Public Function View_RemoveNavigationCompleted() As Long
    If m_pView = 0 Then
        View_RemoveNavigationCompleted = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_NavigationCompleted) Is Nothing Then
        View_RemoveNavigationCompleted = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 16, "View.remove_NavigationCompleted", _
             m_persistentTokens(HK_NavigationCompleted))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveNavigationCompleted: ____ hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveNavigationCompleted OK"
    End If

    Set m_persistentHandlers(HK_NavigationCompleted) = Nothing
    m_persistentTokens(HK_NavigationCompleted) = 0@

    View_RemoveNavigationCompleted = hr
End Function


' ============================================================
' View_OnNavigationCompleted (_9.4b _____, _9.6a ___)
'   ComCallbackHandler __________________.
'
'   args = ICoreWebView2NavigationCompletedEventArgs _ get_IsSuccess
'   (vtable index 3) ______/_______.
'
'   Public Sub __ (Friend __ IDispatch::Invoke _____).
' ============================================================
Public Sub View_OnNavigationCompleted(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 __7: in-callback .
    '   ___ View_OnNavigationCompleted_Core. ____ m_inCallback _ +1 ____ Core ___,
    '   *_________* -1 __. Core ____ Exit Sub _____,
    '   Core __ +1 / -1 ________________ (____ (a)).
    '   Core .
    '   (______ = __________. ____ (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnNavigationCompleted_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnNavigationCompleted: Core ____ " & errNum & " " & errDsc & _
                    " (in-callback ________, _______)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnNavigationCompleted_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnNavigationCompleted: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [__] pSender _ m_pView ______ (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs _ 0 ___ IsSuccess _______"
        ' pArgs .
        m_pageLoaded = True
        ' --- _9.12a: ______ URL _______ (__ return ________) ---
        NotifyUrlState pSender
        Exit Sub
    End If

    ' args.get_IsSuccess .
    Dim isOk As Boolean
    isOk = GetBoolProperty(pArgs, 3, "NavCompletedArgs.get_IsSuccess")
    Debug.Print "  -> isSuccess: " & isOk

    ' ________________ (_9.4c ___)
    ' ____: __/________ NavigationCompleted ____ True ___.
    ' ______________. SPA .
    ' _____________.
    m_pageLoaded = True

    ' --- _9.12a: __ Browser _ URL _____ (_________) ---
    NotifyUrlState pSender
End Sub


' ============================================================
' NotifyUrlState (_9.12a, URL ___ Browser ___)
'
'   sender (ICoreWebView2) _____ URL / ____ / ____ _____,
'   __ Browser _ NotifyUrlChanged ___. Browser .
'   index ______ RaiseEvent TabUrlChanged __ (9.11a _ NotifyTitleChanged
'   _______).
'
'   *Pane _____________* .
'   Browser _____. ___ (____) _____________ (___).
'
'   _____:
'     - View_OnNavigationCompleted : .
'     - View_OnHistoryChanged      : SPA _ pushState / __ / .
'   _______, ______ (____ 60) __________.
' ============================================================
Private Sub NotifyUrlState(ByVal pSender As LongPtr)
    If m_browser Is Nothing Then Exit Sub
    If pSender = 0 Then Exit Sub

    Dim url As String
    url = GetStringProperty(pSender, 4, "View.get_Source")
    Dim canBack As Boolean
    canBack = GetBoolProperty(pSender, 38, "View.get_CanGoBack")
    Dim canFwd As Boolean
    canFwd = GetBoolProperty(pSender, 39, "View.get_CanGoForward")

    m_browser.NotifyUrlChanged Me, url, canBack, canFwd
End Sub


' ============================================================
' View_AddDocumentTitleChanged (_9.4b _____, _9.6a ___)
'
'   ICoreWebView2::add_DocumentTitleChanged (vtable index 46) .
'   _______________.
'
'   args _ IUnknown* (____). _____ sender.get_DocumentTitle
'   (vtable 48) ________.
' ============================================================
Public Function View_AddDocumentTitleChanged() As Long
    If m_pView = 0 Then
        View_AddDocumentTitleChanged = &H80004003
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_DocumentTitleChanged) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddDocumentTitleChanged: _____________"
        View_AddDocumentTitleChanged = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_DocumentTitleChanged, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddDocumentTitleChanged: _____ Acquire ___"
        View_AddDocumentTitleChanged = &H8000FFFF
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As Currency
    Dim hr As Long
    hr = dcf(m_pView, 46, "View.add_DocumentTitleChanged", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddDocumentTitleChanged: add_DocumentTitleChanged __ hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddDocumentTitleChanged = hr
        Exit Function
    End If

    m_persistentTokens(HK_DocumentTitleChanged) = token
    Set m_persistentHandlers(HK_DocumentTitleChanged) = h
    Debug.Print "Wv2Pane.View_AddDocumentTitleChanged OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddDocumentTitleChanged = 0
End Function


' ============================================================
' View_RemoveDocumentTitleChanged (_9.4b _____, _9.6a ___)
'   View_RemoveEvent (vtable index 47) ___.
' ============================================================
Public Function View_RemoveDocumentTitleChanged() As Long
    If m_pView = 0 Then
        View_RemoveDocumentTitleChanged = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_DocumentTitleChanged) Is Nothing Then
        View_RemoveDocumentTitleChanged = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 47, "View.remove_DocumentTitleChanged", _
             m_persistentTokens(HK_DocumentTitleChanged))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveDocumentTitleChanged: ____ hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveDocumentTitleChanged OK"
    End If

    Set m_persistentHandlers(HK_DocumentTitleChanged) = Nothing
    m_persistentTokens(HK_DocumentTitleChanged) = 0@

    View_RemoveDocumentTitleChanged = hr
End Function


' ============================================================
' View_OnDocumentTitleChanged (_9.4b _____, _9.6a ___)
'   ComCallbackHandler __________________.
'
'   sender = ICoreWebView2 _ get_DocumentTitle (vtable index 48) .
'   ______________.
'
'   args _ IUnknown* ________, sender _____________.
'
'   Public Sub __ (Friend __ IDispatch::Invoke _____).
' ============================================================
Public Sub View_OnDocumentTitleChanged(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 __7: in-callback .
    '   ___ View_OnDocumentTitleChanged_Core. ____ m_inCallback _ +1 ____ Core ___,
    '   *_________* -1 __. Core ____ Exit Sub _____,
    '   Core __ +1 / -1 ________________ (____ (a)).
    '   Core .
    '   (______ = __________. ____ (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnDocumentTitleChanged_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnDocumentTitleChanged: Core ____ " & errNum & " " & errDsc & _
                    " (in-callback ________, _______)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnDocumentTitleChanged_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnDocumentTitleChanged: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [__] pSender _ m_pView ______ (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    ' sender (ICoreWebView2).get_DocumentTitle (vtable index 48) .
    Dim title As String
    title = GetStringProperty(pSender, 48, "View.get_DocumentTitle")
    Debug.Print "  -> ____: " & title

    ' --- _9.11c: ________________ (___3, _______) ---
    m_title = title

    ' --- _9.11a: __ Browser __________ (____ m_browser __) ---
    '   NewWindowRequested ___________________.
    '   m_browser ____ (__________) _ Debug.Print .
    '   ________, ________.
    If Not m_browser Is Nothing Then
        m_browser.NotifyTitleChanged Me, title
    End If
End Sub


' ============================================================
' View_AddWebMessageReceived (_9.4c _____, _9.6a ___)
'
'   ICoreWebView2::add_WebMessageReceived (vtable index 34) .
'   _______________.
'
'   JS ___ window.chrome.webview.postMessage(text) .
'   ____. ____ Wv2Pane.View_OnWebMessageReceived _____.
' ============================================================
Public Function View_AddWebMessageReceived() As Long
    If m_pView = 0 Then
        View_AddWebMessageReceived = &H80004003
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_WebMessageReceived) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddWebMessageReceived: _____________"
        View_AddWebMessageReceived = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_WebMessageReceived, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddWebMessageReceived: _____ Acquire ___"
        View_AddWebMessageReceived = &H8000FFFF
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As Currency
    Dim hr As Long
    hr = dcf(m_pView, 34, "View.add_WebMessageReceived", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddWebMessageReceived: add_WebMessageReceived __ hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddWebMessageReceived = hr
        Exit Function
    End If

    m_persistentTokens(HK_WebMessageReceived) = token
    Set m_persistentHandlers(HK_WebMessageReceived) = h
    Debug.Print "Wv2Pane.View_AddWebMessageReceived OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddWebMessageReceived = 0
End Function


' ============================================================
' View_RemoveWebMessageReceived (_9.4c _____, _9.6a ___)
'   View_RemoveEvent (vtable index 35) ___.
' ============================================================
Public Function View_RemoveWebMessageReceived() As Long
    If m_pView = 0 Then
        View_RemoveWebMessageReceived = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_WebMessageReceived) Is Nothing Then
        View_RemoveWebMessageReceived = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 35, "View.remove_WebMessageReceived", _
             m_persistentTokens(HK_WebMessageReceived))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveWebMessageReceived: ____ hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveWebMessageReceived OK"
    End If

    Set m_persistentHandlers(HK_WebMessageReceived) = Nothing
    m_persistentTokens(HK_WebMessageReceived) = 0@

    View_RemoveWebMessageReceived = hr
End Function


' ============================================================
' View_OnWebMessageReceived (_9.4c _____, _9.6a ___)
'   ComCallbackHandler __________________.
'
'   JS _ window.chrome.webview.postMessage(text) .
'   ____. Public Sub __ (Friend __ IDispatch::Invoke _____).
'
'   args = ICoreWebView2WebMessageReceivedEventArgs _ vtable index:
'     3 : get_Source                 (___ URL, ______)
'     4 : get_WebMessageAsJson       (___, ______)
'     5 : TryGetWebMessageAsString   (JS _ string _________, ______)
'
'   TryGetWebMessageAsString _ HRESULT ____, GetStringProperty .
'   HRESULT ________________ (_______).
'   JS _ postMessage(jsonObj) .
'   __ (_____ get_WebMessageAsJson ___).
' ============================================================
Public Sub View_OnWebMessageReceived(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 __7: in-callback .
    '   ___ View_OnWebMessageReceived_Core. ____ m_inCallback _ +1 ____ Core ___,
    '   *_________* -1 __. Core ____ Exit Sub _____,
    '   Core __ +1 / -1 ________________ (____ (a)).
    '   Core .
    '   (______ = __________. ____ (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnWebMessageReceived_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnWebMessageReceived: Core ____ " & errNum & " " & errDsc & _
                    " (in-callback ________, _______)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnWebMessageReceived_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnWebMessageReceived: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [__] pSender _ m_pView ______ (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs _ 0 _______________"
        Exit Sub
    End If

    ' args.get_Source _______ URL .
    Dim source As String
    source = GetStringProperty(pArgs, 3, "WebMsgArgs.get_Source")
    Debug.Print "  -> source: " & source

    ' args.TryGetWebMessageAsString .
    ' JS ____________, .
    Dim msg As String
    msg = GetStringProperty(pArgs, 5, "WebMsgArgs.TryGetWebMessageAsString")
    Debug.Print "  -> message: " & msg

    Dim msgJson As String
    msgJson = GetStringProperty(pArgs, 4, "WebMsgArgs.get_WebMessageAsJson")

    RaiseEvent WebMessageReceived(msg, msgJson)

    ' --- _9.11c: ________ (TabBar _) _________ (_P) ---
    '   ____ OnPaneWebMessage(pane, source, msg) _________.
    '   _____________, _________________,
    '   ______ (TabBar) _________________.
    If Not m_webMsgSink Is Nothing Then
        m_webMsgSink.OnPaneWebMessage Me, source, msg
    End If
End Sub


' ============================================================
' View_AddNewWindowRequested (_9.6a _____)
'
'   ICoreWebView2::add_NewWindowRequested (vtable index 44) .
'   _______________.
'
'   ____________________ (target="_blank" ___,
'   window.open, Ctrl+_____) _____.
'
'   _9.6a _____:
'     URL ________. put_Handled _______, .
'     _______ (= OS _______, ___ OS .
'     __) _____. Bing _____________VBA .
'     _____ (_______________).
'
'   _9.6b ___ put_Handled = TRUE + __ View _ View_Navigate(uri) .
'   ______, Bing ______________________.
' ============================================================
Public Function View_AddNewWindowRequested() As Long
    If m_pView = 0 Then
        View_AddNewWindowRequested = &H80004003   ' E_POINTER
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_NewWindowRequested) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddNewWindowRequested: _____________"
        View_AddNewWindowRequested = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_NewWindowRequested, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddNewWindowRequested: _____ Acquire ___"
        View_AddNewWindowRequested = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As Currency
    Dim hr As Long
    hr = dcf(m_pView, 44, "View.add_NewWindowRequested", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddNewWindowRequested: add_NewWindowRequested __ hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddNewWindowRequested = hr
        Exit Function
    End If

    m_persistentTokens(HK_NewWindowRequested) = token
    Set m_persistentHandlers(HK_NewWindowRequested) = h
    Debug.Print "Wv2Pane.View_AddNewWindowRequested OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddNewWindowRequested = 0
End Function


' ============================================================
' View_RemoveNewWindowRequested (_9.6a _____)
'   View_RemoveEvent (vtable index 45) ___.
'   _____ no-op (S_OK ___).
' ============================================================
Public Function View_RemoveNewWindowRequested() As Long
    If m_pView = 0 Then
        View_RemoveNewWindowRequested = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_NewWindowRequested) Is Nothing Then
        View_RemoveNewWindowRequested = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 45, "View.remove_NewWindowRequested", _
             m_persistentTokens(HK_NewWindowRequested))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveNewWindowRequested: ____ hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveNewWindowRequested OK"
    End If

    Set m_persistentHandlers(HK_NewWindowRequested) = Nothing
    m_persistentTokens(HK_NewWindowRequested) = 0@

    View_RemoveNewWindowRequested = hr
End Function


' ============================================================
' View_OnNewWindowRequested (_9.6a ___, _9.6b ______)
'
'   args = ICoreWebView2NewWindowRequestedEventArgs _ vtable .
'   (__ WebView2.cls _ IDL ______):
'      3 : get_Uri(LPWSTR* uri)            <- 9.6a .
'      4 : put_NewWindow(ICoreWebView2*)
'      5 : get_NewWindow(ICoreWebView2**)
'      6 : put_Handled(BOOL)               <- 9.6b .
'      7 : get_Handled(BOOL*)
'      8 : get_IsUserInitiated(BOOL*)
'      9 : GetDeferral(ICoreWebView2Deferral**)
'     10 : get_WindowFeatures(ICoreWebView2WindowFeatures**)
'
'   _9.6b _____:
'     1. URL _ get_Uri _________ (9.6a ____)
'     2. args.put_Handled(1) __VBA .
'     3. __ View _ View_Navigate(url) .
'
'   put_Handled __________ (9.6a _____):
'     WebView2 ______ Excel ________ WebView2 __________.
'     UIA Inspect _______:
'       - ____: Chrome_WidgetWin_1
'       - Name: "Bing __" (= ______, _______)
'       - ___________________________ WebView2 .
'     ___ Edge _______________, WebView2 .
'     _____. put_Handled = TRUE _________.
'
'   ______ (_ a + _ c, ________):
'     put_Handled ___: WebView2 ______________ (9.6a ______)
'     View_Navigate ___: _________, ________ (_______)
'     ____________ Debug.Print __________.
'
'   __ (_______, _9.6c __):
'     View_Navigate _____ args.put_NewWindow ___ Wv2Pane .
'     ____. put_Handled _____. TabWebView2 __________.
'
'   Public Sub __ (Friend __ IDispatch::Invoke _____).
' ============================================================
Public Sub View_OnNewWindowRequested(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 __7: in-callback .
    '   ___ View_OnNewWindowRequested_Core. ____ m_inCallback _ +1 ____ Core ___,
    '   *_________* -1 __. Core ____ Exit Sub _____,
    '   Core __ +1 / -1 ________________ (____ (a)).
    '   Core .
    '   (______ = __________. ____ (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnNewWindowRequested_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnNewWindowRequested: Core ____ " & errNum & " " & errDsc & _
                    " (in-callback ________, _______)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnNewWindowRequested_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnNewWindowRequested: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [__] pSender _ m_pView ______ (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs _ 0 ___ URL _______"
        Exit Sub
    End If

    ' --- 1. args.get_Uri (vtable index 3) ____ URL ___ (9.6a ____) ---
    Dim url As String
    url = GetStringProperty(pArgs, 3, "NewWindowReqArgs.get_Uri")
    Debug.Print "  -> _________ URL: " & url

    ' --- _9.26/9.27: NewWindowRequested ______________ ---
    '   m_browser _ Nothing (__ Pane __ = SetBrowser ___) .
    '   ____________, __ LEGACY __ (______) ___.
    Dim nwMode As Long
    nwMode = NWM_LEGACY
    If Not (m_browser Is Nothing) Then nwMode = m_browser.NewWindowMode

    If nwMode = NWM_POPUP Then
        ' *___ 1 (9.26 _____)* put_Handled ____, .
        '   Exit __. ______________, .
        '   _______ (POST _______) ______ = ____ 23.
        '   *___* ________________ (______).
        Debug.Print "  [9.26 _____] put_Handled _______________"
        Debug.Print "    (_________________________ = POST __)"
        Debug.Print "    ____________. _______________. "
        Exit Sub

    ElseIf nwMode = NWM_PREWARM Then
        ' *___ 2 (9.27 ___)* __________ CoreWebView2 .
        '   args.put_NewWindow ___. ________ WebView .
        '   _____, POST _____/___/Referer _______________.
        Dim pNew As LongPtr
        pNew = m_browser.ClaimPrewarmViewPtr()

        If pNew = 0 Then
            ' _____ / _ Ready / __ Claim __. ____________.
            Debug.Print "  *[9.27 ______] _________________________*"
            Debug.Print "    -> _____ GET ____ (POST _____ 405 ___). _______. "
            m_browser.RequestPrewarmAsync
            ' _ ______ (put_Handled + RequestNewTabAsync) .
        Else
            Dim hrNw As Long
            hrNw = dcf(pArgs, 4, "NewWindowReqArgs.put_NewWindow", pNew)

            If hrNw <> 0 Then
                ' put_NewWindow ___. ________________.
                Debug.Print "  *[9.27 ______] put_NewWindow __ hr=&H" & Hex(hrNw) & _
                            " -> ______________*"
                m_browser.ReleasePrewarmClaim
                ' _ .
            Else
                Debug.Print "  [9.27 ______] put_NewWindow OK (_____ View=" & pNew & ")"

                ' put_Handled ____ Browser ____________ (__5 ___).
                If m_browser.SetHandledWithNewWindow Then
                    hrNw = dcf(pArgs, 6, "NewWindowReqArgs.put_Handled", 1&)
                    If hrNw <> 0 Then
                        Debug.Print "    put_Handled __ hr=&H" & Hex(hrNw)
                    Else
                        Debug.Print "    put_Handled OK (put_NewWindow ___)"
                    End If
                Else
                    Debug.Print "    put_Handled _____ (SetHandledWithNewWindow = False)"
                End If

                ' __ (m_tabs ___ -> TabAdded -> ActivateTab) ____ Timer ____.
                ' *_______* _____________ (____ 14 ___).
                m_browser.RequestAdoptPrewarmedAsync url
                Exit Sub
            End If
        End If
    End If

    ' --- _9.6b ___: __ View _____ ---
    If url = "" Then
        Debug.Print "  url _____ put_Handled / View_Navigate _____"
        Exit Sub
    End If

    ' --- 2. put_Handled = TRUE ---
    '   WebView2 _______VBA __________.
    '   __ 1& _ BOOL _ TRUE (Win32 __ BOOL = Long, & .
    '   ____ Long ________, ____ 29 ____).
    Dim hr As Long
    hr = dcf(pArgs, 6, "NewWindowReqArgs.put_Handled", 1&)
    If hr <> 0 Then
        Debug.Print "  put_Handled __ hr=&H" & Hex(hr) & _
                    " (WebView2 ________________)"
    Else
        Debug.Print "  put_Handled OK (_____________)"
    End If

    ' --- 3. ______ (_9.10e, _A + _Y) ---
    '   9.6b ____ View _ View_Navigate____Wv2Browser .
    '   __. m_browser _______ AddTabWithUrl ___ (_Y).
    '   *window.open _____ null ___* put_NewWindow __ View .
    '     ____ (_A) ___. ___ (put_NewWindow + Deferral) _______.
    '   m_browser ___ (SetBrowser ___) ____, ________ View .
    '     View_Navigate __________ (__ Pane _________).
    If m_browser Is Nothing Then
        Debug.Print "  [__] m_browser ___ (SetBrowser ___). __ View Navigate _" & _
                    "_______. "
        hr = View_Navigate(url)
        If hr <> 0 Then Debug.Print "  View_Navigate __ hr=&H" & Hex(hr)
        Exit Sub
    End If

    ' *________ (__ Timer)* ___ m_browser.AddTabWithUrl ______,
    '   AddTab _ WaitPaneReady __ Controller _ ControllerCompleted ____,
    '   _____ (NewWindowRequested) _ return ____ WebView2 .
    '   ___________ (______) __ (____ 14). .
    '   RequestNewTabAsync ____, ______ Timer .
    '   AddTabWithUrl ______.
    m_browser.RequestNewTabAsync url
    Debug.Print "  ______ Timer ___ (RequestNewTabAsync): " & url
End Sub


' ============================================================
' SetBrowser / ClearBrowserRef (_9.10e, Wv2Browser __________)
'
'   SetBrowser: Wv2Browser.AddTab ___________ (Me) ___. .
'     Wv2Pane _ View_OnNewWindowRequested __ m_browser.AddTabWithUrl ____.
'   ClearBrowserRef: Wv2Browser.Shutdown _____ (Browser.m_tabs ? Pane.m_browser)
'     ________. _______ UserForm _ m_browser=Nothing ___,
'     _____ Browser _ Pane ___________ (____ 58).
'
'   *____ 58* VBA _______ GC _______________.
'     ___________________ Shutdown_________.
' ============================================================
Public Sub SetBrowser(ByVal browser As Wv2Browser)
    Set m_browser = browser
End Sub

Public Sub ClearBrowserRef()
    Set m_browser = Nothing
    Set m_webMsgSink = Nothing   ' _9.11c: _____________ (____ 58)
End Sub


' ============================================================
' SetWebMsgSink / ClearWebMsgSink (_9.11c, WebMessage _________, _P)
'
'   SetWebMsgSink: JS ___ postMessage ____________, .
'     __________ OnPaneWebMessage(pane, source, msg) _____.
'     TabBar _ Init _______________. .
'     ____ View_AddWebMessageReceived _____ (___ 9.4c ____).
'   ClearWebMsgSink: ______. ClearBrowserRef _______ (___).
'
'   *m_browser (NewWindowRequested ______) _____* .
'     WebMessage ____. _________________.
' ============================================================
Public Sub SetWebMsgSink(ByVal sink As Object)
    Set m_webMsgSink = sink
End Sub

Public Sub ClearWebMsgSink()
    Set m_webMsgSink = Nothing
End Sub


' ============================================================
' View_AddHistoryChanged (_9.9b _____, SPA __)
'
'   ICoreWebView2::add_HistoryChanged (vtable index 13) .
'   _______________. __ ICoreWebView2 .
'   m_pView ___ (m_pView2 ___).
'
'   _______:
'     SPA _ history.pushState / replaceState / back / forward .
'     _______________. ________ (Navigate) .
'     ____________. SPA ________________.
'
'   * ___ View_AddNavigationStarting ____________ *
'     (____ 45: COM _______ WebView2 .
'      pVTableObj = h.Slot + VTABLE_OBJ_OFFSET _____.
'      View_PostWebMessageAsString ________________)
'
'   __:
'     - __: HK_HistoryChanged
'     - vtable index: 13 (add) / 14 (remove)
'     - args _: IUnknown* (____. sender ___ get_Source ____)
'
'   ___: HRESULT (S_OK = 0 ___)
' ============================================================
Public Function View_AddHistoryChanged() As Long
    If m_pView = 0 Then
        View_AddHistoryChanged = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' .
    If Not (m_persistentHandlers(HK_HistoryChanged) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddHistoryChanged: _____________"
        View_AddHistoryChanged = 0
        Exit Function
    End If

    ' ComCallbackHandler _ Acquire (owner = Me ___________________)
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_HistoryChanged, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddHistoryChanged: _____ Acquire ___"
        View_AddHistoryChanged = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    ' add_HistoryChanged ___ (vtable index 13)
    '   __: WebView2 _______ COM .
    '   h.Slot ________ h.Slot + VTABLE_OBJ_OFFSET (____ 45).
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As Currency
    Dim hr As Long
    hr = dcf(m_pView, 13, "View.add_HistoryChanged", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddHistoryChanged: add_HistoryChanged __ hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddHistoryChanged = hr
        Exit Function
    End If

    ' __: token _ ComCallbackHandler .
    m_persistentTokens(HK_HistoryChanged) = token
    Set m_persistentHandlers(HK_HistoryChanged) = h
    Debug.Print "Wv2Pane.View_AddHistoryChanged OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddHistoryChanged = 0
End Function


' ============================================================
' View_RemoveHistoryChanged (_9.9b _____)
'   View_RemoveEvent (vtable index 14) ___. __ IF ___ m_pView.
'   _____ no-op (S_OK ___).
' ============================================================
Public Function View_RemoveHistoryChanged() As Long
    If m_pView = 0 Then
        View_RemoveHistoryChanged = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_HistoryChanged) Is Nothing Then
        View_RemoveHistoryChanged = 0   ' ____ no-op
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 14, "View.remove_HistoryChanged", _
             m_persistentTokens(HK_HistoryChanged))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveHistoryChanged: ____ hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveHistoryChanged OK"
    End If

    Set m_persistentHandlers(HK_HistoryChanged) = Nothing
    m_persistentTokens(HK_HistoryChanged) = 0@

    View_RemoveHistoryChanged = hr
End Function


' ============================================================
' View_OnHistoryChanged (_9.9b _____)
'   ComCallbackHandler __________________.
'
'   args _ IUnknown* ________, sender (= ICoreWebView2) .
'   ___________ (DocumentTitleChanged _______).
'
'   ______ (_ N2-a):
'     - get_Source        (vtable 4)  : ___ URL
'     - get_CanGoBack     (vtable 38) : .
'     - get_CanGoForward  (vtable 39) : .
'   SPA _ pushState ___ URL ____ CanGoBack _ True ___, .
'   SPA ________ 1 ___________.
'
'   Public Sub __ (Friend __ IDispatch::Invoke _____).
' ============================================================
Public Sub View_OnHistoryChanged(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 __7: in-callback .
    '   ___ View_OnHistoryChanged_Core. ____ m_inCallback _ +1 ____ Core ___,
    '   *_________* -1 __. Core ____ Exit Sub _____,
    '   Core __ +1 / -1 ________________ (____ (a)).
    '   Core .
    '   (______ = __________. ____ (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnHistoryChanged_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnHistoryChanged: Core ____ " & errNum & " " & errDsc & _
                    " (in-callback ________, _______)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnHistoryChanged_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' _9.9c (_): pArgs=0 _ HistoryChanged ___ (args _ IUnknown* _____).
    ' sender ___ get_Source / get_CanGoBack / get_CanGoForward ______.
    Debug.Print "Wv2Pane.View_OnHistoryChanged: pSender=" & pSender & _
                ", pArgs=" & pArgs & " (args=0 ___)"
    If pSender <> m_pView Then
        Debug.Print "  [__] pSender _ m_pView ______ (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    ' sender (ICoreWebView2) .
    Dim url As String
    url = GetStringProperty(pSender, 4, "View.get_Source")
    Dim canBack As Boolean
    canBack = GetBoolProperty(pSender, 38, "View.get_CanGoBack")
    Dim canFwd As Boolean
    canFwd = GetBoolProperty(pSender, 39, "View.get_CanGoForward")

    ' _9.9c (_): NavigateToString ____ about:blank _______.
    ' _____URL _______________________.
    If url = "about:blank" Then
        Debug.Print "  -> URL: about:blank (NavigateToString ______, _9.9b ____ 6)"
    Else
        Debug.Print "  -> URL: " & url
    End If
    Debug.Print "  -> CanGoBack: " & canBack & ", CanGoForward: " & canFwd

    ' --- _9.12a: __ Browser _ URL _____ (SPA _ pushState / __ / __) ---
    '   ______ url / canBack / canFwd ________, NotifyUrlState .
    '   ________ (get_* _________).
    If Not m_browser Is Nothing Then
        m_browser.NotifyUrlChanged Me, url, canBack, canFwd
    End If
End Sub


' ============================================================
' View2_AddDOMContentLoaded (_9.9b _____, SPA __, ICoreWebView2_2)
'
'   ICoreWebView2_2::add_DOMContentLoaded (vtable index 64) .
'   _______________. _____ IF (ICoreWebView2_2) .
'   _______, m_pView ____ m_pView2 ___.
'   (= _9.9a _____ EnsureView2 / m_pView2 _______)
'
'   _______:
'     ____ DOM _________ (DOMContentLoaded, _______ 1 _).
'     NavigationCompleted ___, DOM _________________.
'     SPA _____________ (pushState _______ = .
'     HistoryChanged ___).
'
'   * ___ View_AddHistoryChanged / View_AddNavigationStarting .
'      ___________ (____ 45) *
'     ____m_pView _____ EnsureView2() + m_pView2 _______.
'     pVTableObj = h.Slot + VTABLE_OBJ_OFFSET ________.
'
'   __: __ IF _____ View2_ ___ (_9.6a/9.9b ____).
'
'   __ (_______):
'     - __: HK_DOMContentLoaded
'     - ________: m_pView2 (EnsureView2 _____)
'     - vtable index: 64 (add) / 65 (remove)  _ m_pView2 .
'     - args _: ICoreWebView2DOMContentLoadedEventArgs
'                (get_NavigationId = UINT64, vtable 3)
'
'   ___: HRESULT (S_OK = 0 ___)
' ============================================================
Public Function View2_AddDOMContentLoaded() As Long
    ' _2 IF ___. EnsureView2 _ m_pView2 ___ (_________).
    If Not EnsureView2() Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: EnsureView2 __ (m_pView2 ____)"
        View2_AddDOMContentLoaded = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' .
    If Not (m_persistentHandlers(HK_DOMContentLoaded) Is Nothing) Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: _____________"
        View2_AddDOMContentLoaded = 0
        Exit Function
    End If

    ' ComCallbackHandler _ Acquire (owner = Me ___________________)
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_DOMContentLoaded, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: _____ Acquire ___"
        View2_AddDOMContentLoaded = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    ' add_DOMContentLoaded ___ (vtable index 64, m_pView2 _)
    '   __: WebView2 _______ COM .
    '   h.Slot ________ h.Slot + VTABLE_OBJ_OFFSET (____ 45).
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As Currency
    Dim hr As Long
    hr = dcf(m_pView2, 64, "View2.add_DOMContentLoaded", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: add_DOMContentLoaded __ hr=&H" & Hex(hr)
        Set h = Nothing
        View2_AddDOMContentLoaded = hr
        Exit Function
    End If

    ' __: token _ ComCallbackHandler .
    m_persistentTokens(HK_DOMContentLoaded) = token
    Set m_persistentHandlers(HK_DOMContentLoaded) = h
    Debug.Print "Wv2Pane.View2_AddDOMContentLoaded OK: token=" & token & ", handler.slot=" & h.Slot

    View2_AddDOMContentLoaded = 0
End Function


' ============================================================
' View2_RemoveDOMContentLoaded (_9.9b _____, ICoreWebView2_2)
'   remove_DOMContentLoaded (vtable index 65) ___. _2 IF ___ m_pView2.
'   _____ no-op (S_OK ___).
' ============================================================
Public Function View2_RemoveDOMContentLoaded() As Long
    ' _2 IF ___. ________, __ m_pView2 ________.
    ' ____ EnsureView2 ___ (_____ 0 ______).
    If Not EnsureView2() Then
        Debug.Print "Wv2Pane.View2_RemoveDOMContentLoaded: m_pView2 ____"
        View2_RemoveDOMContentLoaded = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_DOMContentLoaded) Is Nothing Then
        View2_RemoveDOMContentLoaded = 0   ' ____ no-op
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView2, 65, "View2.remove_DOMContentLoaded", _
             m_persistentTokens(HK_DOMContentLoaded))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View2_RemoveDOMContentLoaded: ____ hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View2_RemoveDOMContentLoaded OK"
    End If

    Set m_persistentHandlers(HK_DOMContentLoaded) = Nothing
    m_persistentTokens(HK_DOMContentLoaded) = 0@

    View2_RemoveDOMContentLoaded = hr
End Function


' ============================================================
' View2_OnDOMContentLoaded (_9.9b _____, ICoreWebView2_2)
'   ComCallbackHandler __________________.
'
'   args = ICoreWebView2DOMContentLoadedEventArgs _ get_NavigationId
'   (vtable index 3) ____ NavigationId _______.
'   NavigationId _ UINT64 ___ GetLongLongProperty ___ (64bit __).
'
'   __: __ IF _____ View2_On ___.
'
'   Public Sub __ (Friend __ IDispatch::Invoke _____).
'
'   * 32bit _____ *
'     GetLongLongProperty _ #If Win64 _____ 64bit ______.
'     32bit ______ NavigationId ____________.
' ============================================================
Public Sub View2_OnDOMContentLoaded(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 __7: in-callback .
    '   ___ View2_OnDOMContentLoaded_Core. ____ m_inCallback _ +1 ____ Core ___,
    '   *_________* -1 __. Core ____ Exit Sub _____,
    '   Core __ +1 / -1 ________________ (____ (a)).
    '   Core .
    '   (______ = __________. ____ (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View2_OnDOMContentLoaded_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View2_OnDOMContentLoaded: Core ____ " & errNum & " " & errDsc & _
                    " (in-callback ________, _______)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View2_OnDOMContentLoaded_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View2_OnDOMContentLoaded: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [__] pSender _ m_pView ______ (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs _ 0 ___ NavigationId _______"
        Exit Sub
    End If

    ' args.get_NavigationId (vtable index 3, UINT64) .
    Dim navId As Variant
    navId = GetLongLongProperty(pArgs, 3, "DOMContentLoadedArgs.get_NavigationId")
    Debug.Print "  -> navigationId: " & navId
End Sub


' ============================================================
' View_PostWebMessageAsString (_9.4c _____)
'
'   ICoreWebView2::PostWebMessageAsString (vtable index 33) .
'   JS ___________.
'
'   JS ________, ________ HTML ________:
'     window.chrome.webview.addEventListener('message', event => {
'         console.log(event.data);   // .
'     });
'
'   __: WebView2 ___________________.
'         ____ NavigationCompleted ___ (= _______________)
'         _ PostMessage ____, ______________.
'         __ NavigationCompleted __________.
' ============================================================
Public Function View_PostWebMessageAsString(ByVal webMessage As String) As Long
    If m_pView = 0 Then
        View_PostWebMessageAsString = &H80004003
        Exit Function
    End If
    View_PostWebMessageAsString = dcf(m_pView, 33, _
                                     "View.PostWebMessageAsString", _
                                     StrPtr(webMessage))
End Function


' ============================================================
' _____________ (_9.4b _____)
'
'   ICoreWebView2 _______________ Public _____.
'   ____ UserForm ___ m_pane.View_xxx() ____.
'
'   ______ vtable index __________:
'     4  : get_Source
'     6  : NavigateToString
'     40 : GoBack
'     41 : GoForward
'     48 : get_DocumentTitle
' ============================================================

' --- HTML ______ WebView2 ______ ---
Public Function View_NavigateToString(ByVal htmlContent As String) As Long
    If m_pView = 0 Then
        View_NavigateToString = &H80004003
        Exit Function
    End If
    ' .
    m_pageLoaded = False
    View_NavigateToString = dcf(m_pView, 6, "View.NavigateToString", StrPtr(htmlContent))
End Function

' --- VBA _______ JS _____ (_9.18, AddHostObjectToScript, vtable 49) ---
'   objName  : JS __ window.chrome.webview.hostObjects.<objName> .
'   instance : ____ VBA _________ (Instancing 1-Private ___)
'
'   *_ WebView2.cls _________________*
'     Variant __________, __ VarPtr __2__ (VARIANT*) ___.
'     _1______________ LPCWSTR. ___ View_Navigate (StrPtr) .
'     get_Settings (VarPtr) ________.
'
'   *_______ (__A)* ______ NavigateToString _______.
'     ___ Navigate _____, ______ JS _____ hostObjects .
'     ___, _____. ___ undefined _______ ready .
'     __ (__B) __________.
Public Function View_AddHostObjectToScript(ByVal objName As String, ByVal instance As Object) As Long
    If m_pView = 0 Then
        View_AddHostObjectToScript = &H80004003   ' E_POINTER
        Exit Function
    End If
    Dim vObj As Variant
    Set vObj = instance
    View_AddHostObjectToScript = dcf(m_pView, 49, "View.AddHostObjectToScript", _
                                     StrPtr(objName), VarPtr(vObj))
End Function

' --- ___ 1 ___ ---
Public Function View_GoBack() As Long
    If m_pView = 0 Then
        View_GoBack = &H80004003
        Exit Function
    End If
    View_GoBack = dcf(m_pView, 40, "View.GoBack")
End Function

' --- ___ 1 ___ ---
Public Function View_GoForward() As Long
    If m_pView = 0 Then
        View_GoForward = &H80004003
        Exit Function
    End If
    View_GoForward = dcf(m_pView, 41, "View.GoForward")
End Function

' --- __________ URL ___ ---
Public Function View_GetSource() As String
    If m_pView = 0 Then Exit Function
    View_GetSource = GetStringProperty(m_pView, 4, "View.get_Source")
End Function

' --- __________ DocumentTitle ___ ---
Public Function View_GetDocumentTitle() As String
    If m_pView = 0 Then Exit Function
    View_GetDocumentTitle = GetStringProperty(m_pView, 48, "View.get_DocumentTitle")
End Function


' ============================================================
' _9.8b _____: ICoreWebView2 __ IF _______ (View_ _)
'                  + ICoreWebView2Controller _ Ctrl_Close
'                  + ICoreWebView2Settings _________ (Settings_ _)
'
'   ____:
'     _View_ ________ (m_pView __ + GetXxxProperty / dcf) .
'     _Settings __ EnsureSettings _ m_pSettings _______,
'       Wv2Pane ____ (Wv2Settings ___________, 9.8b-A __)
'     _BOOL ___ (View_GetCanGoBack _) _ As Boolean .
'       (HRESULT ______, IsPageLoaded ___, 9.8b-C __)
'     _void _ (View_Reload, View_Stop, Ctrl_Close) _ As Long _ HRESULT .
'       (View_Navigate ___)
' ============================================================

' --- _________ (F5 __, View vtable 31) ---
Public Function View_Reload() As Long
    If m_pView = 0 Then
        View_Reload = &H80004003
        Exit Function
    End If
    ' .
    m_pageLoaded = False
    View_Reload = dcf(m_pView, 31, "View.Reload")
End Function

' --- _______________ (Esc __, View vtable 43) ---
Public Function View_Stop() As Long
    If m_pView = 0 Then
        View_Stop = &H80004003
        Exit Function
    End If
    View_Stop = dcf(m_pView, 43, "View.Stop")
End Function

' --- ________ (View vtable 38, UI _ Back ________) ---
'   ___: True __ View_GoBack .
'   HRESULT _____, ____ False .
Public Function View_GetCanGoBack() As Boolean
    If m_pView = 0 Then Exit Function
    View_GetCanGoBack = GetBoolProperty(m_pView, 38, "View.get_CanGoBack")
End Function

' --- ________ (View vtable 39, UI _ Forward ________) ---
'   ___: True __ View_GoForward .
'   HRESULT _____, ____ False .
Public Function View_GetCanGoForward() As Boolean
    If m_pView = 0 Then Exit Function
    View_GetCanGoForward = GetBoolProperty(m_pView, 39, "View.get_CanGoForward")
End Function

' --- WebView2 _________ PID ___ (View vtable 37, ___) ---
'   ___: _________ Process ID (Long)
'   _______________ WebView2 .
'   HRESULT _____, ____ 0 .
Public Function View_GetBrowserProcessId() As Long
    If m_pView = 0 Then Exit Function
    View_GetBrowserProcessId = GetLongProperty(m_pView, 37, "View.get_BrowserProcessId")
End Function


' ============================================================
' EnsureSettings (Private, ______)
'
'   m_pSettings ____ (= 0) __ View vtable 3 (get_Settings) .
'   ICoreWebView2Settings* ____, m_pSettings ________.
'   ______________.
'
'   COM ___:
'     get_Settings _ [out, retval] ________. .
'     refcount __________, VBA ____ AddRef ___.
'     m_pSettings ______ Class_Terminate ___.
'
'   ___: True __ m_pSettings ___, False .
' ============================================================
Private Function EnsureSettings() As Boolean
    If m_pSettings <> 0 Then
        EnsureSettings = True
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.EnsureSettings: m_pView ____"
        Exit Function
    End If

    Dim pSettings As LongPtr
    Dim hr As Long
    hr = dcf(m_pView, 3, "View.get_Settings", VarPtr(pSettings))
    If hr <> 0 Then
        Debug.Print "Wv2Pane.EnsureSettings: get_Settings __ hr=&H" & Hex(hr)
        Exit Function
    End If
    If pSettings = 0 Then
        Debug.Print "Wv2Pane.EnsureSettings: pSettings _ 0"
        Exit Function
    End If

    m_pSettings = pSettings
    Debug.Print "Wv2Pane.EnsureSettings: m_pSettings __ = " & m_pSettings
    EnsureSettings = True
End Function

' --- JS ________ (Settings vtable 4 = put_IsScriptEnabled) ---
'   ___ True. False ____ JS _________.
'   IsScriptEnabled _ WebView2 .
'   __, _______________ (_________).
Public Function Settings_PutIsScriptEnabled(ByVal v As Boolean) As Long
    If Not EnsureSettings() Then
        Settings_PutIsScriptEnabled = &H80004003
        Exit Function
    End If
    Dim bv As Long
    If v Then bv = 1 Else bv = 0
    Settings_PutIsScriptEnabled = dcf(m_pSettings, 4, "Settings.put_IsScriptEnabled", bv)
End Function

' --- _________________ (Settings vtable 14 = put_AreDefaultContextMenusEnabled) ---
'   ___ True. False ____ Ctrl+.
'   _______. _____________________.
Public Function Settings_PutAreDefaultContextMenusEnabled(ByVal v As Boolean) As Long
    If Not EnsureSettings() Then
        Settings_PutAreDefaultContextMenusEnabled = &H80004003
        Exit Function
    End If
    Dim bv As Long
    If v Then bv = 1 Else bv = 0
    Settings_PutAreDefaultContextMenusEnabled = dcf(m_pSettings, 14, _
                                              "Settings.put_AreDefaultContextMenusEnabled", bv)
End Function

' --- DevTools (F12) ________ (Settings vtable 12 = put_AreDevToolsEnabled) ---
'   ___ True. False ____ F12 _____ DevTools _____.
'   ______ DevTools __________.
Public Function Settings_PutAreDevToolsEnabled(ByVal v As Boolean) As Long
    If Not EnsureSettings() Then
        Settings_PutAreDevToolsEnabled = &H80004003
        Exit Function
    End If
    Dim bv As Long
    If v Then bv = 1 Else bv = 0
    Settings_PutAreDevToolsEnabled = dcf(m_pSettings, 12, "Settings.put_AreDevToolsEnabled", bv)
End Function


' ============================================================
' _9.9a _____: ICoreWebView2_2 __ IF _____ (View2_ ____)
'
'   ____ (__________):
'     __ 1: m_pView2 _____ + EnsureView2 ____ (Settings _____)
'     __ a: IID_ICoreWebView2_2 _ EnsureView2 .
'             (Wv2Thunks _ IID ________, 9.9a _ Wv2Thunks _____)
'     __ X: ____ As Boolean (EnsureSettings ___)
'
'   IID_ICoreWebView2_2 ___:
'     WebView2.h L5078:
'       EXTERN_C __declspec(selectany) const IID IID_ICoreWebView2_2 =
'           {0x9E8F0CF8,0xE670,0x4B5E,{0xB2,0xBC,0x73,0xE0,0x61,0xE3,0x18,0x4C}};
'     MIDL_INTERFACE ___: "9E8F0CF8-E670-4B5E-B2BC-73E061E3184C"
'
'   ICoreWebView2_2 __________ (__ ICoreWebView2 _ 61?67):
'     61 : add_WebResourceResponseReceived
'     62 : remove_WebResourceResponseReceived
'     63 : NavigateWithWebResourceRequest
'     64 : add_DOMContentLoaded         <- _9.9b .
'     65 : remove_DOMContentLoaded      <- _9.9b .
'     66 : get_CookieManager
'     67 : get_Environment
' ============================================================


' ============================================================
' EnsureView2 (Private, ______)
'
'   m_pView2 ____ (= 0) __ QueryInterface(IID_ICoreWebView2_2) .
'   ___ ICoreWebView2_2* ____, m_pView2 ________.
'   ______________.
'
'   COM ___:
'     QueryInterface __________ AddRef _______.
'     m_pView2 _ m_pView .
'     (= __ COM ________ IF _________, ________).
'     m_pView2 ______ Class_Terminate / Ctrl_Close ___ (____).
'
'   QueryInterface _____:
'     IUnknown::QueryInterface (vtable 0):
'       HRESULT QueryInterface(REFIID riid, void** ppvObject);
'     _ 1 __ riid _ GUID ________.
'     _ 2 __ ppvObject _ LongPtr _ out __.
'     dcf ____ VarPtr(localIID) _ VarPtr(pv) ___.
'
'   IID _ EnsureView2 ____________ (_ a __):
'     Wv2Thunks _ m_iidTable _____, _______________.
'     __ 1 ________ (__________ True ____) .
'     __________. Wv2Thunks _ FillGUID _ Public ___.
'
'   ___: True __ m_pView2 ___, False ______ (_ X, EnsureSettings ___)
' ============================================================
Private Function EnsureView2() As Boolean
    If m_pView2 <> 0 Then
        EnsureView2 = True
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.EnsureView2: m_pView ____"
        Exit Function
    End If

    ' --- IID_ICoreWebView2_2 _____ GUID _________ ---
    '   {9E8F0CF8-E670-4B5E-B2BC-73E061E3184C}
    '   FillGUID _ Wv2Thunks _ Public Sub (_9.9a _ Private -> Public __)
    Dim iidView2 As GUID
    FillGUID iidView2, "9E8F0CF8-E670-4B5E-B2BC-73E061E3184C"

    ' --- QueryInterface(riid, &ppv) ___ ---
    '   vtable index 0 _ QueryInterface (IUnknown __).
    '   __:
    '     #1: REFIID riid     -> VarPtr(iidView2)  (GUID _______)
    '     #2: void** ppvObject -> VarPtr(pv)        (___ LongPtr ___)
    '   ___: HRESULT (S_OK = 0 ___)
    Dim pv As LongPtr
    Dim hr As Long
    hr = dcf(m_pView, 0, "View.QueryInterface(ICoreWebView2_2)", _
             VarPtr(iidView2), VarPtr(pv))
    If hr <> 0 Then
        Debug.Print "Wv2Pane.EnsureView2: QueryInterface __ hr=&H" & Hex(hr)
        Exit Function
    End If
    If pv = 0 Then
        Debug.Print "Wv2Pane.EnsureView2: QueryInterface ____ pv _ 0"
        Exit Function
    End If

    m_pView2 = pv
    Debug.Print "Wv2Pane.EnsureView2: m_pView2 __ = " & m_pView2
    EnsureView2 = True
End Function


' ============================================================
' View2_Ensure (_9.9a ___, Public ______)
'
'   Immediate _______ ?UserForm1.GetActivePane.View2_Ensure .
'   ___ m_pView2 _______________.
'
'   9.9b ___ View2_AddDOMContentLoaded .
'   EnsureView2 _________, ___________________.
'
'   _ i __: ______ + Test_AllEvents_Real _______________.
'             _____ (Test_View2_Ensure_Real) ______.
'
'   ___: True __ m_pView2 ___ (____ or ______)
' ============================================================
Public Function View2_Ensure() As Boolean
    View2_Ensure = EnsureView2()
End Function


' ============================================================
' EnsureView3 (_9.10d-2a ___, ICoreWebView2_3 __ IF _____)
'
'   ICoreWebView2_3 _ QueryInterface ______ m_pView3 ________.
'   EnsureView2 (_9.9a) ______. __ m_pView __ QI .
'   (ICoreWebView2_3 : ICoreWebView2_2 : ICoreWebView2 .
'    ____________).
'
'   SetVirtualHostNameToFolderMapping (vtable 71) ________.
'
'   IID_ICoreWebView2_3 = {A0D6DF20-3B92-416D-AA0C-437A9C727857}
'     (WebView2.h: 0xA0D6DF20,0x3B92,0x416D,{0xAA,0x0C,0x43,0x7A,0x9C,0x72,0x78,0x57})
'
'   ___: True __ m_pView3 ___ (____ or ______)
' ============================================================
Private Function EnsureView3() As Boolean
    If m_pView3 <> 0 Then
        EnsureView3 = True
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.EnsureView3: m_pView ____"
        Exit Function
    End If

    ' --- IID_ICoreWebView2_3 _____ GUID _________ ---
    '   FillGUID _ Wv2Thunks _ Public Sub (_9.9a _ Public __)
    Dim iidView3 As GUID
    FillGUID iidView3, "A0D6DF20-3B92-416D-AA0C-437A9C727857"

    ' --- QueryInterface(riid, &ppv) ___ (vtable index 0) ---
    Dim pv As LongPtr
    Dim hr As Long
    hr = dcf(m_pView, 0, "View.QueryInterface(ICoreWebView2_3)", _
             VarPtr(iidView3), VarPtr(pv))
    If hr <> 0 Then
        Debug.Print "Wv2Pane.EnsureView3: QueryInterface __ hr=&H" & Hex(hr)
        Exit Function
    End If
    If pv = 0 Then
        Debug.Print "Wv2Pane.EnsureView3: QueryInterface ____ pv _ 0"
        Exit Function
    End If

    m_pView3 = pv
    Debug.Print "Wv2Pane.EnsureView3: m_pView3 __ = " & m_pView3
    EnsureView3 = True
End Function


' ============================================================
' View3_Ensure (_9.10d-2a ___, Public ______)
'   Immediate __ ?UserForm1.GetActivePane.View3_Ensure _ m_pView3 _____.
'   View2_Ensure (_9.9a) ___.
' ============================================================
Public Function View3_Ensure() As Boolean
    View3_Ensure = EnsureView3()
End Function


' ============================================================
' View3_SetVirtualHostNameToFolderMapping (_9.10d-2a ___, ICoreWebView2_3)
'
'   ICoreWebView2_3::SetVirtualHostNameToFolderMapping (vtable index 71) ___,
'   _______________________. .
'   https://<hostName>/ ___ secure origin ______________,
'   about:blank origin (____ 6) _ pushState/replaceState ________.
'
'   *____________* COM ______ (____) _ WebView2 .
'   ______, ____ 45 (VTABLE_OBJ_OFFSET) ___. View_Navigate .
'   _______________ (View_AddNavigationStarting .
'   ______).
'
'   _____ (WebView2.h):
'     HRESULT SetVirtualHostNameToFolderMapping(
'       LPCWSTR hostName,     ' ______ (_ "appassets.example")
'       LPCWSTR folderPath,   ' .
'       COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND accessKind)  ' DENY=0/ALLOW=1/DENY_CORS=2
'
'   *_________* Navigate ____________. __ View3 .
'   _____, __ EnsureView3 ___.
'
'   accessKind _ Win32 _ enum (4 byte int). .
'   (___ 1 = ALLOW ___). dcf __ Long .
'   (put_Handled 1& ___, _______).
'
'   ___: HRESULT (S_OK = 0 ___)
' ============================================================
Public Function View3_SetVirtualHostNameToFolderMapping( _
    ByVal hostName As String, _
    ByVal folderPath As String, _
    ByVal accessKind As Long) As Long

    If Not EnsureView3() Then
        Debug.Print "Wv2Pane.View3_SetVirtualHostNameToFolderMapping: EnsureView3 __"
        View3_SetVirtualHostNameToFolderMapping = &H80004003   ' E_POINTER
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView3, 71, "View3.SetVirtualHostNameToFolderMapping", _
             StrPtr(hostName), StrPtr(folderPath), accessKind)

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View3_SetVirtualHostNameToFolderMapping: __ hr=&H" & _
                    Hex(hr) & " (host=" & hostName & ", folder=" & folderPath & ")"
    Else
        Debug.Print "Wv2Pane.View3_SetVirtualHostNameToFolderMapping OK: https://" & _
                    hostName & "/ -> " & folderPath & " (accessKind=" & accessKind & ")"
    End If

    View3_SetVirtualHostNameToFolderMapping = hr
End Function


' ============================================================
' Ctrl_Close (Controller vtable 24)
'
'   WebView2 controller ________. Controller ____ VBA .
'   ______________.
'
'   _9.8c ____ (_ E 1+3 ______):
'     Close ___________, _____:
'       (1) ____ Controller::Close ___ (vtable 24)
'       (2) m_pSettings / m_pView2 / m_pView / m_pCtrl .
'           VBA ___ Release ___, ____ 0 .
'           (_9.9a _ m_pView2 ___, Release __________ = ____ 43)
'       (3) m_persistentTokens _ 0 ____ (Cleanup ___ remove __)
'           _ _______ WebView2 __ Close ____ Release _______,
'              VBA __ remove_xxx _________ (= __ refCount=0 .
'              _____ remove_xxx ____ AV ____)
'       (4) m_closed = True ____ (View_xxx / Settings_xxx ___ return)
'       (5) m_state = Cs_Failed (Init ______________)
'
'   _____:
'     - Close __ Class_Terminate _____, CleanupAllPersistentHandlers .
'       m_persistentTokens ____ 0 _______ (= __)
'     - Release _ m_pSettings/m_pView2/m_pView/m_pCtrl ____ 0 .
'     - Close _________ View_xxx _____ m_closed ___ return
'
'   __:
'     - Ctrl_Close _____ Controller::Close ___ HRESULT
'       (= ___ VBA ____________, __________)
'     - __ Wv2Pane _________ Init .
'       (= _________ New Wv2Pane ___________)
' ============================================================
Public Function Ctrl_Close() As Long
    If m_closed Then
        Debug.Print "Wv2Pane.Ctrl_Close: __ Close __, __"
        Ctrl_Close = 0  ' __ Close ______ S_OK .
        Exit Function
    End If
    If m_pCtrl = 0 Then
        Ctrl_Close = &H80004003
        Exit Function
    End If

    Debug.Print "Wv2Pane.Ctrl_Close: m_pCtrl = " & m_pCtrl
    Dim hr As Long
    hr = dcf(m_pCtrl, 24, "Ctrl.Close")

    ' --- (2) (3) (4) (5): Close _______ ---
    '
    ' Close _______ WebView2 ________________ Release .
    ' ___ (_______, _9.8b ____Ctrl_Close _____).
    ' _____ VBA __ m_pSettings/m_pView2/m_pView/m_pCtrl _ Release ____,
    ' VBA __ AddRef ______.
    '
    ' ___ Class_Terminate ___ Settings -> View2 -> View -> Ctrl (_______,
    ' _9.9a _ m_pView2 ___, ____ 43).
    If m_pSettings <> 0 Then
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pSettings = " & m_pSettings
        ComRelease m_pSettings
        m_pSettings = 0
    End If
    If m_pView3 <> 0 Then
        ' _9.10d-2a ___: View3 ____ IF, m_pView2 ____ Release (____ 51).
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pView3 = " & m_pView3
        ComRelease m_pView3
        m_pView3 = 0
    End If
    If m_pView2 <> 0 Then
        ' _9.9a ___: View2 _ View __ QI ________ IF, View ____ Release.
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pView2 = " & m_pView2
        ComRelease m_pView2
        m_pView2 = 0
    End If
    If m_pView <> 0 Then
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pView = " & m_pView
        ComRelease m_pView
        m_pView = 0
    End If
    If m_pCtrl <> 0 Then
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pCtrl = " & m_pCtrl
        ComRelease m_pCtrl
        m_pCtrl = 0
    End If

    ' (3) _______ token _ 0 ____:
    '   WebView2 _ Close ____ Release ________,
    '   ___ remove_xxx _________________ remove .
    '   _______. token _ 0 ______ CleanupAllPersistentHandlers .
    '   ______ (token=0 _________________, _9.4a).
    '
    '   _9.9b: ___________ LBound/UBound ___.
    '   _____ HK_DOMContentLoaded .
    '   (= CleanupAllPersistentHandlers ______________).
    Dim k As Long
    For k = LBound(m_persistentHandlers) To UBound(m_persistentHandlers)
        m_persistentTokens(k) = 0@
        Set m_persistentHandlers(k) = Nothing
    Next k

    ' (4) m_closed .
    m_closed = True
    m_state = Cs_Failed   ' Close .

    Ctrl_Close = hr
End Function


' ============================================================
' _9.8c _____: View_ExecuteScript + OnExecuteScriptCompleted
'                  (_ a = ____ OK, _ _ = ____________,
'                   _ X = __ callback ID __)
' ============================================================


' --- ___ JavaScript _ WebView2 ____ (View vtable 29) ---
'
'   __:
'     jsCode  : ____ JS ___ (String)
'
'   ___ (As Long):
'     >= 1    : _____ callback ID (_________)
'     <= 0    : ___ (m_pView = 0, Close __, dcf ___)
'                 -1 -> m_closed (Close __)
'                 -2 -> m_pView = 0 (View ___)
'                 -3 -> Acquire .
'                 -4 -> dcf __ (________ HRESULT ___,
'                       ___ Debug.Print ___)
'
'   __:
'     1. callback ID ___ (m_executeScriptCounter _ 1 ___)
'     2. AcquireHandlerFor(HK_ExecuteScriptCompleted, Me) .
'     3. _____ callbackId ___ (CallbackId _____ Let)
'     4. View::ExecuteScript(jsCode, handler.Slot) .
'     5. _____ h _ Nothing ___ (= Wv2Pane _ ComCallbackHandler ______,
'        WebView2 _________ AddRef __________)
'     6. callback ID .
'
'   __:
'     WebView2 ______ JS __ -> ____ handler.Invoke .
'     -> ComCallbackHandler.Invoke[HK_ExecuteScriptCompleted] .
'     -> __________ Wv2Pane.OnExecuteScriptCompleted .
'     -> WebView2 _ handler _ Release -> refCount=0 -> auto-released
'
'   _____:
'     ____ View_ExecuteScript _ 10 ________, .
'     ____ + ____ callback ID __________ (_ a ___).
Public Function View_ExecuteScript(ByVal jsCode As String) As Long
    If m_closed Then
        Debug.Print "Wv2Pane.View_ExecuteScript: Close _______"
        View_ExecuteScript = -1
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.View_ExecuteScript: m_pView ____"
        View_ExecuteScript = -2
        Exit Function
    End If

    ' --- 1. callback ID ___ ---
    m_executeScriptCounter = m_executeScriptCounter + 1
    Dim cbId As Long
    cbId = m_executeScriptCounter

    ' --- 2. _____ Acquire ---
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_ExecuteScriptCompleted, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_ExecuteScript: _____ Acquire __"
        View_ExecuteScript = -3
        Exit Function
    End If

    ' --- 3. callback ID ___ ---
    h.callbackId = cbId

    ' --- 4. ExecuteScript ___ (vtable 29) ---
    '   _____: HRESULT ExecuteScript(LPCWSTR javaScript, IExecuteScriptCompletedHandler*)
    '
    '   __: WebView2 ________ COM .
    '   h.Slot ________, h.Slot + VTABLE_OBJ_OFFSET.
    '   _____________, _____ VTable .
    '   ______________ (= ___ add_NavigationStarting .
    '   ______, ____ 8 _).
    '
    '   _______ WebView2 ___________ vtable .
    '   ________.
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim hr As Long
    hr = dcf(m_pView, 29, "View.ExecuteScript", StrPtr(jsCode), pVTableObj)

    ' --- 5. _________ (Wv2Pane ______) ---
    '   WebView2 _ AddRef _______________, _______ OK.
    '   _____ m_owner (= Me Wv2Pane) _________.
    Set h = Nothing

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_ExecuteScript: dcf __ hr=&H" & Hex(hr) & _
                    " (callbackId=" & cbId & " ___, ___________)"
        View_ExecuteScript = -4
        Exit Function
    End If

    Debug.Print "Wv2Pane.View_ExecuteScript OK: callbackId=" & cbId & _
                ", jsLen=" & Len(jsCode)
    View_ExecuteScript = cbId
End Function


' ============================================================
' View_ExecuteScriptSafe (_9.9c _____, JS ____________)
'
'   _________ JS __ try-catch ______ ExecuteScript ___.
'   _____ resultJson ______________ JSON ___:
'     ___: {"ok":true, "value":<JS ____>}
'     ___: {"ok":false, "errorName":"<___>", "errorMessage":"<_____>",
'              "stack":"<________>"}
'
'   _____, 9.8c ______JS ____ HRESULT=S_OK + result=null.
'   __________ VBA ________________.
'
'   ____ (__________):
'     _F1: ___ View_ExecuteScriptSafe (__, Safe = try-catch ___)
'     _F2: VBA ___ JSON ______ OnExecuteScriptCompleted .
'           (VBA _ JSON ____ 64bit _______. ____ InStr ______)
'     _F3: ______ (________ " ____________)
'
'   * __ (_______) *
'     ___________________________.
'     (= var __r = (USER_CODE) _____________)
'
'     ____ NG / OK __:
'       NG: "throw new Error('x')"   <- throw __ (statement). .
'                                       ______, wrappedJs .
'                                       ________, ExecuteScript .
'                                       resultJson = "null" .
'                                       (= try-catch ______)
'       OK: "(()=>{ throw new Error('x') })()"
'                                     <- .
'                                       (____ throw ___________,
'                                        catch _________)
'       OK: "(function(){ var x=1; return x+2; })()"
'                                     <- ___ + return .
'       OK: "nonexistentFunction()"   <- _ (______). .
'                                       ReferenceError _ catch .
'
'     * ____ (_9.9c) _______ *
'       VBA _________________, JS .
'       ExecuteScript _ HRESULT=S_OK + result="null" ______.
'       _____ (ReferenceError _) _______. .
'       ____________, _____________.
'       (______ catch ______ eval() _________)
'
'     ________ " _____________ \" .
'       (VBA _ "" _________ JS _ \" _____).
'
'   ___:
'     cbid = m_pane.View_ExecuteScriptSafe("1 + 2")
'       -> resultJson = "{\"ok\":true,\"value\":3}"
'     cbid = m_pane.View_ExecuteScriptSafe("(()=>{throw new Error('boom')})()")
'       -> resultJson = "{\"ok\":false,\"errorName\":\"Error\",
'                         \"errorMessage\":\"boom\",\"stack\":\"...\"}"
'       _ "throw new Error('boom')" ______ resultJson="null" .
'         (______)
'     cbid = m_pane.View_ExecuteScriptSafe("nonexistentFunction()")
'       -> resultJson = "{\"ok\":false,\"errorName\":\"ReferenceError\",...}"
'
'   ___: View_ExecuteScript ___ (callback ID _________)
' ============================================================
Public Function View_ExecuteScriptSafe(ByVal userCode As String) As Long
    ' ________ try-catch _______ JSON ___ JS _______.
    ' ______ 1 ____ (ExecuteScript ____________________ < ___).
    Dim wrappedJs As String
    wrappedJs = "(function(){try{var __r=(" & userCode & ");" & _
                "return {ok:true,value:__r};}" & _
                "catch(e){return {ok:false,errorName:e.name," & _
                "errorMessage:e.message,stack:e.stack};}})()"

    Debug.Print "Wv2Pane.View_ExecuteScriptSafe: wrappedJsLen=" & Len(wrappedJs)
    View_ExecuteScriptSafe = View_ExecuteScript(wrappedJs)
End Function


' --- ExecuteScript __________ (______________) ---
'
'   ComCallbackHandler.Handler_Invoke[HK_ExecuteScriptCompleted] .
'   m_owner.OnExecuteScriptCompleted _____. Public .
'   ______ Public __ (= ____ 14 _).
'
'   __:
'     callbackId  : View_ExecuteScript ______ ID
'     errorCode   : HRESULT (0 = S_OK)
'     resultJson  : JS _____ JSON .
'                   _:
'                     ExecuteScript("1 + 2")     -> "3"
'                     ExecuteScript("'foo'")     -> "\"foo\""
'                     ExecuteScript("document.title")
'                                                -> "\"__ - Microsoft Bing\""
'                     ExecuteScript("undefined") -> "null" (_ undefined _ JSON __ null)
'                     ExecuteScript("throw 1")   -> "" (errorCode __ 0)
'
'   _ _ ____ (_9.9c ___):
'     1. Wv2Pane ______ (= 9.8c ____, _______________)
'     2. UserForm1.OnExecuteScriptResult ____ (= 9.9c ___)
'
'   * UserForm1 _________ (_ E1-a-i) *
'     Wv2Pane _ UserForm1 _____ **_____**. VBA _ UserForm .
'     ______________________ (= UserForm1 .
'     _______) __, _____ UserForm1.OnExecuteScriptResult .
'     _______________.
'
'     ____Wv2Pane _ UserForm1 .
'     _________ (= _________________) __,
'     9.10 _ TabWebView2 .
'     ________, 9.9c ________________.
'
'     UserForm1 __ On Error Resume Next ___________,
'     UserForm1.OnExecuteScriptResult ______________.
Public Sub OnExecuteScriptCompleted( _
    ByVal callbackId As Long, _
    ByVal errorCode As Long, _
    ByVal resultJson As String)

    ' --- (1) Wv2Pane ____ (9.8c ___, _ E2-X = __) ---
    Debug.Print "Wv2Pane.OnExecuteScriptCompleted: callbackId=" & callbackId & _
                ", errorCode=" & errorCode & _
                ", resultJson=" & resultJson

    ' --- (1b) D-1: _________ (__2___5) ---
    '   EvalSync ______ callbackId _________.
    '   ______ (____________) ___________.
    '   EvalSync ______ View_ExecuteScript .
    '   _____, _________.
    Dim k As String
    Dim ent As Variant
    Dim found As Boolean

    k = "cb" & callbackId
    If Not m_evalPending Is Nothing Then
        found = False
        On Error Resume Next
        ent = m_evalPending(k)
        found = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0

        If found Then
            If ent(0) = EVAL_DISCARDED Then
                Debug.Print "Wv2Pane.OnExecuteScriptCompleted: *____* _ " & k & _
                            " ____________ (__________)"
                m_evalPending.Remove k
            Else
                m_evalPending.Remove k
                m_evalPending.Add Array(EVAL_DONE, errorCode, resultJson), k
            End If
        End If
    End If

    ' --- (2) UserForm1 ____ (_9.9c ___, _ _ ____) ---
    '   VBA ______ UserForm _____ (______, _ E1-a-i).
    '   UserForm1 _________ (_______, _____ Terminate __)
    '   ____: VBA _ UserForm1 ________________,
    '   __________________. Terminate .
    '   ______ + Initialize ___.
    '   ___, __ ExecuteScript ______ Wv2Pane .
    '   = UserForm1 ______________________,
    '   ___________________.
    RaiseEvent ExecuteScriptResult(callbackId, errorCode, resultJson)
End Sub


' ============================================================
' *D-1 ___: EvalSync _____*
' ============================================================


' ============================================================
' EvalSync (D-1) - JS .
'
'   __:
'     jsExpr     : ____ JS *_* (____. ___________)
'     timeoutSec : _______ (__ 5 _). 0 ___________.
'
'   ___ (As String):
'     __ : JS _____ JSON.stringify _____. LastEvalOk = True
'     __ : ___. ___ LastEvalError ___. LastEvalOk = False
'
'   __:
'     1. _______ (Close __ / View ___ / *in-callback*)
'     2. View_ExecuteScriptSafe _ JSON.stringify(_) ____ callbackId .
'     3. .
'     4. DoEvents _________ (WaitPaneReady ____)
'     5. __ -> JSON _______ / __ -> .
'
'   *__ JSON.stringify _____ (____ (e))*
'     Safe _____ ok _ value ____, value .
'     ______ VBA __________________. JSON.stringify .
'     ______ value _*_____*___, _____ 1 _____.
'
'   *DoEvents _________/___________ (____ 20)*
' ============================================================
Public Function EvalSync(ByVal jsExpr As String, _
                         Optional ByVal timeoutSec As Single = 5) As String
    Dim res As String
    Dim cbId As Long
    Dim k As String
    Dim t As Single
    Dim ent As Variant
    Dim timedOut As Boolean
    Dim errCode As Long
    Dim resultJson As String
    Dim v As String
    Dim eName As String
    Dim eMsg As String

    res = ""
    m_lastEvalOk = False
    m_lastEvalError = ""

    ' --- 1. ____ ---
    If m_closed Then
        m_lastEvalError = "closed"
        Wv2Log.LogW "Wv2Pane.EvalSync: Close _______"
        EvalSync = ""
        Exit Function
    End If

    If m_pView = 0 Then
        m_lastEvalError = "no-view"
        Wv2Log.LogW "Wv2Pane.EvalSync: m_pView ____"
        EvalSync = ""
        Exit Function
    End If

    If m_inCallback > 0 Then
        ' *__7 ____* __________ 14 ____________.
        m_lastEvalError = "in-callback"
        Wv2Log.LogW "Wv2Pane.EvalSync: *___________________* " & _
                    "(__=" & m_inCallback & "). ______ JS _________ " & _
                    "SafeTimer ____________ (____ 14). " & _
                    "________________________ " & _
                    "ResetCallbackGuard _______. "
        EvalSync = ""
        Exit Function
    End If

    If timeoutSec <= 0 Then timeoutSec = EVAL_TIMEOUT_DEFAULT
    If m_evalPending Is Nothing Then Set m_evalPending = New Collection

    ' --- 2. Safe _________ (__4) ---
    cbId = View_ExecuteScriptSafe("JSON.stringify(" & jsExpr & ")")
    If cbId <= 0 Then
        m_lastEvalError = "execute-failed(" & cbId & ")"
        Wv2Log.LogE "Wv2Pane.EvalSync: ExecuteScript ______ rc=" & cbId
        EvalSync = ""
        Exit Function
    End If

    ' --- 3. _________ ---
    k = "cb" & cbId
    m_evalPending.Add Array(EVAL_WAITING, 0&, ""), k

    m_evalDepth = m_evalDepth + 1
    If m_evalDepth > 1 Then
        Wv2Log.LogD "Wv2Pane.EvalSync: _____ __=" & m_evalDepth & " (" & k & ")"
    End If

    ' --- 4. ____ (WaitPaneReady ____) ---
    timedOut = False
    t = Timer
    Do
        DoEvents
        ent = m_evalPending(k)
        If ent(0) = EVAL_DONE Then Exit Do
        If (Timer - t) > timeoutSec Then
            timedOut = True
            ' *__5* ___________. __________.
            m_evalPending.Remove k
            m_evalPending.Add Array(EVAL_DISCARDED, 0&, ""), k
            Exit Do
        End If
    Loop

    m_evalDepth = m_evalDepth - 1

    If timedOut Then
        m_lastEvalError = "timeout(" & timeoutSec & "s)"
        Wv2Log.LogE "Wv2Pane.EvalSync: *______* " & k & " ________ " & _
                    "(______________)"
        EvalSync = ""
        Exit Function
    End If

    ' --- 5. ______ ---
    errCode = ent(1)
    resultJson = ent(2)
    m_evalPending.Remove k

    If errCode <> 0 Then
        m_lastEvalError = "hresult(&H" & Hex$(errCode) & ")"
        Wv2Log.LogE "Wv2Pane.EvalSync: ExecuteScript ___ " & m_lastEvalError
        EvalSync = ""
        Exit Function
    End If

    If InStr(1, resultJson, """ok"":true") > 0 Then
        If Wv2Json.JsonPickStr(resultJson, "value", v) Then
            m_lastEvalOk = True
            res = v
        Else
            ' JSON.stringify(undefined) _ undefined ___, ___ stringify .
            ' value _______. _______ undefined _____.
            m_lastEvalOk = True
            res = "undefined"
        End If

    ElseIf InStr(1, resultJson, """ok"":false") > 0 Then
        Wv2Json.JsonPickStr resultJson, "errorName", eName
        Wv2Json.JsonPickStr resultJson, "errorMessage", eMsg
        m_lastEvalError = eName & ": " & eMsg
        Wv2Log.LogW "Wv2Pane.EvalSync: JS __ " & m_lastEvalError

    Else
        ' Safe _______ ok _____, _______ JS .
        ' (ExecuteScript _ S_OK + null ______. _9.9c ___).
        m_lastEvalError = "syntax-error-or-null (resultJson=" & Left$(resultJson, 60) & ")"
        Wv2Log.LogE "Wv2Pane.EvalSync: " & m_lastEvalError
    End If

    EvalSync = res
End Function


' ============================================================
' LastEvalOk / LastEvalError (D-1) - ___ EvalSync .
'   EvalSync __________, ____JS .
'   __________________.
' ============================================================
Public Property Get LastEvalOk() As Boolean
    LastEvalOk = m_lastEvalOk
End Property

Public Property Get LastEvalError() As String
    LastEvalError = m_lastEvalError
End Property


' ============================================================
' InCallbackDepth / ResetCallbackGuard / Debug_SetInCallback (D-1 __7)
'
'   ResetCallbackGuard:
'     ___________________________________.
'     __________________, EvalSync _ in-callback .
'     _______________.
'
'   Debug_SetInCallback:
'     *____* _______ Test_D1_Guard ____________.
'     ____________ (Debug_ ____ Wv2Browser _ Debug_* _____).
' ============================================================
Public Property Get InCallbackDepth() As Long
    InCallbackDepth = m_inCallback
End Property

Public Sub ResetCallbackGuard()
    Debug.Print "Wv2Pane.ResetCallbackGuard: m_inCallback = " & m_inCallback & " -> 0"
    m_inCallback = 0
End Sub

Public Sub Debug_SetInCallback(ByVal n As Long)
    Debug.Print "Wv2Pane.Debug_SetInCallback: " & m_inCallback & " -> " & n & " (____)"
    m_inCallback = n
End Sub


' ============================================================
' JsQuote (D-2, Friend) - VBA ____ JS .
'
'   *VBA ______________ + _________________*
'     Wv2Json.JsonEscape ___________________ JSON .
'     ______, ___ Chr$(34) ___. _______*___________*
'     __, VBA __________ (_______ JS-in-VBA _____).
'
'   _:
'     JsQuote("input[name='q']")   -> "input[name='q']"          (____)
'     JsQuote("a[title=""x""]")    -> ____ JSON .
'                                    (_______*___________*)
'     JsQuote("C:\temp")           -> "C:\\temp"
'
'   Friend _______ Wv2Element.GetAttribute _____________.
' ============================================================
Friend Function JsQuote(ByVal s As String) As String
    JsQuote = Chr$(34) & Wv2Json.JsonEscape(s) & Chr$(34)
End Function


' ============================================================
' JsRegistryPrologue (D-2, Private) - .
'
'   D-2 __ JS ___________ (__1 _c).
'   ______, __ R ___________.
'   *_________, __________* = _________ stale.
'
'   *JS _________________* (________).
' ============================================================
Private Function JsRegistryPrologue() As String
    JsRegistryPrologue = _
        "if(typeof window.__wv2==='undefined'||!window.__wv2){" & _
        "window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]};}" & _
        "var R=window.__wv2;"
End Function


' ============================================================
' GetElementById / QuerySelector (D-2) - ___ 1 .
'
'   _____ Wv2Element _, ________ Nothing ___ (__4 _a).
'
'   *Nothing ____ 2 __*
'     Nothing + LastEvalOk = True  ... ________ (__)
'     Nothing + LastEvalOk = False ... __. LastEvalError .
'
'   ___:
'     Dim p As Wv2Pane, el As Wv2Element
'     Set p = UserForm1.GetActivePane
'     Set el = p.QuerySelector("input[name='q']")
'     If el Is Nothing Then
'         Debug.Print "__ (ok=" & p.LastEvalOk & " err=" & p.LastEvalError & ")"
'     Else
'         Debug.Print el.TagName & " / " & el.Value
'     End If
' ============================================================
'   *D-3 __ ImplicitWaitSec > 0 ______________*
'     ___ 0 ___, _______ D-2 ____1 ____________.
Public Function GetElementById(ByVal elementId As String) As Wv2Element
    Set GetElementById = FindWithWait("document.getElementById(" & JsQuote(elementId) & ")", _
                                      "GetElementById", elementId, m_implicitWaitSec)
End Function

Public Function QuerySelector(ByVal selector As String) As Wv2Element
    Set QuerySelector = FindWithWait("document.querySelector(" & JsQuote(selector) & ")", _
                                     "QuerySelector", selector, m_implicitWaitSec)
End Function


' ============================================================
' FindOne (D-2, Private) - .
'
'   finderExpr : ___ 1 ___ JS _ (________ null _____)
'   apiName / arg : .
'
'   JS ______:
'     {g:'<__>', found:true,  h:<__>}   .
'     {g:'<__>', found:false, h:0}        .
'
'   *h __________* Wv2Json.JsonGetNum ________ (____64).
'     found _________ h ___ 0 _____.
' ============================================================
'   quiet (D-3): True __________________. .
'                1 ________, _________________.
Private Function FindOne(ByVal finderExpr As String, _
                         ByVal apiName As String, _
                         ByVal arg As String, _
                         Optional ByVal quiet As Boolean = False) As Wv2Element
    Dim js As String
    Dim res As String
    Dim gen As String
    Dim h As Long
    Dim el As Wv2Element

    Set FindOne = Nothing

    js = "(function(){" & JsRegistryPrologue() & _
         "var e=" & finderExpr & ";" & _
         "if(!e){return {g:R.gen,found:false,h:0};}" & _
         "R.els.push(e);" & _
         "return {g:R.gen,found:true,h:R.els.length-1};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        Debug.Print "Wv2Pane." & apiName & ": __ (" & arg & ") err=" & m_lastEvalError
        Exit Function
    End If

    ' ____________ (________)
    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    If InStr(1, res, """found"":true") = 0 Then
        ' ________. *_____* ___ m_lastEvalOk _ True ___.
        If Not quiet Then
            Debug.Print "Wv2Pane." & apiName & ": ______ (" & arg & ")"
        End If
        Exit Function
    End If

    h = Wv2Json.JsonGetNum(res, "h")

    Set el = New Wv2Element
    el.InitElement Me, h, gen
    Set FindOne = el

    Debug.Print "Wv2Pane." & apiName & ": __ (" & arg & ") handle=" & h & " gen=" & gen
End Function


' ============================================================
' CurrentDomGen (D-2) - .
'
'   *_________* _JS _______________.
'   stale ________, ____________.
' ============================================================
Public Property Get CurrentDomGen() As String
    CurrentDomGen = m_domGen
End Property


' ============================================================
' ClearElementRegistry (D-2, __3) - .
'
'   JS __________________.
'   *___ Wv2Element ____ stale ___. *
'
'   ______________ els __________. .
'   (_______________).
'
'   ___: True = __. False = __ (___ LastEvalError)
' ============================================================
Public Function ClearElementRegistry() As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    js = "(function(){" & _
         "window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]};" & _
         "return window.__wv2.gen;})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        Wv2Log.LogE "Wv2Pane.ClearElementRegistry: __ err=" & m_lastEvalError
        ClearElementRegistry = False
        Exit Function
    End If

    gen = Wv2Json.JsonUnescape(res)
    m_domGen = gen
    Wv2Log.LogD "Wv2Pane.ClearElementRegistry: ___________ (_____=" & gen & ")"
    ClearElementRegistry = True
End Function


' ============================================================
' Elem_Read (D-2, Friend) - Wv2Element ______ 1 .
'
'   __:
'     elemHandle : JS _ R.els[] .
'     elemGen    : .
'     valueExpr  : __ e (____) ___ JS _. _ "e.innerText"
'     outVal     : ______ (_________)
'     outStatus  : ____. .
'   ___:
'     True  = ______ (__ null / undefined ______ True. outVal ____)
'     False = stale _____. ___ outStatus
'
'   JS ______:
'     {s:'gen', g:'<__>'}                 _____ = stale
'     {s:'nil', g:'<__>'}                 ________ = stale .
'     {s:'ok',  g:'<__>', has:false, v:''} __ null / undefined
'     {s:'ok',  g:'<__>', has:true,  v:'<_>'}
'
'   *____ String(x) ______________*
'     JSON ________ VBA _____________. ____ API .
'     ______ String ___, JS _________ (____31 _ undefined
'     _________).
' ============================================================
Friend Function Elem_Read(ByVal elemHandle As Long, _
                          ByVal elemGen As String, _
                          ByVal valueExpr As String, _
                          ByRef outVal As String, _
                          ByRef outStatus As String) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    outVal = ""
    outStatus = ""
    Elem_Read = False

    ' --- ____: _______________, _____ stale ______ ---
    '     (____________________________, _______)
    If Len(m_domGen) > 0 Then
        If elemGen <> m_domGen Then
            outStatus = "stale"
            Exit Function
        End If
    End If

    js = "(function(){" & JsRegistryPrologue() & _
         "if(R.gen!==" & JsQuote(elemGen) & "){return {s:'gen',g:R.gen};}" & _
         "var e=R.els[" & CStr(elemHandle) & "];" & _
         "if(!e){return {s:'nil',g:R.gen};}" & _
         "var x=" & valueExpr & ";" & _
         "if(x===null||x===undefined){return {s:'ok',g:R.gen,has:false,v:''};}" & _
         "return {s:'ok',g:R.gen,has:true,v:String(x)};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        outStatus = m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    If InStr(1, res, """s"":""ok""") = 0 Then
        outStatus = "stale"
        Exit Function
    End If

    If InStr(1, res, """has"":true") = 0 Then
        ' __ null / undefined. *_______* (_____ / value _______)
        outVal = ""
        Elem_Read = True
        Exit Function
    End If

    Wv2Json.JsonPickStr res, "v", outVal
    Elem_Read = True
End Function


' ============================================================
' Elem_CheckStale (D-2, Friend) - .
'
'   ___: True = stale (______)
'   outStatus: ____________. .
'
'   *___________ True (___) ___. *
'     ____________ outStatus _____.
' ============================================================
Friend Function Elem_CheckStale(ByVal elemHandle As Long, _
                                ByVal elemGen As String, _
                                ByRef outStatus As String) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    outStatus = ""
    Elem_CheckStale = True

    ' ____ (Elem_Read _____)
    If Len(m_domGen) > 0 Then
        If elemGen <> m_domGen Then Exit Function
    End If

    js = "(function(){" & JsRegistryPrologue() & _
         "var ok=(R.gen===" & JsQuote(elemGen) & ")&&(!!R.els[" & CStr(elemHandle) & "]);" & _
         "return {s:ok?'ok':'stale',g:R.gen};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        outStatus = m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    Elem_CheckStale = (InStr(1, res, """s"":""ok""") = 0)
End Function


' ============================================================
' Elem_Exec (D-3, Friend) - ____________ 1 .
'
'   __:
'     elemHandle  : JS _ R.els[] .
'     elemGen     : .
'     actionStmts : __ e (____) ___ JS *_* (__________).
'                   _________________ rv ________.
'     outVal      : rv ___ (_____). .
'     outStatus   : ____. .
'   ___:
'     True  = ___ JS .
'     False = stale _____. ___ outStatus
'
'   *__ = _JS ____________________________*
'     ___ disabled ______ Click _______________.
'     _______________ (Elem_Read) _______.
'
'   ________________ Elem_Read ___ (__________).
' ============================================================
Friend Function Elem_Exec(ByVal elemHandle As Long, _
                          ByVal elemGen As String, _
                          ByVal actionStmts As String, _
                          ByRef outVal As String, _
                          ByRef outStatus As String) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    outVal = ""
    outStatus = ""
    Elem_Exec = False

    ' --- ____: _______________, _____ stale ______ ---
    If Len(m_domGen) > 0 Then
        If elemGen <> m_domGen Then
            outStatus = "stale"
            Exit Function
        End If
    End If

    js = "(function(){" & JsRegistryPrologue() & _
         "if(R.gen!==" & JsQuote(elemGen) & "){return {s:'gen',g:R.gen};}" & _
         "var e=R.els[" & CStr(elemHandle) & "];" & _
         "if(!e){return {s:'nil',g:R.gen};}" & _
         "var rv='';" & actionStmts & _
         "return {s:'ok',g:R.gen,v:String(rv)};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        outStatus = m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    If InStr(1, res, """s"":""ok""") = 0 Then
        outStatus = "stale"
        Exit Function
    End If

    Wv2Json.JsonPickStr res, "v", outVal
    Elem_Exec = True
End Function

' ============================================================
' *D-3b ___: DOM ____*
'
'   __1 ___ (2026-08-22) __ExecuteScript _ Promise .
'   __, _____3 ______*VBA _______*_____.
'   JS _________ (postMessage ________) _____. .
'   _____, D-1 ____ EvalSync _________.
' ============================================================


' ============================================================
' ImplicitWaitSec (D-3, __2 _C _ opt-in __)
'
'   0 ________, QuerySelector / GetElementById _*______*
'   ___________ (SeleniumVBA _ ImplicitMaxWait _____).
'
'   *___ 0 (__)* _________________.
'   WaitFor / WaitGone __________*____________*.
'
'   ___:
'     p.ImplicitWaitSec = 5      ' ________ 5 .
'     p.ImplicitWaitSec = 0      ' .
' ============================================================
Public Property Get ImplicitWaitSec() As Single
    ImplicitWaitSec = m_implicitWaitSec
End Property

Public Property Let ImplicitWaitSec(ByVal waitSec As Single)
    If waitSec < 0 Then waitSec = 0
    m_implicitWaitSec = waitSec
    Wv2Log.LogD "Wv2Pane.ImplicitWaitSec = " & waitSec & " _"
End Property


' ============================================================
' WaitFor (D-3, __2 _C ____ API)
'   ____________*_____*__.
'
'   __:
'     selector   : CSS .
'     timeoutSec : __ (_). __ 10. 0 .
'   ___:
'     _____            -> Wv2Element
'     __________  -> *Nothing + LastEvalOk = True* (__)
'     __ (_______) -> *Nothing + LastEvalOk = False* (LastEvalError ___)
'
'   *_______ (__5 _A)*
'     display:none ____________. SPA __ DOM _______, .
'     _______. ____________ WaitVisible _______.
'
'   *______________*
'     in-callback (____14) ____________, ________.
'     ______, _________________ timeoutSec _____.
'
'   *___________/___________* (____20)
'
'   *_______________*
'     View ________ EvalSync _ no-view ____, .
'     *_______ Nothing ___*. __________, .
'     EvalSync ______________ (____ D4WaitPane ___).
'
'   ___:
'     Dim el As Wv2Element
'     Set el = p.WaitFor("#result", 5)
'     If el Is Nothing Then
'         Debug.Print IIf(p.LastEvalOk, "______", "__: " & p.LastEvalError)
'     End If
' ============================================================
Public Function WaitFor(ByVal selector As String, _
                        Optional ByVal timeoutSec As Single = WAIT_TIMEOUT_DEFAULT) As Wv2Element
    If timeoutSec <= 0 Then timeoutSec = WAIT_TIMEOUT_DEFAULT

    Set WaitFor = FindWithWait("document.querySelector(" & JsQuote(selector) & ")", _
                               "WaitFor", selector, timeoutSec)
End Function


' ============================================================
' WaitGone (D-3)
'   ____________*______*__.
'
'   ___:
'     True                  ... ___ (_______ 1 ______)
'     False + LastEvalOk=True  ... __________ (_________)
'     False + LastEvalOk=False ... __ (___ LastEvalError)
'
'   *_______________* (D-3 _____)
'     FindOne ____, ____________________ R.els .
'     push _____ (10 ____ 70 _____). WaitGone .
'     ____, found ____________ (ExistsNow) ___.
'
'   ___:
'     If Not p.WaitGone(".spinner", 10) Then Debug.Print "_______"
' ============================================================
Public Function WaitGone(ByVal selector As String, _
                         Optional ByVal timeoutSec As Single = WAIT_TIMEOUT_DEFAULT) As Boolean
    Dim t0 As Single
    Dim found As Boolean
    Dim ok As Boolean

    WaitGone = False
    If timeoutSec <= 0 Then timeoutSec = WAIT_TIMEOUT_DEFAULT

    t0 = Timer
    Do
        found = ExistsNow(selector, ok)

        If Not ok Then
            ' ___________. LastEvalOk = False _____.
            Debug.Print "Wv2Pane.WaitGone: __ (" & selector & ") err=" & m_lastEvalError
            Exit Function
        End If

        If Not found Then
            WaitGone = True
            Exit Function
        End If

        If ElapsedSec(t0) >= timeoutSec Then
            Debug.Print "Wv2Pane.WaitGone: ____ (" & selector & ") " & _
                        timeoutSec & " ___________"
            Exit Function
        End If

        PollPause
    Loop
End Function


' ============================================================
' FindWithWait (D-3, Private) - ______ FindOne .
'
'   waitSec <= 0 __ 1 _____ (= D-2 _________).
'   *______ (Nothing + LastEvalOk=True) ______*
'   __ (LastEvalOk=False) _______.
' ============================================================
Private Function FindWithWait(ByVal finderExpr As String, _
                              ByVal apiName As String, _
                              ByVal arg As String, _
                              ByVal waitSec As Single) As Wv2Element
    Dim el As Wv2Element
    Dim t0 As Single

    t0 = Timer
    Do
        Set el = FindOne(finderExpr, apiName, arg, (waitSec > 0))

        If Not el Is Nothing Then
            Set FindWithWait = el
            Exit Function
        End If

        If Not m_lastEvalOk Then Exit Function
        If waitSec <= 0 Then Exit Function
        If ElapsedSec(t0) >= waitSec Then
            Debug.Print "Wv2Pane." & apiName & ": ____ (" & arg & ") " & _
                        waitSec & " ___________"
            Exit Function
        End If

        PollPause
    Loop
End Function


' ============================================================
' ExistsNow (D-3, Private) - _________ 1 .
'
'   *___________* WaitGone _______ (________).
'   outOk : JS ___________. False ___________.
' ============================================================
Private Function ExistsNow(ByVal selector As String, ByRef outOk As Boolean) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    ExistsNow = False
    outOk = False

    js = "(function(){" & JsRegistryPrologue() & _
         "var e=document.querySelector(" & JsQuote(selector) & ");" & _
         "return {g:R.gen,found:!!e};})()"

    res = EvalSync(js)
    If Not m_lastEvalOk Then Exit Function

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    outOk = True
    ExistsNow = (InStr(1, res, """found"":true") > 0)
End Function


' ============================================================
' PollPause / ElapsedSec (D-3, Private) - .
'
'   PollPause : POLL_INTERVAL_SEC (100ms) __ DoEvents ___.
'               *Sleep _____* WebView2 __________,
'               _____________________.
'
'   ElapsedSec: *Timer ___ 0 __ 0 ___* __, .
'               _______________. .
'               (D-3 _____. _____ 1:53 _____________).
' ============================================================
Private Sub PollPause()
    Dim t0 As Single

    t0 = Timer
    Do
        DoEvents
        If ElapsedSec(t0) >= POLL_INTERVAL_SEC Then Exit Do
    Loop
End Sub

Private Function ElapsedSec(ByVal sinceTimer As Single) As Single
    Dim d As Single

    d = Timer - sinceTimer
    If d < 0 Then d = d + 86400
    ElapsedSec = d
End Function

' ============================================================
' *D-4a ___: ____ SPA ____*
' ============================================================


' ============================================================
' SpaProbeState (D-4a) - .
'
'   ___: __ JSON (___). _______ + LastEvalOk = False
'
'   JSON ___:
'     v  ___        h  ___ (True/False)   g  ____    rp .
'     f  ___ fetch  x  ___ XHR
'     q  ___ DOM _______ ms   qn ____________ ms
'     m  ___ DOM __  sm ____ DOM .
'     n  _____        sn .
'     up ___________ ms      ob MutationObserver .
'     sda DOM _____ arm ___     sdh .
'     sna _______ arm ___     snh .
'     sua URL _____ arm ___      suh .
'     lgN _______                 lgOn .
'     il  *________* (__ 5 _. "xhr 1234ms https://..." __)
'     ifn __________*_____*__ (__________)
'     ifo ___*________*__ (StaleInflightMs ______)
'
'   *____________, __________*
'   ________ Ensure, __ Read__ 2 _______.
'
'   ___:
'     Debug.Print p.SpaProbeState
'     -> {""v"":1,""h"":true,""g"":1,...}
' ============================================================
Public Function SpaProbeState() As String
    Dim js As String

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {v:P.v,h:HP(P),g:P.gen,rp:P.repairs," & _
         "f:P.fetchN,x:P.xhrN," & _
         "q:Math.round(N()-P.lastMut),qn:Math.round(N()-P.lastNet)," & _
         "m:P.mutN,sm:P.skipMut,n:P.netN,sn:P.skipNet," & _
         "up:Math.round(N()-P.since),ob:!!P.ob," & _
         "sda:!!P.sdArmed,sdh:P.sdHits,sna:!!P.snArmed,snh:P.snHits," & _
         "sua:!!P.suArmed," & _
         "suh:(P.suArmed&&String(location.href).indexOf(P.suPat)>=0)?1:0," & _
         "lgN:P.log.length,lgOn:!!P.logOn," & _
         "il:P.ifl.map(function(o){" & _
         "return o.w+' '+Math.round(N()-o.t)+'ms '+o.u;}).slice(0,5).join(' # ')," & _
         "ifn:P.ifl.filter(function(o){return (N()-o.t)<" & StaleMsJs() & ";}).length," & _
         "ifo:P.ifl.filter(function(o){return (N()-o.t)>=" & StaleMsJs() & ";}).length};})()"

    SpaProbeState = EvalSync(js)

    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.SpaProbeState: __ err=" & m_lastEvalError
        SpaProbeState = ""
    End If
End Function


' ============================================================
' SpaProbeHealthy (D-4a) - .
'   *___________ False ___* (___). _______ LastEvalOk.
' ============================================================
Public Function SpaProbeHealthy() As Boolean
    Dim st As String

    st = SpaProbeState()
    If Not m_lastEvalOk Then Exit Function

    SpaProbeHealthy = (InStr(1, st, """h"":true") > 0)
End Function


' ============================================================
' SpaProbeReset (D-4a) - .
'
'   *_____* _____________________.
'   _______________, ___________________.
'
'   *__ fetch / XHR ______* _______, .
'   ____________________. ______ window.fetch _____.
'
'   *_____________________* (window.__wv2p _ null .
'   ___). ___________________, ____________.
'   null _______ 1 _______, ________________.
' ============================================================
Public Function SpaProbeReset() As Boolean
    Dim js As String

    js = "(function(){var o=window.__wv2p;if(o){o.v=-1;}" & SpaProbeJs() & _
         "return {g:P.gen,h:HP(P)};})()"

    EvalSync js
    SpaProbeReset = m_lastEvalOk

    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.SpaProbeReset: __ err=" & m_lastEvalError
    Else
        Wv2Log.LogD "Wv2Pane.SpaProbeReset: __________"
    End If
End Function


' ============================================================
' SpaProbeJs (D-4a, Private) - ________ JS (____________)
'
'   _____, __ P ________________.
'   ______, *__________*. _____ HP(P) ____.
'
'   *JS _________________* (________).
'   *_____*
'     fetch          ... _____ +1, __/___ -1
'     XHR            ... send _ +1, loadend _ -1 (_____ __wv2d ___)
'     MutationObserver ... childList / characterData / subtree / attributes
'   _____ (ignoreNet / ignoreSel) _*_______*__,
'   _____________ (D-4b _______).
' ============================================================
Private Function SpaProbeJs() As String
    Dim js As String

    js = "var V=" & CStr(SPA_PROBE_VER) & ",N=function(){return performance.now();};"
    js = js & "var XP=(window.XMLHttpRequest?XMLHttpRequest.prototype:null);"
    js = js & "var HP=function(p){return !!p&&p.v===V" & _
                      "&&(typeof window.fetch!=='function'||window.fetch===p.fw)" & _
                      "&&(!XP||(XP.open===p.xo&&XP.send===p.xs))&&!!p.ob;};"
    js = js & "var P=window.__wv2p;"
    js = js & "if(!HP(P)){var prev=P;"

    ' --- _________ (__________) ---
    js = js & "var q={v:V,gen:(prev?((prev.gen||0)+1):1)," & _
                      "repairs:(prev?((prev.repairs||0)+1):0)," & _
                      "since:N(),lastMut:N(),lastNet:N()," & _
                      "fetchN:0,xhrN:0,mutN:0,netN:0,skipMut:0,skipNet:0," & _
                      "sdArmed:false,sdSel:'',sdHits:0,sdObs:null," & _
                      "snArmed:false,snPat:'',snHits:0," & _
                      "suArmed:false,suPat:''," & _
                      "ifl:[],log:[],logOn:(prev&&prev.logOn?true:false)," & _
                      "ignoreNet:(prev&&prev.ignoreNet?prev.ignoreNet:[])," & _
                      "ignoreSel:(prev&&prev.ignoreSel?prev.ignoreSel:'')};"

    ' --- _____ (__ q ________________) ---
    js = js & "var ig=function(u){var a=q.ignoreNet,i;" & _
                      "for(i=0;i<a.length;i++){if(u.indexOf(a[i])>=0){return true;}}return false;};"
    js = js & "var sk=function(t){if(!q.ignoreSel){return false;}" & _
                      "try{var e=(t&&t.nodeType===1)?t:(t?t.parentElement:null);" & _
                      "return !!(e&&e.closest(q.ignoreSel));}catch(x){return false;}};"

    ' --- D-4e: in-flight ________ ---
    '   ifl ______________. *_____________*
    '   log _______________ (_____. logOn ____).
    js = js & "var lg=function(k,d){if(!q.logOn){return;}" & _
                      "q.log.push(Math.round(N())+' '+k+' '+String(d).slice(0,120));" & _
                      "if(q.log.length>200){q.log.shift();}};q.lg=lg;"
    js = js & "var ifA=function(u,w){var o={u:String(u).slice(0,150),t:N(),w:w};" & _
                      "q.ifl.push(o);lg('NET+',w+' '+o.u);return o;};"
    js = js & "var ifD=function(o){var i=q.ifl.indexOf(o);" & _
                      "if(i>=0){q.ifl.splice(i,1);}" & _
                      "lg('NET-',Math.round(N()-o.t)+'ms '+o.u);};"

    ' --- D-4c: _____________ arm _ ---
    '   *_____ lastMut _____* = .
    '   (_____________. 1 ____________).
    js = js & "var sn=function(u){if(q.snArmed&&q.snPat&&" & _
                      "String(u).indexOf(q.snPat)>=0){q.snHits++;q.lastMut=N();" & _
                      "lg('SIG','net hit '+String(u).slice(0,80));}};q.sn=sn;"
    js = js & "q.armUrl=function(pat){q.suPat=pat;q.suArmed=true;return 0;};"
    js = js & "q.armDom=function(sel){var n;" & _
                      "try{n=document.querySelector(sel);}catch(e){return -2;}" & _
                      "if(!n){return -1;}" & _
                      "if(q.sdObs){try{q.sdObs.disconnect();}catch(e){}}" & _
                      "q.sdSel=sel;q.sdHits=0;q.sdArmed=true;" & _
                      "q.sdObs=new MutationObserver(function(ms){" & _
                      "q.sdHits+=ms.length;q.lastMut=N();" & _
                      "lg('SIG','dom hit x'+ms.length);});" & _
                      "q.sdObs.observe(n," & _
                      "{childList:true,characterData:true,subtree:true,attributes:true});" & _
                      "return 0;};"
    js = js & "q.armNet=function(pat){q.snPat=pat;q.snHits=0;q.snArmed=true;return 0;};"
    js = js & "q.disarm=function(){if(q.sdObs){try{q.sdObs.disconnect();}catch(e){}}" & _
                      "q.sdObs=null;q.sdArmed=false;q.snArmed=false;q.suArmed=false;" & _
                      "return 0;};"

    ' --- fetch ____ (*___ window.fetch ___*) ---
    js = js & "if(typeof window.fetch==='function'){"
    js = js & "var of=(prev&&prev.of&&prev.fw===window.fetch)?prev.of:window.fetch;q.of=of;"
    '   *____________________*
    '   ___________________, arm ___________.
    js = js & "q.fw=function(u,o){" & _
                      "var r=(typeof u==='string')?u:((u&&u.url)?u.url:'');" & _
                      "var sp=ig(r);if(sp){q.skipNet++;}else{q.netN++;q.fetchN++;}" & _
                      "var en=sp?null:ifA(r,'fetch');" & _
                      "var d=function(){if(!sp){q.fetchN--;q.lastNet=N();ifD(en);}sn(r);};" & _
                      "return of.apply(this,arguments).then(" & _
                      "function(v){d();return v;},function(e){d();throw e;});};"
    js = js & "window.fetch=q.fw;}"

    ' --- XHR ____ ---
    js = js & "if(XP){"
    js = js & "var oo=(prev&&prev.oo&&prev.xo===XP.open)?prev.oo:XP.open;"
    js = js & "var os=(prev&&prev.os&&prev.xs===XP.send)?prev.os:XP.send;q.oo=oo;q.os=os;"
    js = js & "q.xo=function(m,u){try{this.__wv2u=String(u);}catch(e){}" & _
                      "return oo.apply(this,arguments);};"
    js = js & "q.xs=function(){var s=this;try{var u=s.__wv2u||'';" & _
                      "var sp=ig(u);if(sp){q.skipNet++;}else{q.netN++;q.xhrN++;}" & _
                      "var en=sp?null:ifA(u,'xhr');" & _
                      "s.addEventListener('loadend',function(){" & _
                      "if(!s.__wv2d){s.__wv2d=true;" & _
                      "if(!sp){q.xhrN--;q.lastNet=N();ifD(en);}sn(u);}});}catch(e){}" & _
                      "return os.apply(this,arguments);};"
    js = js & "XP.open=q.xo;XP.send=q.xs;}"

    ' --- DOM _____ ---
    js = js & "try{q.ob=new MutationObserver(function(ms){" & _
                      "var i,c=0,t;for(i=0;i<ms.length;i++){t=ms[i].target;" & _
                      "if(ms[i].type==='characterData'){t=(t&&t.parentElement)?t.parentElement:t;}" & _
                      "if(sk(t)){q.skipMut++;}else{c++;}}" & _
                      "if(c>0){q.mutN+=c;q.lastMut=N();lg('DOM','x'+c);}});"
    js = js & "q.ob.observe(document," & _
                      "{childList:true,characterData:true,subtree:true,attributes:true});" & _
                      "}catch(e){q.ob=null;}"

    js = js & "window.__wv2p=q;P=q;}"

    SpaProbeJs = js
End Function

' ============================================================
' *D-4b ___: ____*
' ============================================================


' ============================================================
' WaitSettled (D-4b) - ___ DOM .
'
'   __:
'     timeoutSec : __ (_). __ 10. 0 .
'     stableMs   : ___ (ms). __ 500. ____ DOM .
'   ___:
'     True                     ... .
'     False + LastEvalOk=True  ... ___________ (__)
'     False + LastEvalOk=False ... __ (___ LastEvalError)
'
'   *** True ____________________ ***
'     _____. __________________________.
'     _______ D-4c ______________.
'
'   *___________/___________* (____20)
'     10 ____ 10 _____ DoEvents _____.
'
'   ___:
'     el.Click
'     If p.WaitSettled(10) Then
'         Set r = p.QuerySelector("#result")
'     Else
'         Debug.Print "________: " & p.LastSettleInfo
'     End If
' ============================================================
Public Function WaitSettled(Optional ByVal timeoutSec As Single = WAIT_TIMEOUT_DEFAULT, _
                            Optional ByVal stableMs As Long = SETTLE_STABLE_MS_DEFAULT) As Boolean
    Dim t0 As Single
    Dim st As String
    Dim inflight As Long
    Dim quietMs As Long
    Dim quietEff As Long
    Dim elapsedMs As Long
    Dim polls As Long
    Dim sigOk As Boolean
    Dim inflightList As String
    Dim stale As Long

    WaitSettled = False
    m_lastSettle = ""
    If timeoutSec <= 0 Then timeoutSec = WAIT_TIMEOUT_DEFAULT
    If stableMs < 0 Then stableMs = SETTLE_STABLE_MS_DEFAULT

    t0 = Timer
    Do
        polls = polls + 1
        st = SpaProbeState()

        If Not m_lastEvalOk Then
            ' ___________ (____104).
            m_lastSettle = "error " & m_lastEvalError
            Wv2Log.LogW "Wv2Pane.WaitSettled: __ err=" & m_lastEvalError
            m_sigDom = False
            m_sigNet = False
            m_sigUrl = False
            Exit Function
        End If

        ' *_____________* (D-4f)
        '   ifn = __________ StaleInflightMs ___________.
        '   _________________________________,
        '   *________________________________*
        '   ___ WaitSettled __________. .
        '   AddIgnoreNetwork __________.
        inflight = Wv2Json.JsonGetNum(st, "ifn")
        stale = Wv2Json.JsonGetNum(st, "ifo")
        quietMs = Wv2Json.JsonGetNum(st, "q")

        ' --- D-4c: arm ____________ ---
        '   *arm ____________ arm _____ = ________*
        '   ____ arm ______, _______ (fail-visible).
        If m_sigDom Or m_sigNet Or m_sigUrl Then
            If (m_sigDom And InStr(1, st, """sda"":true") = 0) Or _
               (m_sigNet And InStr(1, st, """sna"":true") = 0) Or _
               (m_sigUrl And InStr(1, st, """sua"":true") = 0) Then
                m_lastSettle = "signal-lost (_____ arm ____)"
                Wv2Log.LogW "Wv2Pane.WaitSettled: *signal-lost* " & _
                            "arm __________________. " & _
                            "_____________ (arm _____)"
                DisarmSignals
                Exit Function
            End If
        End If

        sigOk = True
        If m_sigDom Then
            If Wv2Json.JsonGetNum(st, "sdh") = 0 Then sigOk = False
        End If
        If m_sigNet Then
            If Wv2Json.JsonGetNum(st, "snh") = 0 Then sigOk = False
        End If
        If m_sigUrl Then
            If Wv2Json.JsonGetNum(st, "suh") = 0 Then sigOk = False
        End If

        ' --- *_________________* (D-4b _____) ---
        '   _____ q _____ DOM ___________, .
        '   ___________ 1 _________________. .
        '   *___________________________*
        '   (________ ms _____________).
        '   __________________, *____ stableMs ____*.
        '   _________ q _____, ______________.
        '   ___ stableMs __________________.
        '   ___ D-4c _______ (arm) ___.
        elapsedMs = CLng(ElapsedSec(t0) * 1000)
        quietEff = quietMs
        If quietEff > elapsedMs Then quietEff = elapsedMs

        If inflight = 0 And quietEff >= stableMs And sigOk Then
            m_lastSettle = "stable q=" & quietMs & "ms(eff=" & quietEff & ") slack=" & _
                           (quietEff - stableMs) & "ms polls=" & polls & _
                           " elapsed=" & Format$(ElapsedSec(t0), "0.00") & "s" & _
                           IIf(stale > 0, " (___ " & stale & " ____)", "") & _
                           SigInfoOf(st)
            WaitSettled = True
            DisarmSignals
            Exit Function
        End If

        If ElapsedSec(t0) >= timeoutSec Then
            ' *____________, ____________* (D-4d ___)
            '   Maps _______ XHR _ 1 _____________.
            '   URL _____ AddIgnoreNetwork ____.
            Wv2Json.JsonPickStr st, "il", inflightList

            m_lastSettle = "timeout inflight=" & inflight & _
                           IIf(stale > 0, "(___ " & stale & " ____)", "") & _
                           " q=" & quietMs & _
                           "ms (stableMs=" & stableMs & ") polls=" & polls & _
                           SigInfoOf(st)
            If Len(inflightList) > 0 Then
                m_lastSettle = m_lastSettle & " | ___: " & inflightList
            End If
            Wv2Log.LogD "Wv2Pane.WaitSettled: ____ " & m_lastSettle
            DisarmSignals
            Exit Function
        End If

        PollPause
    Loop
End Function


' ============================================================
' LastSettleInfo (D-4b) - .
'
'   ___: stable q=<__ ms> slack=<__ ms> polls=<____> elapsed=<_>
'   ___: timeout inflight=<______> q=<__ ms> ... / error <__>
'
'   *slack ____ (100ms __) ________* __________.
'   _______________________, stableMs .
'   ______ (D-4c) ____________.
' ============================================================
Public Property Get LastSettleInfo() As String
    LastSettleInfo = m_lastSettle
End Property


' ============================================================
' AutoWaitAfterAction (D-4b, __7 _C _ opt-in __)
'
'   True ____, Wv2Element _ *.Value = / .Click / .SetAttribute ______*
'   ___ WaitSettled ___. __ False.
'
'   *D-3 ________* ________JS _____________,
'   ______ p.LastSettleInfo / p.WaitSettled ______.
'   ImplicitWaitSec (D-3) __________, ___ opt-in____.
' ============================================================
Public Property Get AutoWaitAfterAction() As Boolean
    AutoWaitAfterAction = m_autoWait
End Property

Public Property Let AutoWaitAfterAction(ByVal onOff As Boolean)
    m_autoWait = onOff
    Wv2Log.LogD "Wv2Pane.AutoWaitAfterAction = " & onOff
End Property


' ============================================================
' AddIgnoreNetwork / ClearIgnoreNetwork / IgnoreSelectors (D-4b, __5)
'
'   ________________. *______* ________.
'
'   AddIgnoreNetwork : URL _*____*____. .
'                      _: p.AddIgnoreNetwork "google-analytics"
'   IgnoreSelectors  : DOM __________ CSS ____ (______)
'                      _: p.IgnoreSelectors = ".clock, .ticker, #ads"
'                      *closest _____* __, .
'
'   *____ VBA ____* ________________.
'   ____________________________.
'
'   *** __________________ ***
'     ____________, __________________.
'     _________, ____________________.
' ============================================================
Public Function AddIgnoreNetwork(ByVal urlPattern As String) As Boolean
    Dim v As Variant

    If Len(urlPattern) = 0 Then Exit Function
    If m_ignoreNet Is Nothing Then Set m_ignoreNet = New Collection

    For Each v In m_ignoreNet
        If CStr(v) = urlPattern Then
            AddIgnoreNetwork = True
            Exit Function
        End If
    Next v

    m_ignoreNet.Add urlPattern
    Wv2Log.LogD "Wv2Pane.AddIgnoreNetwork: " & urlPattern & _
                " (_ " & m_ignoreNet.Count & " _)"
    AddIgnoreNetwork = True
End Function

Public Sub ClearIgnoreNetwork()
    Set m_ignoreNet = New Collection
    Wv2Log.LogD "Wv2Pane.ClearIgnoreNetwork: __________"
End Sub

Public Property Get IgnoreSelectors() As String
    IgnoreSelectors = m_ignoreSel
End Property

Public Property Let IgnoreSelectors(ByVal cssSelectors As String)
    m_ignoreSel = cssSelectors
    Wv2Log.LogD "Wv2Pane.IgnoreSelectors = " & cssSelectors
End Property


' ============================================================
' SpaProbeApplyJs (D-4b, Private) - ______________ JS
'
'   ___________________. __ P (____) .
'   _____. *__________* (_________________).
' ============================================================
Private Function SpaProbeApplyJs() As String
    Dim js As String
    Dim arr As String
    Dim v As Variant

    If Not m_ignoreNet Is Nothing Then
        For Each v In m_ignoreNet
            If Len(arr) > 0 Then arr = arr & ","
            arr = arr & JsQuote(CStr(v))
        Next v
    End If

    ' *_______* (D-4b _________)
    '   ______________, *_______________*.
    '   IgnoreSelectors = "" _ ClearIgnoreNetwork _______,
    '   ________________. ________ VBA __________.
    js = "P.ignoreNet=[" & arr & "];P.ignoreSel=" & _
         JsQuote(m_ignoreSel) & ";"

    SpaProbeApplyJs = js
End Function

' ============================================================
' *D-4c: ______ (arm ________)*
'
'   *__________________________________*
'   ____________. _______, ______*______ arm*_,
'   __ WaitSettled _____________.
'
'   _ .
'     arm _______*__ WaitSettled 1 _______*. _____ 2 .
'     ____ 2 _ arm ___. ________________.
'
'   _ .
'     ________ 1 ___________, *.
'     ______* (_____________).
'
'   _ *arm ____________ (fail-visible)*
'     ________________ arm ____. .
'     _____________, __________________.
'     VBA ___ arm ________, ______ signal-lost ______.
'
'   _ .
'     *____________________* .
'     _________, ________________.
'
'   ___ (arm ____ act __):
'     p.ArmContentSignal "#result-table"     ' .
'     el.Click
'     If Not p.WaitSettled(10) Then Debug.Print p.LastSettleInfo
' ============================================================


' ============================================================
' ArmContentSignal (D-4c) - .
'
'   *arm _______________ False* (fail-fast)
'   ___________, ___________________,
'   __________________________.
'   *_________________* ____________.
'
'   ___: True = arm ___ / False = .
'           ___ LastEvalError (not-found / bad-selector / ___)
' ============================================================
Public Function ArmContentSignal(ByVal selector As String) As Boolean
    Dim js As String
    Dim res As String

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {r:P.armDom(" & JsQuote(selector) & ")};})()"

    res = EvalSync(js)
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.ArmContentSignal: __ err=" & m_lastEvalError
        Exit Function
    End If

    ' *JsonGetNum ________ (____64)* __, ____________.
    If InStr(1, res, """r"":0") = 0 Then
        If InStr(1, res, "-2") > 0 Then
            m_lastEvalError = "bad-selector"
        Else
            m_lastEvalError = "not-found"
        End If
        m_lastEvalOk = False
        Wv2Log.LogW "Wv2Pane.ArmContentSignal: arm ____ (" & selector & _
                    ") __=" & m_lastEvalError
        Exit Function
    End If

    m_sigDom = True
    Wv2Log.LogD "Wv2Pane.ArmContentSignal: arm __ (" & selector & ")"
    ArmContentSignal = True
End Function


' ============================================================
' ArmNetworkSignal (D-4c) - .
'
'   urlPattern _ URL _*____*. _______________.
'   *_________* (___________). .
'   ______________. LastSettleInfo _ net:miss _____,
'   _____ D-4d _________ URL _______.
' ============================================================
Public Function ArmNetworkSignal(ByVal urlPattern As String) As Boolean
    Dim js As String

    If Len(urlPattern) = 0 Then Exit Function

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {r:P.armNet(" & JsQuote(urlPattern) & ")};})()"

    EvalSync js
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.ArmNetworkSignal: __ err=" & m_lastEvalError
        Exit Function
    End If

    m_sigNet = True
    Wv2Log.LogD "Wv2Pane.ArmNetworkSignal: arm __ (" & urlPattern & ")"
    ArmNetworkSignal = True
End Function


' ============================================================
' DisarmSignals (D-4c) - arm .
'   WaitSettled ______________ (______).
'   arm ___________________.
' ============================================================
Public Function DisarmSignals() As Boolean
    m_sigDom = False
    m_sigNet = False
    m_sigUrl = False

    EvalSync "(function(){var p=window.__wv2p;" & _
             "if(p&&p.disarm){p.disarm();}return 1;})()"

    DisarmSignals = m_lastEvalOk
End Function


' ============================================================
' SigInfoOf (D-4c, Private) - .
' ============================================================
Private Function SigInfoOf(ByVal st As String) As String
    Dim r As String

    If m_sigDom Then
        r = r & " dom:"
        If Wv2Json.JsonGetNum(st, "sdh") > 0 Then r = r & "hit" Else r = r & "miss"
    End If

    If m_sigNet Then
        r = r & " net:"
        If Wv2Json.JsonGetNum(st, "snh") > 0 Then r = r & "hit" Else r = r & "miss"
    End If

    If m_sigUrl Then
        r = r & " url:"
        If Wv2Json.JsonGetNum(st, "suh") > 0 Then r = r & "hit" Else r = r & "miss"
    End If

    If Len(r) > 0 Then r = " signal=" & Trim$(r)
    SigInfoOf = r
End Function

' ============================================================
' ArmUrlSignal (D-4e) - URL .
'
'   *D-4d ___________* SPA ____ DOM .
'   *URL ____* _____. Google Maps ____________,
'   ______ URL _ /place/... _ /@__,__,___ ____.
'   __________________________________.
'
'   pushState ______ DOM _____________,
'   *location.href _____________* ______.
'
'   ______. __ WaitSettled 1 _______.
'   *arm ______________________* __,
'   ____ URL ________________.
'
'   ___:
'     p.ArmUrlSignal "/place/"
'     el.Click
'     If p.WaitSettled(15) Then Debug.Print p.View_GetSource()
' ============================================================
Public Function ArmUrlSignal(ByVal urlFragment As String) As Boolean
    Dim js As String

    If Len(urlFragment) = 0 Then Exit Function

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {r:P.armUrl(" & JsQuote(urlFragment) & ")};})()"

    EvalSync js
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.ArmUrlSignal: __ err=" & m_lastEvalError
        Exit Function
    End If

    m_sigUrl = True
    Wv2Log.LogD "Wv2Pane.ArmUrlSignal: arm __ (" & urlFragment & ")"
    ArmUrlSignal = True
End Function


' ============================================================
' SpaProbeLogging (D-4e) - _____ ON / OFF
'
'   *___ False* ______, _______ 200 _________:
'     NET+ <fetch|xhr> <URL>     .
'     NET- <__ms> <URL>        .
'     DOM  x<__>               DOM .
'     SIG  <__> hit ...        arm .
'
'   *________________________*
'   _______ (WebDriver BiDi for SeleniumVBA) _ Discovery Log ____.
'   arm ___ 1 ___, _____ URL _ DOM ______, ____ arm ___ /
'   ___________.
'
'   *__ ON ____* _________________.
' ============================================================
Public Property Get SpaProbeLogging() As Boolean
    Dim st As String

    st = SpaProbeState()
    If Not m_lastEvalOk Then Exit Property

    SpaProbeLogging = (InStr(1, st, """lgOn"":true") > 0)
End Property

Public Property Let SpaProbeLogging(ByVal onOff As Boolean)
    Dim js As String

    js = "(function(){" & SpaProbeJs() & _
         "P.logOn=" & IIf(onOff, "true", "false") & ";" & _
         "if(!P.logOn){P.log=[];}return {r:1};})()"

    EvalSync js
    Wv2Log.LogD "Wv2Pane.SpaProbeLogging = " & onOff & _
                IIf(m_lastEvalOk, "", " (__: " & m_lastEvalError & ")")
End Property


' ============================================================
' SpaProbeDrainLog (D-4e) - ______________ Wv2Log .
'
'   *_______________* (______).
'   ___: _____. 0 _____________.
'
'   1 __ EvalSync ______________, *1 ____ 60 ___*.
'   _____________ (____ 60 __, _____________).
' ============================================================
Public Function SpaProbeDrainLog() As Long
    Dim js As String
    Dim res As String
    Dim txt As String
    Dim lines As Variant
    Dim i As Long

    js = "(function(){" & SpaProbeJs() & _
         "var a=P.log.splice(0,60);" & _
         "return {n:a.length,s:a.join(String.fromCharCode(1))};})()"

    res = EvalSync(js, 10)
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.SpaProbeDrainLog: __ err=" & m_lastEvalError
        Exit Function
    End If

    SpaProbeDrainLog = Wv2Json.JsonGetNum(res, "n")
    If SpaProbeDrainLog = 0 Then Exit Function

    Wv2Json.JsonPickStr res, "s", txt
    lines = Split(txt, ChrW$(1))

    Wv2Log.LogI "  --- ____ (" & SpaProbeDrainLog & " _) ---"
    For i = LBound(lines) To UBound(lines)
        Wv2Log.LogI "    " & lines(i)
    Next i
End Function

' ============================================================
' StaleInflightMs (D-4f) - ________________ (ms)
'
'   *__ 10000 (10 _)* 0 ______ (__________).
'
'   _ _____ (D-4d ___)
'     Google Maps ______ /search?tbm=map _*______*.
'     _____. _____ 2 ______ DOM ____ URL _________,
'     __ 1 ________ inflight=0 ____, _____________.
'     ____________________________, .
'     (AddIgnoreNetwork) ______, ________________.
'
'   _ ________ (___)
'     *___________, _______________*
'     _______ DOM ________________,
'     ________________ WaitSettled .
'     (True ____________________, ________).
'     __________ AddIgnoreNetwork _______,
'     D-4c ____________________.
' ============================================================
Public Property Get StaleInflightMs() As Long
    StaleInflightMs = m_staleMs
End Property

Public Property Let StaleInflightMs(ByVal ms As Long)
    If ms < 0 Then ms = 0
    m_staleMs = ms
    Wv2Log.LogD "Wv2Pane.StaleInflightMs = " & ms & _
                IIf(ms = 0, " (__: ______)", "")
End Property


' ============================================================
' StaleMsJs (D-4f, Private) - ______ JS .
'   0 (__) ________________________.
' ============================================================
Private Function StaleMsJs() As String
    If m_staleMs <= 0 Then
        StaleMsJs = "1e15"
    Else
        StaleMsJs = CStr(m_staleMs)
    End If
End Function

' ============================================================
' *D-6 ___: QuerySelectorAll _____*
' ============================================================


' ============================================================
' QuerySelectorAll (D-6) - .
'
'   __:
'     selector : CSS .
'     maxCount : ____ (__ 200). 0 .
'   ___:
'     Wv2Element _ Collection. *__________ Collection* (Nothing ____)
'     _________ Collection ___. ___ LastEvalOk ___ (____93).
'
'   *__________ LastAllTruncated = True*
'     ________ = __________________.
'
'   *______* querySelectorAll _______.
'
'   ___:
'     Dim els As Collection, el As Wv2Element
'     Set els = p.QuerySelectorAll("a[href]")
'     Debug.Print els.Count & " _"
'     For Each el In els
'         Debug.Print el.GetAttribute("href")
'     Next el
'
'   *1 ________ 1 ____* (D-2 _____) __, 200 .
'   _____ 200 __ = 8 ______. ___________ EvalSync .
'   JS __ 1 _______.
' ============================================================
Public Function QuerySelectorAll(ByVal selector As String, _
                                 Optional ByVal maxCount As Long = ALL_MAX_DEFAULT) As Collection
    Dim js As String
    Dim res As String
    Dim gen As String
    Dim handles As String
    Dim parts As Variant
    Dim i As Long
    Dim el As Wv2Element
    Dim outCol As Collection

    Set outCol = New Collection
    Set QuerySelectorAll = outCol
    m_allTruncated = False

    If maxCount <= 0 Then maxCount = ALL_MAX_DEFAULT

    ' *D-2 _ FindOne _ push ___________*
    '   __________________ (Wv2Json __________.
    '   __________ Split ______ _ __5).
    js = "(function(){" & JsRegistryPrologue() & _
         "var q=document.querySelectorAll(" & JsQuote(selector) & ");" & _
         "var mx=" & CStr(maxCount) & ",h=[],i;" & _
         "for(i=0;i<q.length&&i<mx;i++){R.els.push(q[i]);h.push(R.els.length-1);}" & _
         "return {g:R.gen,n:q.length,k:h.length,cut:(q.length>mx)," & _
         "h:h.join(','),tot:R.els.length};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        Debug.Print "Wv2Pane.QuerySelectorAll: __ (" & selector & ") err=" & m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    m_allTruncated = (InStr(1, res, """cut"":true") > 0)
    If m_allTruncated Then
        Wv2Log.LogW "Wv2Pane.QuerySelectorAll: *_____* (" & selector & _
                    ") _____ " & Wv2Json.JsonGetNum(res, "n") & _
                    " ____ " & maxCount & " _____"
    End If

    Wv2Json.JsonPickStr res, "h", handles
    If Len(handles) = 0 Then
        Debug.Print "Wv2Pane.QuerySelectorAll: 0 _ (" & selector & ")"
        WarnIfElementsPiled Wv2Json.JsonGetNum(res, "tot")
        Exit Function
    End If

    parts = Split(handles, ",")
    For i = LBound(parts) To UBound(parts)
        Set el = New Wv2Element
        el.InitElement Me, CLng(parts(i)), gen
        outCol.Add el
    Next i

    WarnIfElementsPiled Wv2Json.JsonGetNum(res, "tot")

    Debug.Print "Wv2Pane.QuerySelectorAll: " & outCol.Count & " _ (" & _
                selector & ") gen=" & gen
End Function


' ============================================================
' LastAllTruncated (D-6) - ___ QuerySelectorAll .
'   *True __________ = ____________*
'   ______ maxCount _____, _________.
' ============================================================
Public Property Get LastAllTruncated() As Boolean
    LastAllTruncated = m_allTruncated
End Property


' ============================================================
' ElementCount (D-6) - JS .
'
'   *___________* (D-2 __3). .
'   _______, QuerySelectorAll _ 1 __ 200 ____________.
'   _____________, ClearElementRegistry _____.
'   *________ Wv2Element ____ stale ___*
'
'   ___: _______. _____ -1 (LastEvalError ___)
' ============================================================
Public Property Get ElementCount() As Long
    Dim res As String

    res = EvalSync("(function(){" & JsRegistryPrologue() & _
                  "return {n:R.els.length};})()")

    If Not m_lastEvalOk Then
        ElementCount = -1
        Exit Property
    End If

    ElementCount = Wv2Json.JsonGetNum(res, "n")
End Property


' ============================================================
' WarnIfElementsPiled (D-6, Private) - .
'   *______* __________ Wv2Element ____ stale ___.
'   _________, _________________.
' ============================================================
Private Sub WarnIfElementsPiled(ByVal total As Long)
    If total < ELEMENT_WARN_COUNT Then Exit Sub

    Wv2Log.LogW "Wv2Pane: *________ " & total & " _______* " & _
                "(_____________). ______ ClearElementRegistry _" & _
                "_____. ______ Wv2Element ____ stale ___. "
End Sub



' ============================================================
' *N-1 : ______________ (WebResourceRequested)*
'
'   ______ N-1 ________. .
'     (1) _____   NetCaptureStart / NetCaptureStop / NetLogDrain .
'     (2) COM ___   View_AddWebResourceRequested __ 4 .
'     (3) ___       View_OnWebResourceRequested (+ in-callback ___)
'     (4) ___       NetRingAlloc / NetPush / NetElapsedMs / NetRingIndex
'     (5) _____   NetCtxName / NetCtxFromName
'
'   ___ (__):
'     p.NetCaptureStart          ' *Navigate _________ DOCUMENT _____*
'     ... __ 1 _____ ...
'     p.NetLogDrain              ' .
'     p.NetCaptureStop
' ============================================================


' ============================================================
' NetCaptureStart (N-1) - .
'
'   contexts  : __________________. ___ DOCUMENT,XHR,FETCH.
'               __________. "ALL" _______ (___ CSS ___).
'                 DOCUMENT / STYLESHEET / IMAGE / MEDIA / FONT / SCRIPT /
'                 XHR (= XMLHTTPREQUEST) / FETCH / TEXTTRACK / EVENTSOURCE /
'                 WEBSOCKET / MANIFEST / SIGNEDEXCHANGE / PING /
'                 CSPVIOLATIONREPORT / OTHER / ALL
'               *______* SDK _____________________.
'   uriFilter : URL _____. __ "*" (___). * _ ? ____.
'
'   *__________________*
'     ___________________, ________ (____106 ___).
'
'   *_____ 1 __________________ False ___*
'     _____________ 1 _______, .
'     _________________.
'
'   ___: True _____.
' ============================================================
Public Function NetCaptureStart(Optional ByVal contexts As String = "", _
                                Optional ByVal uriFilter As String = "*") As Boolean
    Dim parts As Variant
    Dim i     As Long
    Dim ctx   As Long
    Dim added As Long
    Dim hr    As Long
    Dim spec  As String

    If m_closed Then
        Wv2Log.LogW "Wv2Pane.NetCaptureStart: _____________"
        Exit Function
    End If
    If m_pView = 0 Then
        Wv2Log.LogW "Wv2Pane.NetCaptureStart: m_pView ____ (_________)"
        Exit Function
    End If

    If m_netOn Then NetCaptureStop

    spec = Trim$(contexts)
    If LenB(spec) = 0 Then spec = NET_CTX_DEFAULT
    If LenB(Trim$(uriFilter)) = 0 Then uriFilter = "*"

    ' --- (1) _________ ---
    hr = View_AddWebResourceRequested()
    If hr <> 0 Then
        Wv2Log.LogE "Wv2Pane.NetCaptureStart: add_WebResourceRequested __ hr=&H" & Hex(hr)
        Exit Function
    End If

    ' --- (2) _______ (*1 _______ 1 _____*) ---
    parts = Split(spec, ",")
    For i = LBound(parts) To UBound(parts)
        ctx = NetCtxFromName(CStr(parts(i)))
        If ctx < 0 Then
            Wv2Log.LogW "Wv2Pane.NetCaptureStart: __ [" & Trim$(CStr(parts(i))) & _
                        "] ___________"
        Else
            hr = View_AddWebResourceRequestedFilter(uriFilter, ctx)
            If hr <> 0 Then
                Wv2Log.LogW "Wv2Pane.NetCaptureStart: _________ " & _
                            NetCtxName(ctx) & " hr=&H" & Hex(hr)
            Else
                m_netFilters.Add uriFilter & vbTab & ctx
                added = added + 1
            End If
        End If
    Next i

    If added = 0 Then
        Wv2Log.LogE "Wv2Pane.NetCaptureStart: *_____ 1 ________* " & _
                    "___________ 1 _______, ____________. "
        View_RemoveWebResourceRequested
        Exit Function
    End If

    ' --- (3) _________ (N-3) ---
    '   *_______________* ______. .
    '   ____________ (__1 _A). _2 IF .
    '   *_______________* __, __________.
    m_netRespOn = False
    hr = View2_AddWebResourceResponseReceived()
    If hr = 0 Then
        m_netRespOn = True
    Else
        Wv2Log.LogW "Wv2Pane.NetCaptureStart: *_____________* hr=&H" & _
                    Hex(hr) & " __________________. "
    End If

    NetLogClear
    m_netOn = True
    NetCaptureStart = True

    Wv2Log.LogI "Wv2Pane.NetCaptureStart: ____ uri=" & uriFilter & _
                " __=" & spec & " (____ " & added & " _)"
End Function


' ============================================================
' NetCaptureStop (N-1) - .
'
'   *____________* (____106 ___)
'   ______________________, .
'   __ NetCaptureStart ___________________.
'
'   ___: True ____________.
' ============================================================
Public Function NetCaptureStop() As Boolean
    Dim v     As Variant
    Dim parts As Variant
    Dim hr    As Long
    Dim allOk As Boolean

    If Not m_netOn Then
        NetCaptureStop = True
        Exit Function
    End If

    allOk = True

    ' --- (1) _______ ---
    '   *_______ / View ______ COM _____*
    '   _________________. _______________.
    If m_closed Or m_pView = 0 Then
        Wv2Log.LogD "Wv2Pane.NetCaptureStop: View ______________"
    Else
        For Each v In m_netFilters
            parts = Split(CStr(v), vbTab)
            hr = View_RemoveWebResourceRequestedFilter(CStr(parts(0)), CLng(parts(1)))
            If hr <> 0 Then
                Wv2Log.LogW "Wv2Pane.NetCaptureStop: _________ " & _
                            CStr(parts(0)) & " / " & NetCtxName(CLng(parts(1))) & _
                            " hr=&H" & Hex(hr)
                allOk = False
            End If
        Next v
    End If
    Set m_netFilters = New Collection

    ' --- (2) _______ ---
    hr = View_RemoveWebResourceRequested()
    If hr <> 0 Then allOk = False

    ' --- (3) _________ (N-3) ---
    hr = View2_RemoveWebResourceResponseReceived()
    If hr <> 0 Then allOk = False
    m_netRespOn = False

    ' --- (4) N-4: ______ ResponseView ____ ---
    '   *________________* _______.
    NetRespPendingClear

    m_netOn = False
    NetCaptureStop = allOk

    Wv2Log.LogI "Wv2Pane.NetCaptureStop: ____ (___ " & m_netTotal & " _ / " & _
                "__ " & m_netCount & " _ / __ " & m_netDropped & " _)"
End Function


' ============================================================
' NetCaptureOn (N-1) - .
' ============================================================
Public Property Get NetCaptureOn() As Boolean
    NetCaptureOn = m_netOn
End Property


' ============================================================
' NetLogCount / NetLogTotal / NetLogDropped (N-1)
'
'   NetLogCount   : ____________ (______)
'   NetLogTotal   : _______*____* (________)
'   NetLogDropped : *_____________*
'
'   *________________, NetLogDrain __ 0 _____* (N-1b)
'   ___________. 0 ________ NetLogClear ___.
'
'   *NetLogCount _ NetLogTotal _________________*
'   ______________ (D-6 _ LastAllTruncated _____).
' ============================================================
Public Property Get NetLogCount() As Long
    NetLogCount = m_netCount
End Property

Public Property Get NetLogTotal() As Long
    NetLogTotal = m_netTotal
End Property

Public Property Get NetLogDropped() As Long
    NetLogDropped = m_netDropped
End Property


' ============================================================
' NetLogCapacity (N-1) - ______ (_)
'   *____________* (______). _____________.
' ============================================================
Public Property Get NetLogCapacity() As Long
    NetLogCapacity = m_netCap
End Property

Public Property Let NetLogCapacity(ByVal n As Long)
    If n < 1 Then n = 1
    If n > NET_CAP_MAX Then
        Wv2Log.LogW "Wv2Pane.NetLogCapacity: " & n & " ___ " & NET_CAP_MAX & " _________"
        n = NET_CAP_MAX
    End If
    m_netCap = n
    NetRingAlloc
    Wv2Log.LogD "Wv2Pane.NetLogCapacity = " & m_netCap & " (___________)"
End Property


' ============================================================
' NetLogClear (N-1) - .
'   *__ ms ___________* (__ 1 __ +0ms _____)
'   *______ (NetLogTotal / NetLogDropped) _ 0 ___*
'   ____________ NetLogDrain _____ (_______).
' ============================================================
Public Sub NetLogClear()
    NetRingAlloc
End Sub


' ============================================================
' NetLogLine (N-1) - ____ 1 .
'
'   idx _*____ 1 __*. ________.
'   1 __ vbTab ____ 7 __ (N-3 _ 5 -> 7):
'     ____ / __ms / ____ / ___ / URL / _____ / .
'   *____________* _____ (0_4) ________.
'   *______________* (__7 __)
'
'   *____________, ______________*
'   (N-5 _ PowerShell _________________)
' ============================================================
Public Property Get NetLogLine(ByVal idx As Long) As String
    Dim k As Long
    k = NetRingIndex(idx)
    If k < 0 Then Exit Property

    NetLogLine = m_netSeq(k) & vbTab & m_netMs(k) & vbTab & _
                 m_netMethod(k) & vbTab & NetCtxName(m_netCtx(k)) & vbTab & m_netUri(k) & _
                 vbTab & NetStatusText(m_netStatus(k)) & vbTab & m_netReason(k)
End Property


' ============================================================
' NetLogDrain (N-1) - ______ Wv2Log .
'
'   *______________* (D-4 _ SpaProbeDrainLog _____).
'   *N-3 ___* ______________, .
'   _______, ____________ NetRespUnmatched ______.
'   *____________* (____105). ________________.
'
'   ___: _____. 0 _____________.
' ============================================================
Public Function NetLogDrain() As Long
    Dim i As Long
    Dim n As Long

    n = m_netCount
    If n = 0 Then
        Wv2Log.LogI "  --- ________: _______ (___ " & m_netTotal & " _) ---"
        Exit Function
    End If

    Wv2Log.LogI "  --- ________ " & n & " _ (___ " & m_netTotal & " _) ---"
    Wv2Log.LogI "        #   __ms  ____ __       __  URL"
    For i = 1 To n
        Wv2Log.LogI "      " & NetLogFormat(i)
    Next i

    If m_netDropped > 0 Then
        ' *________* __, ___________________.
        '   _______________, *______________* .
        '   ______. ______ NetLogClear _____.
        Wv2Log.LogW "  *______ " & m_netDropped & " ____________* " & _
                    "NetLogCapacity _____, ____ NetLogDrain ____. "
    End If

    ' --- N-3: ____________ ---
    '   *_______________* _________________,
    '   ___ CSS ___________. _________ NetLogDropped __.
    If m_netUnmatched > 0 Then
        Wv2Log.LogI "  (____________ " & m_netUnmatched & " _ _ " & _
                    "___ CSS ______________, _____)"
    End If

    NetLogDrain = n
    ' *________ (_____)* __3 _Y
    NetRingEmpty
End Function


' ============================================================
' View_AddWebResourceRequested (N-1, ICoreWebView2 vtable 55)
'
'   *index 55 _ tools/vtable.py _ WebView2.h ________* (____115).
'   ____________ 54 ______ (OpenDevToolsWindow ______).
'
'   ___ View2_AddDOMContentLoaded ____________. .
'   _m_pView2 ____ m_pView__vtable 64/65 ____ 55/56___.
'   pVTableObj = h.Slot + VTABLE_OBJ_OFFSET ________ (____45).
'
'   ___: HRESULT (S_OK = 0 ___)
' ============================================================
Public Function View_AddWebResourceRequested() As Long
    If m_closed Then
        View_AddWebResourceRequested = &H80004003   ' E_POINTER
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: m_pView ____"
        View_AddWebResourceRequested = &H80004003
        Exit Function
    End If

    ' .
    If Not (m_persistentHandlers(HK_WebResourceRequested) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: _____________"
        View_AddWebResourceRequested = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_WebResourceRequested, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: _____ Acquire ___"
        View_AddWebResourceRequested = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As Currency
    Dim hr As Long
    hr = dcf(m_pView, 55, "View.add_WebResourceRequested", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: __ hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddWebResourceRequested = hr
        Exit Function
    End If

    m_persistentTokens(HK_WebResourceRequested) = token
    Set m_persistentHandlers(HK_WebResourceRequested) = h
    Debug.Print "Wv2Pane.View_AddWebResourceRequested OK: token=" & token & _
                ", handler.slot=" & h.Slot

    View_AddWebResourceRequested = 0
End Function


' ============================================================
' View_RemoveWebResourceRequested (N-1, ICoreWebView2 vtable 56)
'   _____ no-op (S_OK ___).
' ============================================================
Public Function View_RemoveWebResourceRequested() As Long
    If m_persistentHandlers(HK_WebResourceRequested) Is Nothing Then
        View_RemoveWebResourceRequested = 0   ' ____ no-op
        Exit Function
    End If

    Dim hr As Long
    If m_pView <> 0 Then
        hr = dcf(m_pView, 56, "View.remove_WebResourceRequested", _
                 m_persistentTokens(HK_WebResourceRequested))
        If hr <> 0 Then
            Debug.Print "Wv2Pane.View_RemoveWebResourceRequested: ____ hr=&H" & Hex(hr)
        Else
            Debug.Print "Wv2Pane.View_RemoveWebResourceRequested OK"
        End If
    Else
        ' m_pView _____ remove ________. ______.
        Debug.Print "Wv2Pane.View_RemoveWebResourceRequested: m_pView _ 0 _________"
    End If

    Set m_persistentHandlers(HK_WebResourceRequested) = Nothing
    m_persistentTokens(HK_WebResourceRequested) = 0@

    View_RemoveWebResourceRequested = hr
End Function


' ============================================================
' View_AddWebResourceRequestedFilter (N-1, ICoreWebView2 vtable 57)
'
'   *_____ 1 _____________ 1 _____*
'   _____ 1 _____. NetCaptureStart ____ 3 ______.
'
'   uriFilter       : URL _____. * _ ? ____ ("*" ____)
'   resourceContext : COREWEBVIEW2_WEB_RESOURCE_CONTEXT __ (0=ALL, 1=DOCUMENT, ...)
'
'   ___: HRESULT (S_OK = 0 ___)
' ============================================================
Public Function View_AddWebResourceRequestedFilter( _
    ByVal uriFilter As String, _
    ByVal resourceContext As Long) As Long

    If m_pView = 0 Then
        View_AddWebResourceRequestedFilter = &H80004003
        Exit Function
    End If
    View_AddWebResourceRequestedFilter = dcf(m_pView, 57, _
                                             "View.AddWebResourceRequestedFilter", _
                                             StrPtr(uriFilter), resourceContext)
End Function


' ============================================================
' View_RemoveWebResourceRequestedFilter (N-1, ICoreWebView2 vtable 58)
'   *_________ (uri, context) ____________*
' ============================================================
Public Function View_RemoveWebResourceRequestedFilter( _
    ByVal uriFilter As String, _
    ByVal resourceContext As Long) As Long

    If m_pView = 0 Then
        View_RemoveWebResourceRequestedFilter = &H80004003
        Exit Function
    End If
    View_RemoveWebResourceRequestedFilter = dcf(m_pView, 58, _
                                                "View.RemoveWebResourceRequestedFilter", _
                                                StrPtr(uriFilter), resourceContext)
End Function


' ============================================================
' View_OnWebResourceRequested (N-1)
'   ComCallbackHandler __________________. Public Sub __.
'
'   *___________________* (____)
'     WebResourceRequested _________________, .
'     Err.Raise ___*_________________*. ___ 1 .
'     __________, 1 __________________.
'     ___________*____ 1 __(err)________*.
'     ________ NetLogDrain _________.
' ============================================================
Public Sub View_OnWebResourceRequested(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim errNum As Long
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnWebResourceRequested_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        On Error Resume Next
        NetPush "(err)", -1, "*_______ " & errNum & " " & errDsc & "*"
    End If
End Sub


' ============================================================
' View_OnWebResourceRequested_Core (N-1, Private)
'
'   *__________ (__5)*
'     ___ Wv2Log _ Debug.Print _____. EvalSync ___ (____14 / 20).
'     dcf / GetStringProperty ___ funcName _____, .
'     ______________ (_________________).
'
'   args _ vtable (tools/vtable.py ___):
'     3 = get_Request(ICoreWebView2WebResourceRequest**)
'     7 = get_ResourceContext(COREWEBVIEW2_WEB_RESOURCE_CONTEXT*)
'   request _ vtable:
'     3 = get_Uri(LPWSTR*)   5 = get_Method(LPWSTR*)
'
'   *get_Request _ AddRef _____* ____ Release __ (COM __).
'   1 ______________, _______________.
' ============================================================
Private Sub View_OnWebResourceRequested_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim pReq As LongPtr
    Dim hr   As Long
    Dim uri  As String
    Dim mth  As String
    Dim ctx  As Long

    If pArgs = 0 Then
        NetPush "(err)", -1, "*args _ 0*"
        Exit Sub
    End If

    ' __ (________ 0 = ALL ___. get_ResourceContext ___ ALL .
    ' ______, ___ ALL ____________________)
    ctx = GetLongProperty(pArgs, 7, "")

    hr = dcf(pArgs, 3, "", VarPtr(pReq))
    If hr <> 0 Or pReq = 0 Then
        NetPush "(err)", ctx, "*get_Request __ hr=&H" & Hex(hr) & "*"
        Exit Sub
    End If

    uri = GetStringProperty(pReq, 3, "")
    mth = GetStringProperty(pReq, 5, "")

    NetPush mth, ctx, uri

    ' --- N-2: __ (_______) ---
    '   *__ OFF* ___, ___ NetDetailOn = True _________.
    '   pReq _ Release ______ (args _ pReq _ Invoke _________).
    If m_detOn Then NetReadDetail pReq, mth, uri

    ComRelease pReq
End Sub


' ============================================================
' NetReadDetail (N-2, Private) - 1 .
'
'   *_________* (__1 _A). args _ Invoke ___________,
'   _____________. __________.
'
'   *________ GET ____ (__7 __)*
'     GET _________. ___________________.
'     URL ___________ NetDetailUriFilter.
' ============================================================
Private Sub NetReadDetail(ByVal pReq As LongPtr, _
                          ByVal mth As String, _
                          ByVal uri As String)
    Dim headers  As String
    Dim cType    As String
    Dim nBytes   As Long
    Dim trunc    As Boolean
    Dim isText   As Boolean
    Dim bodyText As String
    Dim hexHead  As String
    Dim buf()    As Byte

    ' URL ___ (_______)
    If LenB(m_detUriFilter) > 0 Then
        If InStr(1, uri, m_detUriFilter, vbTextCompare) = 0 Then Exit Sub
    End If

    headers = NetReadHeaders(pReq)
    cType = NetHeaderValue(headers, "content-type")

    ' *GET _________*
    If UCase$(mth) <> "GET" Then
        If NetReadBody(pReq, nBytes, trunc, buf) Then
            If nBytes > 0 Then
                isText = NetIsTextType(cType)
                If isText Then
                    bodyText = Utf8BytesToString(buf, nBytes)
                Else
                    hexHead = NetHexHead(buf, nBytes)
                End If
            End If
        End If
    End If

    NetDetailPush m_netTotal, mth, uri, headers, cType, _
                  nBytes, trunc, isText, bodyText, hexHead
End Sub


' ============================================================
' NetRingAlloc (N-1, Private) - .
'   _________, ____________.
' ============================================================
Private Sub NetRingAlloc()
    If m_netCap < 1 Then m_netCap = NET_CAP_DEFAULT

    ReDim m_netSeq(0 To m_netCap - 1)
    ReDim m_netMs(0 To m_netCap - 1)
    ReDim m_netCtx(0 To m_netCap - 1)
    ReDim m_netMethod(0 To m_netCap - 1)
    ReDim m_netUri(0 To m_netCap - 1)
    ReDim m_netStatus(0 To m_netCap - 1)
    ReDim m_netReason(0 To m_netCap - 1)

    m_netHead = 0
    m_netCount = 0
    m_netTotal = 0
    m_netDropped = 0
    m_netUnmatched = 0
    m_netT0 = GetTickCount()
End Sub


' ============================================================
' NetRingEmpty (N-1b, Private) - *____* .
'   ______ (m_netTotal / m_netDropped) ___ ms ________.
'   NetLogDrain ___________________.
' ============================================================
Private Sub NetRingEmpty()
    m_netHead = 0
    m_netCount = 0
End Sub


' ============================================================
' NetPush (N-1, Private) - 1 .
'   *_____________, ________* (_______)
' ============================================================
Private Sub NetPush(ByVal mth As String, ByVal ctx As Long, ByVal uri As String)
    m_netTotal = m_netTotal + 1

    m_netSeq(m_netHead) = m_netTotal
    m_netMs(m_netHead) = NetElapsedMs()
    m_netCtx(m_netHead) = ctx
    m_netMethod(m_netHead) = mth
    m_netUri(m_netHead) = uri
    ' *_______________* _______________ (N-3)
    m_netStatus(m_netHead) = 0
    m_netReason(m_netHead) = ""

    m_netHead = m_netHead + 1
    If m_netHead >= m_netCap Then m_netHead = 0

    If m_netCount < m_netCap Then
        m_netCount = m_netCount + 1
    Else
        m_netDropped = m_netDropped + 1
    End If
End Sub


' ============================================================
' NetElapsedMs (N-1, Private) - _________ ms
'   GetTickCount _ Long ____ 49.7 ______. __ Double .
'   ___, ______ 2^32 ______.
' ============================================================
Private Function NetElapsedMs() As Long
    Dim d As Double
    d = CDbl(GetTickCount()) - CDbl(m_netT0)
    If d < 0 Then d = d + 4294967296#
    If d > 2147483647# Then d = 2147483647#
    NetElapsedMs = CLng(d)
End Function


' ============================================================
' NetRingIndex (N-1, Private) - ___ 1 __ -> .
'   _____ -1.
' ============================================================
Private Function NetRingIndex(ByVal idx As Long) As Long
    Dim k As Long

    If idx < 1 Or idx > m_netCount Then
        NetRingIndex = -1
        Exit Function
    End If

    k = m_netHead - m_netCount + (idx - 1)
    Do While k < 0
        k = k + m_netCap
    Loop
    NetRingIndex = k Mod m_netCap
End Function


' ============================================================
' NetLogFormat (N-1, Private) - 1 _______ 1 .
' ============================================================
Private Function NetLogFormat(ByVal idx As Long) As String
    Dim k As Long
    k = NetRingIndex(idx)
    If k < 0 Then Exit Function

    NetLogFormat = Right$("     " & m_netSeq(k), 5) & "  " & _
                   Right$("       " & m_netMs(k), 7) & "  " & _
                   Left$(m_netMethod(k) & "        ", 8) & " " & _
                   Left$(NetCtxName(m_netCtx(k)) & "          ", 10) & " " & _
                   Right$("     " & NetStatusText(m_netStatus(k)), 5) & " " & _
                   m_netUri(k)
End Function


' ============================================================
' NetStatusText (N-3, Private) - .
'   *0 _____________________* (__7 __)
'   0 _______________________, ______ (____111).
' ============================================================
Private Function NetStatusText(ByVal st As Long) As String
    If st = 0 Then Exit Function
    If st < 0 Then NetStatusText = "(___)": Exit Function
    NetStatusText = CStr(st)
End Function


' ============================================================
' NetCtxName (N-1, Private) - COREWEBVIEW2_WEB_RESOURCE_CONTEXT .
'   *__ tools/vtable.py --enum COREWEBVIEW2_WEB_RESOURCE_CONTEXT ___*
' ============================================================
Private Function NetCtxName(ByVal ctx As Long) As String
    Select Case ctx
        Case 0:  NetCtxName = "ALL"
        Case 1:  NetCtxName = "DOCUMENT"
        Case 2:  NetCtxName = "STYLESHEET"
        Case 3:  NetCtxName = "IMAGE"
        Case 4:  NetCtxName = "MEDIA"
        Case 5:  NetCtxName = "FONT"
        Case 6:  NetCtxName = "SCRIPT"
        Case 7:  NetCtxName = "XHR"
        Case 8:  NetCtxName = "FETCH"
        Case 9:  NetCtxName = "TEXTTRACK"
        Case 10: NetCtxName = "EVENTSRC"
        Case 11: NetCtxName = "WEBSOCKET"
        Case 12: NetCtxName = "MANIFEST"
        Case 13: NetCtxName = "SIGNEDEX"
        Case 14: NetCtxName = "PING"
        Case 15: NetCtxName = "CSPREPORT"
        Case 16: NetCtxName = "OTHER"
        Case Else: NetCtxName = "CTX" & ctx
    End Select
End Function


' ============================================================
' NetCtxFromName (N-1, Private) - __ (_____) -> .
'   _______ -1. *______* SDK _______________.
' ============================================================
Private Function NetCtxFromName(ByVal ctxToken As String) As Long
    Dim t As String

    t = UCase$(Trim$(ctxToken))
    t = Replace$(t, "_", "")
    t = Replace$(t, " ", "")

    If LenB(t) = 0 Then
        NetCtxFromName = -1
        Exit Function
    End If

    If IsNumeric(t) Then
        NetCtxFromName = CLng(t)
        Exit Function
    End If

    Select Case t
        Case "ALL":                             NetCtxFromName = 0
        Case "DOCUMENT", "DOC":                 NetCtxFromName = 1
        Case "STYLESHEET", "CSS":               NetCtxFromName = 2
        Case "IMAGE", "IMG":                    NetCtxFromName = 3
        Case "MEDIA":                           NetCtxFromName = 4
        Case "FONT":                            NetCtxFromName = 5
        Case "SCRIPT", "JS":                    NetCtxFromName = 6
        Case "XHR", "XMLHTTPREQUEST":           NetCtxFromName = 7
        Case "FETCH":                           NetCtxFromName = 8
        Case "TEXTTRACK":                       NetCtxFromName = 9
        Case "EVENTSOURCE", "EVENTSRC":         NetCtxFromName = 10
        Case "WEBSOCKET", "WS":                 NetCtxFromName = 11
        Case "MANIFEST":                        NetCtxFromName = 12
        Case "SIGNEDEXCHANGE", "SIGNEDEX":      NetCtxFromName = 13
        Case "PING":                            NetCtxFromName = 14
        Case "CSPVIOLATIONREPORT", "CSPREPORT": NetCtxFromName = 15
        Case "OTHER":                           NetCtxFromName = 16
        Case Else:                              NetCtxFromName = -1
    End Select
End Function


' ============================================================
' *N-2 : _____________*
'
'   N-1 ___ (URL / ____ / __) ___, *1 ____________*___.
'
'   _ *_________ OFF (__8 __)*
'     N-1 _________________________. .
'     _________ (1 _____ COM ____ 60 ___) .
'     ______, _______. `NetDetailOn = True` _________.
'
'   _ *__________ (__2)*
'     `get_Content` ___ `IStream` ______. __________,
'     ____ WebView2 _____________*________*.
'     _________ `Seek(0, STREAM_SEEK_SET)` ______.
'     *_________________* __ ____ POST .
'     ______, _________________ (____112).
'
'   _ *_________ (__3 _M)*
'     `Cookie` / `Authorization` _________________.
'     *____________, ________* __, .
'     `NetRedact = False` _____. ______.
'
'   _ ____ Public .
'     NetDetailOn        As Boolean  *__ False* .
'     NetDetailCapacity  As Long     ________ (__ 50)
'     NetBodyMaxBytes    As Long     ______ (__ 65536)
'     NetRedact          As Boolean  *__ True* .
'     AddRedactHeader(name) / ClearRedactHeaders()
'     NetDetailUriFilter As String   _____ URL _____ (_____)
'     NetDetailCount     As Long
'     NetDetailDropped   As Long     *________* (_______)
'     NetDetailLine(i)   As String   seq/____/URL/____/___/.
'     NetDetailHeaders(i) As String  ___ (__________)
'     NetDetailBody(i)   As String   _____ (___________)
'     NetDetail(i)       As Boolean  1 __ Wv2Log .
'     NetDetailAll()     As Long     .
'     NetDetailClear()
'
'   _ vtable index (*tools/vtable.py ___*, ____115)
'       ICoreWebView2WebResourceRequest::get_Content  ->IStream**   7
'       ICoreWebView2WebResourceRequest::get_Headers                9
'       ICoreWebView2HttpRequestHeaders::GetIterator                8
'       ...HttpHeadersCollectionIterator::GetCurrentHeader          3
'       ...::get_HasCurrentHeader / ::MoveNext                    4 / 5
'       IStream::Read / Seek        (Windows SDK objidl.h __)   3 / 5
' ============================================================


' ============================================================
' NetDetailOn (N-2) - __ (_______) .
'   *__ False* ______ N-1 _____.
' ============================================================
Public Property Get NetDetailOn() As Boolean
    NetDetailOn = m_detOn
End Property

Public Property Let NetDetailOn(ByVal onOff As Boolean)
    m_detOn = onOff
    Wv2Log.LogD "Wv2Pane.NetDetailOn = " & onOff & _
                IIf(onOff, " (__________)", " (____)")
End Property


' ============================================================
' NetDetailCapacity (N-2) - ________ (_)
'   *_________ (__5 _S)* ___________, .
'   _________ 1 ________. _____________.
'   *____________*
' ============================================================
Public Property Get NetDetailCapacity() As Long
    NetDetailCapacity = m_detCap
End Property

Public Property Let NetDetailCapacity(ByVal n As Long)
    If n < 1 Then n = 1
    If n > NET_DET_CAP_MAX Then
        Wv2Log.LogW "Wv2Pane.NetDetailCapacity: " & n & " ___ " & _
                    NET_DET_CAP_MAX & " _________"
        n = NET_DET_CAP_MAX
    End If
    m_detCap = n
    NetDetailRingAlloc
    Wv2Log.LogD "Wv2Pane.NetDetailCapacity = " & m_detCap & " (___________)"
End Property


' ============================================================
' NetBodyMaxBytes (N-2) - ________ (___)
'   *_____________* __, _____________.
'   _________ NetDetailLine __________ (_______).
' ============================================================
Public Property Get NetBodyMaxBytes() As Long
    NetBodyMaxBytes = m_detBodyMax
End Property

Public Property Let NetBodyMaxBytes(ByVal n As Long)
    If n < 0 Then n = 0
    If n > NET_BODY_MAX_LIMIT Then n = NET_BODY_MAX_LIMIT
    m_detBodyMax = n
    Wv2Log.LogD "Wv2Pane.NetBodyMaxBytes = " & m_detBodyMax
End Property


' ============================================================
' NetRedact (N-2) - .
'   *__ True* Cookie / Authorization _____________.
'   *_______________* __, False ____________.
' ============================================================
Public Property Get NetRedact() As Boolean
    NetRedact = m_detRedact
End Property

Public Property Let NetRedact(ByVal onOff As Boolean)
    m_detRedact = onOff
    If Not onOff Then
        Wv2Log.LogW "Wv2Pane.NetRedact = False *Cookie ___________* " & _
                    "_________________. "
    Else
        Wv2Log.LogD "Wv2Pane.NetRedact = True (__________)"
    End If
End Property


' ============================================================
' AddRedactHeader / ClearRedactHeaders (N-2)
'   ________ / ___________. __________.
'   ______ NetRedactDefaults __.
' ============================================================
Public Sub AddRedactHeader(ByVal headerName As String)
    Dim t As String
    t = LCase$(Trim$(headerName))
    If LenB(t) = 0 Then Exit Sub
    If InStr(1, m_detRedactList, "|" & t & "|") > 0 Then Exit Sub
    m_detRedactList = m_detRedactList & t & "|"
    Wv2Log.LogD "Wv2Pane.AddRedactHeader: " & t
End Sub

Public Sub ClearRedactHeaders()
    m_detRedactList = "|"
    Wv2Log.LogW "Wv2Pane.ClearRedactHeaders: *______ 1 _______* " & _
                "Cookie ______. "
End Sub


' ============================================================
' NetDetailUriFilter (N-2) - _____ URL .
'   _____. *______________* __, .
'   ________.
' ============================================================
Public Property Get NetDetailUriFilter() As String
    NetDetailUriFilter = m_detUriFilter
End Property

Public Property Let NetDetailUriFilter(ByVal pattern As String)
    m_detUriFilter = pattern
    Wv2Log.LogD "Wv2Pane.NetDetailUriFilter = [" & pattern & "]" & _
                IIf(LenB(pattern) = 0, " (____)", "")
End Property


' ============================================================
' NetDetailCount (N-2) - .
' ============================================================
Public Property Get NetDetailCount() As Long
    NetDetailCount = m_detCount
End Property


' ============================================================
' NetDetailDropped (N-2) - .
'   *_______* ___ NetLogDropped _____.
'   NetDetailClear / NetDetailCapacity ____ 0 ___.
' ============================================================
Public Property Get NetDetailDropped() As Long
    NetDetailDropped = m_detDropped
End Property


' ============================================================
' NetDetailLine (N-2) - __ 1 ________ (____1 __)
'   vbTab ____ 6 __:
'     ____ / ____ / URL / ________ / ____ / Content-Type
'   *_______________* (NetDetailHeaders / NetDetailBody ___)
' ============================================================
Public Property Get NetDetailLine(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailLine = m_detSeq(k) & vbTab & m_detMethod(k) & vbTab & m_detUri(k) & _
                    vbTab & m_detBytes(k) & vbTab & CStr(m_detTrunc(k)) & _
                    vbTab & m_detCType(k)
End Property


' ============================================================
' NetDetailHeaders (N-2) - __ 1 .
'   1 _ 1 ____ vbLf ___. *NetRedact = True __________*
' ============================================================
Public Property Get NetDetailHeaders(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailHeaders = NetRedactApply(m_detHeaders(k))
End Property


' ============================================================
' NetDetailBody (N-2) - __ 1 .
'   *___________________* (__4 _P).
'   _________. _____ NetDetailLine _ 4 ________.
' ============================================================
Public Property Get NetDetailBody(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailBody = m_detBody(k)
End Property


' ============================================================
' NetDetail (N-2) - __ 1 __ Wv2Log .
'   *__ (NetLogDrain) ______ (__6 _V)* 1 ___________,
'   _____________ (____105 ___).
'   ___: True _____.
' ============================================================
Public Function NetDetail(ByVal idx As Long) As Boolean
    Dim k As Long
    Dim i As Long
    Dim lines As Variant

    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Function

    Wv2Log.LogI "  --- __ #" & m_detSeq(k) & " ---"
    Wv2Log.LogI "      " & m_detMethod(k) & " " & m_detUri(k)

    If LenB(m_detHeaders(k)) > 0 Then
        lines = Split(NetRedactApply(m_detHeaders(k)), vbLf)
        Wv2Log.LogI "      [___ " & (UBound(lines) - LBound(lines) + 1) & " _]"
        For i = LBound(lines) To UBound(lines)
            Wv2Log.LogI "        " & lines(i)
        Next i
    Else
        Wv2Log.LogI "      [___] ______"
    End If

    If m_detBytes(k) = 0 Then
        Wv2Log.LogI "      [___] __"
    ElseIf m_detIsText(k) Then
        Wv2Log.LogI "      [___] " & m_detBytes(k) & " ___ (____" & _
                    IIf(m_detTrunc(k), "_*______*", "") & ")"
        lines = Split(Replace$(m_detBody(k), vbCrLf, vbLf), vbLf)
        For i = LBound(lines) To UBound(lines)
            Wv2Log.LogI "        " & lines(i)
        Next i
    Else
        Wv2Log.LogI "      [___] " & m_detBytes(k) & " ___ (____" & _
                    IIf(m_detTrunc(k), "_*______*", "") & _
                    ") __: " & m_detHex(k)
    End If

    ' --- N-3 / N-4: __ ---
    If m_detStatus(k) = 0 Then
        Wv2Log.LogI "      [__] _______"
    Else
        Wv2Log.LogI "      [__] " & NetStatusText(m_detStatus(k)) & " " & m_detReason(k)
        If LenB(m_detRespHeaders(k)) > 0 Then
            lines = Split(NetRedactApply(m_detRespHeaders(k)), vbLf)
            Wv2Log.LogI "      [_____ " & (UBound(lines) - LBound(lines) + 1) & " _]"
            For i = LBound(lines) To UBound(lines)
                Wv2Log.LogI "        " & lines(i)
            Next i
        End If

        ' --- N-4: _______ ---
        '   *encoding ___ hex _____* (__2 _P')
        '   ________ hex _______ (gzip __ 1F 8B).
        If m_detRespBytes(k) > 0 Or LenB(m_detRespEnc(k)) > 0 Then
            Wv2Log.LogI "      [____] " & m_detRespBytes(k) & " ___" & _
                        IIf(m_detRespTrunc(k), " *______*", "") & _
                        IIf(LenB(m_detRespEnc(k)) > 0, " encoding=" & m_detRespEnc(k), "") & _
                        IIf(m_detRespIsText(k), " (____)", " (____)")
            Wv2Log.LogI "        __: " & m_detRespHex(k)
            If LenB(m_detRespBody(k)) > 0 Then
                lines = Split(Replace$(Left$(m_detRespBody(k), 2000), vbCrLf, vbLf), vbLf)
                For i = LBound(lines) To UBound(lines)
                    Wv2Log.LogI "        " & lines(i)
                Next i
            End If
        End If
    End If

    NetDetail = True
End Function


' ============================================================
' NetDetailAll (N-2) - ___________ Wv2Log .
'   *___________* __. _________.
'   *________* (NetLogDrain ___. ___________)
' ============================================================
Public Function NetDetailAll() As Long
    Dim i As Long

    If m_detCount = 0 Then
        Wv2Log.LogI "  --- __: _______ ---"
        Exit Function
    End If

    Wv2Log.LogI "  --- __ " & m_detCount & " _ ---"
    If m_detDropped > 0 Then
        Wv2Log.LogW "  *________________ " & m_detDropped & " ___* " & _
                    "NetDetailCapacity _____, NetDetailUriFilter _____. "
    End If
    For i = 1 To m_detCount
        NetDetail i
    Next i
    NetDetailAll = m_detCount
End Function


' ============================================================
' NetDetailClear (N-2) - .
' ============================================================
Public Sub NetDetailClear()
    NetDetailRingAlloc
End Sub


' ============================================================
' NetReadHeaders (N-2, Private) - _______ 1 .
'
'   ICoreWebView2HttpRequestHeaders::GetIterator (8) _________,
'   GetCurrentHeader (3) / get_HasCurrentHeader (4) / MoveNext (5) ___.
'
'   *_________* ____ (NET_HDR_MAX) ___.
'   *GetCurrentHeader _ LPWSTR _ 2 ___* __, ____ CoTaskMemFree __.
'
'   ___: "name: value" _ vbLf ______. _________.
' ============================================================
Private Function NetReadHeaders(ByVal pReq As LongPtr) As String
    Dim pHeaders As LongPtr
    Dim hr       As Long

    If pReq = 0 Then Exit Function

    hr = dcf(pReq, 9, "", VarPtr(pHeaders))          ' request.get_Headers
    If hr <> 0 Or pHeaders = 0 Then Exit Function

    ' *____ GetIterator _ 8* (____ 7. _____ _ ____115)
    NetReadHeaders = NetReadHeadersIter(pHeaders, 8)
    ComRelease pHeaders
End Function


' ============================================================
' NetReadHeadersIter (N-3, Private) - ___________ 1 .
'
'   _____________. *____ GetIterator _ index __* .
'   ________ (___ 8 / ___ 7).
'   ____________ ICoreWebView2HttpHeadersCollectionIterator.
'     3 = GetCurrentHeader(LPWSTR* name, LPWSTR* value)
'     4 = get_HasCurrentHeader(BOOL*)   5 = MoveNext(BOOL*)
'
'   *_________* ____ (NET_HDR_MAX) ___.
'   *GetCurrentHeader _ LPWSTR _ 2 ___* __, ____ CoTaskMemFree __.
' ============================================================
Private Function NetReadHeadersIter(ByVal pHeaders As LongPtr, _
                                    ByVal iterIndex As Long) As String
    Dim pIter    As LongPtr
    Dim hr       As Long
    Dim n        As Long
    Dim pName    As LongPtr
    Dim pValue   As LongPtr
    Dim moreFlag As Long
    Dim sb       As String

    ' *__________ (N-3b)*
    '   ________________. .
    '   ______, __ 1 _________ (____111).
    If pHeaders = 0 Then
        NetReadHeadersIter = "(__: ___________ 0)"
        Exit Function
    End If

    hr = dcf(pHeaders, iterIndex, "", VarPtr(pIter))   ' GetIterator
    If hr <> 0 Then
        NetReadHeadersIter = "(__: GetIterator ___ hr=&H" & Hex(hr) & _
                             " index=" & iterIndex & ")"
        Exit Function
    End If
    If pIter = 0 Then
        NetReadHeadersIter = "(__: GetIterator _ S_OK __ iter _ 0 index=" & _
                             iterIndex & ")"
        Exit Function
    End If

    Do While n < NET_HDR_MAX
        If Not GetBoolProperty(pIter, 4, "") Then Exit Do   ' get_HasCurrentHeader

        pName = 0
        pValue = 0
        hr = dcf(pIter, 3, "", VarPtr(pName), VarPtr(pValue))   ' GetCurrentHeader
        If hr = 0 Then
            If LenB(sb) > 0 Then sb = sb & vbLf
            sb = sb & PtrToString(pName) & ": " & PtrToString(pValue)
            n = n + 1
        End If
        If pName <> 0 Then CoTaskMemFree pName
        If pValue <> 0 Then CoTaskMemFree pValue
        If hr <> 0 Then Exit Do

        moreFlag = 0
        hr = dcf(pIter, 5, "", VarPtr(moreFlag))     ' MoveNext(BOOL*)
        If hr <> 0 Or moreFlag = 0 Then Exit Do
    Loop

    If n >= NET_HDR_MAX Then
        sb = sb & vbLf & "(*" & NET_HDR_MAX & " _______*)"
    End If
    If n = 0 Then
        ' *___________ 1 ____*
        sb = "(__: ____ 0 _. HasCurrentHeader _____ False" & _
             " index=" & iterIndex & ")"
    End If

    ComRelease pIter
    NetReadHeadersIter = sb
End Function


' ============================================================
' NetReadBody (N-2, Private) - .
'
'   ICoreWebView2WebResourceRequest::get_Content (7) ___ IStream .
'   IStream::Read (3) ___, *_______ Seek (5) ______*.
'
'   *___________ (__2)*
'     IStream ______. _________, ____ WebView2 .
'     ______________*________*.
'
'   __ (_____):
'     outBytes  : _______ (_____________)
'     outTrunc  : .
'     outBuf()  : _______ (0 __, outBytes ____)
'
'   ___: True ________________ (0 _____ True).
'           False _________(GET __) ________.
' ============================================================
Private Function NetReadBody(ByVal pReq As LongPtr, _
                             ByRef outBytes As Long, _
                             ByRef outTrunc As Boolean, _
                             ByRef outBuf() As Byte) As Boolean
    Dim pStream As LongPtr
    Dim hr      As Long
    Dim cap     As Long
    Dim got     As Long
    Dim chunk   As Long
    Dim cbRead  As Long

    outBytes = 0
    outTrunc = False

    If pReq = 0 Then Exit Function

    hr = dcf(pReq, 7, "", VarPtr(pStream))           ' get_Content
    If hr <> 0 Or pStream = 0 Then Exit Function     ' _____ (GET __)

    cap = m_detBodyMax
    If cap < 1 Then
        ' __ 0 = ____. _______________________.
        ComRelease pStream
        NetReadBody = True
        Exit Function
    End If

    ' *______ NetReadStream ______ (N-4)*
    '   _______________________. .
    '   *________ Seek(0) ________* ____.
    If Not NetReadStream(pStream, cap, got, outTrunc, outBuf) Then
        ComRelease pStream
        Exit Function
    End If

    ' ** __________ ** (__2 _X)
    '   Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition)
    '   dlibMove _ 8 _______ (Currency 0@), dwOrigin _ 4 ___, ___ NULL ____ (CLngPtr(0)).
    hr = dcf(pStream, 5, "", 0@, STREAM_SEEK_SET, CLngPtr(0))
    If hr <> 0 Then
        ' *______* ________________________.
        ' _______________, ___________.
        outTrunc = True
        outBytes = got
        ComRelease pStream
        NetReadBody = True
        Exit Function
    End If

    ComRelease pStream
    outBytes = got
    NetReadBody = True
End Function


' ============================================================
' NetReadStream (N-4, Private) - IStream .
'
'   *___________* _________________.
'     ______ (N-2) ... *___________* ________ Seek(0)
'     _____ (N-4)   ... .
'
'   *_______ 1 __________________________*
'   _________________________________.
'
'   ___: True _______ (0 _____ True).
' ============================================================
Private Function NetReadStream(ByVal pStream As LongPtr, _
                               ByVal cap As Long, _
                               ByRef outBytes As Long, _
                               ByRef outTrunc As Boolean, _
                               ByRef outBuf() As Byte) As Boolean
    Dim hr     As Long
    Dim got    As Long
    Dim chunk  As Long
    Dim cbRead As Long

    outBytes = 0
    outTrunc = False
    If pStream = 0 Or cap < 1 Then Exit Function

    ReDim outBuf(0 To cap - 1)

    Do While got < cap
        chunk = cap - got
        If chunk > NET_BODY_CHUNK Then chunk = NET_BODY_CHUNK
        cbRead = 0
        hr = dcf(pStream, 3, "", VarPtr(outBuf(got)), chunk, VarPtr(cbRead))  ' Read
        If hr <> 0 Then Exit Do
        If cbRead <= 0 Then Exit Do
        got = got + cbRead
    Loop

    ' *________* _ 1 ____________ (_________)
    If got >= cap Then
        Dim probe(0 To 0) As Byte
        cbRead = 0
        hr = dcf(pStream, 3, "", VarPtr(probe(0)), 1&, VarPtr(cbRead))
        If hr = 0 And cbRead > 0 Then outTrunc = True
    End If

    outBytes = got
    NetReadStream = True
End Function


' ============================================================
' NetDetailPush (N-2, Private) - ___ 1 .
'   *______________* (_______).
' ============================================================
Private Sub NetDetailPush(ByVal seqNo As Long, _
                          ByVal mth As String, _
                          ByVal uri As String, _
                          ByVal headers As String, _
                          ByVal cType As String, _
                          ByVal nBytes As Long, _
                          ByVal trunc As Boolean, _
                          ByVal isText As Boolean, _
                          ByVal bodyText As String, _
                          ByVal hexHead As String)
    m_detSeq(m_detHead) = seqNo
    m_detMethod(m_detHead) = mth
    m_detUri(m_detHead) = uri
    m_detHeaders(m_detHead) = headers
    m_detCType(m_detHead) = cType
    m_detBytes(m_detHead) = nBytes
    m_detTrunc(m_detHead) = trunc
    m_detIsText(m_detHead) = isText
    m_detBody(m_detHead) = bodyText
    m_detHex(m_detHead) = hexHead
    ' N-3: __________ (_______________)
    m_detStatus(m_detHead) = 0
    m_detReason(m_detHead) = ""
    m_detRespHeaders(m_detHead) = ""
    ' N-4: __________ (_______________)
    m_detRespBody(m_detHead) = ""
    m_detRespBytes(m_detHead) = 0
    m_detRespTrunc(m_detHead) = False
    m_detRespIsText(m_detHead) = False
    m_detRespEnc(m_detHead) = ""
    m_detRespHex(m_detHead) = ""
    m_detRespGot(m_detHead) = False

    m_detHead = m_detHead + 1
    If m_detHead >= m_detCap Then m_detHead = 0
    If m_detCount < m_detCap Then
        m_detCount = m_detCount + 1
    Else
        m_detDropped = m_detDropped + 1
    End If
End Sub


' ============================================================
' NetDetailRingAlloc / NetDetailRingIndex (N-2, Private)
' ============================================================
Private Sub NetDetailRingAlloc()
    If m_detCap < 1 Then m_detCap = NET_DET_CAP_DEFAULT

    ReDim m_detSeq(0 To m_detCap - 1)
    ReDim m_detMethod(0 To m_detCap - 1)
    ReDim m_detUri(0 To m_detCap - 1)
    ReDim m_detHeaders(0 To m_detCap - 1)
    ReDim m_detCType(0 To m_detCap - 1)
    ReDim m_detBytes(0 To m_detCap - 1)
    ReDim m_detTrunc(0 To m_detCap - 1)
    ReDim m_detIsText(0 To m_detCap - 1)
    ReDim m_detBody(0 To m_detCap - 1)
    ReDim m_detHex(0 To m_detCap - 1)
    ReDim m_detRespHeaders(0 To m_detCap - 1)
    ReDim m_detStatus(0 To m_detCap - 1)
    ReDim m_detReason(0 To m_detCap - 1)
    ReDim m_detRespBody(0 To m_detCap - 1)
    ReDim m_detRespBytes(0 To m_detCap - 1)
    ReDim m_detRespTrunc(0 To m_detCap - 1)
    ReDim m_detRespIsText(0 To m_detCap - 1)
    ReDim m_detRespEnc(0 To m_detCap - 1)
    ReDim m_detRespHex(0 To m_detCap - 1)
    ReDim m_detRespGot(0 To m_detCap - 1)

    m_detHead = 0
    m_detCount = 0
    m_detDropped = 0
End Sub

Private Function NetDetailRingIndex(ByVal idx As Long) As Long
    Dim k As Long

    If idx < 1 Or idx > m_detCount Then
        NetDetailRingIndex = -1
        Exit Function
    End If

    k = m_detHead - m_detCount + (idx - 1)
    Do While k < 0
        k = k + m_detCap
    Loop
    NetDetailRingIndex = k Mod m_detCap
End Function


' ============================================================
' NetRedactApply (N-2, Private) - .
'   *________* ________________,
'   NetRedact = False ____________ (__3 _M).
' ============================================================
Private Function NetRedactApply(ByVal headers As String) As String
    Dim lines As Variant
    Dim i     As Long
    Dim p     As Long
    Dim nm    As String
    Dim out   As String

    If LenB(headers) = 0 Then Exit Function
    If Not m_detRedact Then
        NetRedactApply = headers
        Exit Function
    End If

    lines = Split(headers, vbLf)
    For i = LBound(lines) To UBound(lines)
        p = InStr(1, CStr(lines(i)), ":")
        If p > 1 Then
            nm = LCase$(Trim$(Left$(CStr(lines(i)), p - 1)))
            If InStr(1, m_detRedactList, "|" & nm & "|") > 0 Then
                lines(i) = Left$(CStr(lines(i)), p) & " <___ " & _
                           (Len(CStr(lines(i))) - p - 1) & " __>"
            End If
        End If
        If LenB(out) > 0 Then out = out & vbLf
        out = out & CStr(lines(i))
    Next i

    NetRedactApply = out
End Function


' ============================================================
' NetIsTextType (N-2, Private) - Content-Type .
'   *_________ False (______)* ___.
'   Content-Type ______, .
'   application/x-www-form-urlencoded _____ False .
'   (_____________ = ____111).
' ============================================================
Private Function NetIsTextType(ByVal cType As String) As Boolean
    Dim t As String
    t = LCase$(cType)
    If LenB(t) = 0 Then Exit Function

    If InStr(1, t, "text/") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "json") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "xml") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "x-www-form-urlencoded") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "javascript") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "charset=") > 0 Then NetIsTextType = True: Exit Function
End Function


' ============================================================
' NetHexHead (N-2, Private) - ________ 16 .
' ============================================================
Private Function NetHexHead(ByRef b() As Byte, ByVal cb As Long) As String
    Dim i As Long
    Dim n As Long
    Dim s As String

    n = cb
    If n > NET_HEX_HEAD Then n = NET_HEX_HEAD
    For i = 0 To n - 1
        s = s & Right$("0" & Hex$(b(i)), 2) & " "
    Next i
    If cb > n Then s = s & "..."
    NetHexHead = RTrim$(s)
End Function


' ============================================================
' NetHeaderValue (N-2, Private) - ________ 1 .
'   __________. _______.
' ============================================================
Private Function NetHeaderValue(ByVal headers As String, ByVal headerName As String) As String
    Dim lines As Variant
    Dim i     As Long
    Dim p     As Long
    Dim nm    As String

    If LenB(headers) = 0 Then Exit Function

    lines = Split(headers, vbLf)
    For i = LBound(lines) To UBound(lines)
        p = InStr(1, CStr(lines(i)), ":")
        If p > 1 Then
            nm = LCase$(Trim$(Left$(CStr(lines(i)), p - 1)))
            If nm = LCase$(headerName) Then
                NetHeaderValue = Trim$(Mid$(CStr(lines(i)), p + 1))
                Exit Function
            End If
        End If
    Next i
End Function


' ============================================================
' NetRedactDefaults (N-2, Private) - .
'   *_____________* _______ AddRedactHeader.
' ============================================================
Private Function NetRedactDefaults() As String
    NetRedactDefaults = "|cookie|set-cookie|authorization|proxy-authorization|" & _
                        "x-api-key|x-auth-token|x-csrf-token|api-key|"
End Function

' ============================================================
' *N-3 : _______________*
'
'   N-1 ____*__*_, N-2 ____*_____*___.
'   ___ 1 ___________________________.
'
'   _ *_______________ (N-3 _____)*
'     ICoreWebView2_2 _ vtable _ AddWebResourceResponseReceivedFilter .
'     _____ (61/62 ___ NavigateWithWebResourceRequest). .
'     *___ CSS __________*. ___ args _____ Request .
'     Response ___, *__ (COREWEBVIEW2_WEB_RESOURCE_CONTEXT) _______*
'     ______________.
'
'     ___ *______________________ (__1 _A)*.
'     N-1 _____________________, _____.
'     _________ NetCaptureStart "ALL" .
'     ______________. *___________*
'
'   _ _____ (__2 _P / __6 _Y)
'     args __ NavigationId _________________,
'     *URI + _________________*___*_______*.
'     *__ URL __________________. .
'     ______* __ ___________. ______ (301 -> 200) .
'     _______ 2 _______ 1 _ 1 ___.
'
'   _ *_______________ (__7 __)*
'     0 ________________________, __ (.
'     N-5 ________) _______ (____111).
'
'   _ ____ Public .
'     NetLogStatus(i)        As Long    _________ (0 = _______)
'     NetDetailRespHeaders(i) As String _____ (__________)
'     NetRespOn              As Boolean .
'     NetRespUnmatched       As Long    *______________*
'
'   _ vtable index (*tools/vtable.py ___*, ____115)
'       ICoreWebView2_2::add_WebResourceResponseReceived            61
'       ICoreWebView2_2::remove_WebResourceResponseReceived         62
'       ...ResponseReceivedEventArgs::get_Request / get_Response   3 / 4
'       ICoreWebView2WebResourceResponseView::get_Headers            3
'       ...::get_StatusCode (int*) / get_ReasonPhrase              4 / 5
'       *ICoreWebView2HttpResponseHeaders::GetIterator = 7*
'
'   *_* ___ ICoreWebView2HttpRequestHeaders::GetIterator _ 8 __,
'   ___ ICoreWebView2HttpResponseHeaders::GetIterator _ *7*.
'   _____ put_ ____ AppendHeader ______________.
'   _____________ 1 ____ GetHeaders ___ (____115).
' ============================================================


' ============================================================
' NetLogStatus (N-3) - .
'   *0 _____________* ___ 0 ___________.
'   __ (NetLogLine / NetLogFormat) __ 0 ______ (__7 __).
' ============================================================
Public Property Get NetLogStatus(ByVal idx As Long) As Long
    Dim k As Long
    k = NetRingIndex(idx)
    If k < 0 Then Exit Property
    NetLogStatus = m_netStatus(k)
End Property


' ============================================================
' NetRespOn (N-3) - .
'   *False _________________*
'   ICoreWebView2_2 ______________.
' ============================================================
Public Property Get NetRespOn() As Boolean
    NetRespOn = m_netRespOn
End Property


' ============================================================
' NetRespUnmatched (N-3) - .
'
'   *________________*
'     _________________, ___ CSS ________.
'     ___ DOCUMENT / XHR / FETCH _________, .
'     __________. *_______________*
'
'   *_________ NetLogDropped __*
'     ______________________, _________.
'     ___________ NetLogDropped _____.
' ============================================================
Public Property Get NetRespUnmatched() As Long
    NetRespUnmatched = m_netUnmatched
End Property


' ============================================================
' NetDetailRespHeaders (N-3) - __ 1 .
'   1 _ 1 ____ vbLf ___. *NetRedact = True __________*
'   (Set-Cookie ____________)
' ============================================================
Public Property Get NetDetailRespHeaders(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailRespHeaders = NetRedactApply(m_detRespHeaders(k))
End Property


' ============================================================
' View2_AddWebResourceResponseReceived (N-3, ICoreWebView2_2 vtable 61)
'
'   *_2 IF ________ m_pView2 ___* (____45).
'   ___ View2_AddDOMContentLoaded ____________.
'
'   ___: HRESULT (S_OK = 0 ___)
' ============================================================
Public Function View2_AddWebResourceResponseReceived() As Long
    If m_closed Then
        View2_AddWebResourceResponseReceived = &H80004003   ' E_POINTER
        Exit Function
    End If
    If Not EnsureView2() Then
        Debug.Print "Wv2Pane.View2_AddWebResourceResponseReceived: EnsureView2 __"
        View2_AddWebResourceResponseReceived = &H80004003
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_WebResourceResponseReceived) Is Nothing) Then
        Debug.Print "Wv2Pane.View2_AddWebResourceResponseReceived: ______"
        View2_AddWebResourceResponseReceived = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_WebResourceResponseReceived, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View2_AddWebResourceResponseReceived: Acquire ___"
        View2_AddWebResourceResponseReceived = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As Currency
    Dim hr As Long
    hr = dcf(m_pView2, 61, "View2.add_WebResourceResponseReceived", _
             pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View2_AddWebResourceResponseReceived: __ hr=&H" & Hex(hr)
        Set h = Nothing
        View2_AddWebResourceResponseReceived = hr
        Exit Function
    End If

    m_persistentTokens(HK_WebResourceResponseReceived) = token
    Set m_persistentHandlers(HK_WebResourceResponseReceived) = h
    Debug.Print "Wv2Pane.View2_AddWebResourceResponseReceived OK: token=" & token

    View2_AddWebResourceResponseReceived = 0
End Function


' ============================================================
' View2_RemoveWebResourceResponseReceived (N-3, vtable 62, m_pView2)
'   _____ no-op (S_OK ___).
' ============================================================
Public Function View2_RemoveWebResourceResponseReceived() As Long
    If m_persistentHandlers(HK_WebResourceResponseReceived) Is Nothing Then
        View2_RemoveWebResourceResponseReceived = 0
        Exit Function
    End If

    Dim hr As Long
    If m_pView2 <> 0 Then
        hr = dcf(m_pView2, 62, "View2.remove_WebResourceResponseReceived", _
                 m_persistentTokens(HK_WebResourceResponseReceived))
        If hr <> 0 Then
            Debug.Print "Wv2Pane.View2_RemoveWebResourceResponseReceived: ____ hr=&H" & Hex(hr)
        End If
    Else
        Debug.Print "Wv2Pane.View2_RemoveWebResourceResponseReceived: m_pView2 _ 0 _________"
    End If

    Set m_persistentHandlers(HK_WebResourceResponseReceived) = Nothing
    m_persistentTokens(HK_WebResourceResponseReceived) = 0@

    View2_RemoveWebResourceResponseReceived = hr
End Function


' ============================================================
' View2_OnWebResourceResponseReceived (N-3)
'   ComCallbackHandler __________________. Public Sub __.
'   *N-1 __________________* (________________)
' ============================================================
Public Sub View2_OnWebResourceResponseReceived(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim errNum As Long
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View2_OnWebResourceResponseReceived_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        On Error Resume Next
        m_netUnmatched = m_netUnmatched + 1
    End If
End Sub


' ============================================================
' View2_OnWebResourceResponseReceived_Core (N-3, Private)
'
'   *__________* (N-1 / N-2 ___). ___ Debug.Print _____.
'   *_________* __ (GetContent) _______ N-4 __.
'   ________________________.
'
'   args _ vtable: 3 = get_Request, 4 = get_Response
'   ResponseView _ vtable: 3 = get_Headers, 4 = get_StatusCode, 5 = get_ReasonPhrase
' ============================================================
Private Sub View2_OnWebResourceResponseReceived_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim pReq  As LongPtr
    Dim pResp As LongPtr
    Dim hr    As Long
    Dim uri   As String
    Dim mth   As String
    Dim st    As Long
    Dim reason As String
    Dim k     As Long

    If pArgs = 0 Then Exit Sub

    ' --- _______________ Request ___ ---
    hr = dcf(pArgs, 3, "", VarPtr(pReq))             ' get_Request
    If hr <> 0 Or pReq = 0 Then Exit Sub
    uri = GetStringProperty(pReq, 3, "")             ' request.get_Uri
    mth = GetStringProperty(pReq, 5, "")             ' request.get_Method
    ComRelease pReq

    ' --- ___________________________ ---
    '   *______________ (__1 _A / __5 _W)*
    '   _________________, ___ CSS _________.
    '   ____________________ = __.
    k = NetMatchRequest(uri, mth)
    If k < 0 Then
        m_netUnmatched = m_netUnmatched + 1
        Exit Sub
    End If

    ' --- Response ___ ---
    hr = dcf(pArgs, 4, "", VarPtr(pResp))            ' get_Response
    If hr <> 0 Or pResp = 0 Then
        m_netStatus(k) = -1                          ' *____________*
        m_netReason(k) = "(________)"
        Exit Sub
    End If

    st = GetLongProperty(pResp, 4, "")               ' get_StatusCode
    reason = GetStringProperty(pResp, 5, "")         ' get_ReasonPhrase

    m_netStatus(k) = st
    m_netReason(k) = reason

    ' --- _________ ON _____ (__4 _U) ---
    If m_detOn Then NetAttachRespHeaders m_netSeq(k), pResp, st, reason

    ' --- N-4: _________ (*_________*) ---
    '   ResponseView _ args __________, _______ GetContent .
    '   __. ____________, _______________.
    If m_respBodyOn And m_detOn Then NetStartRespContent m_netSeq(k), pResp

    ComRelease pResp
End Sub


' ============================================================
' NetMatchRequest (N-3, Private) - .
'
'   *URI + _________________*___*______*
'   (__2 _P / __6 _Y). _____________________.
'
'   *_______________, ____________________*
'   __ URL ____ 2 _____________ (Req1 -> Req2 -> Resp -> Resp)
'   _, _______*___________*. _______ FIFO ___,
'   ___ 1 _________________ (____________).
'
'   *______* __ URL ______________. .
'   ______________. args _ NavigationId .
'   _________, _______________.
'
'   ___: _______. ________ -1.
' ============================================================
Private Function NetMatchRequest(ByVal uri As String, ByVal mth As String) As Long
    Dim i As Long
    Dim k As Long

    NetMatchRequest = -1
    If m_netCount = 0 Then Exit Function

    ' *___ (i = 1 -> m_netCount)* = .
    For i = 1 To m_netCount
        k = NetRingIndex(i)
        If k >= 0 Then
            If m_netStatus(k) = 0 Then                       ' .
                If m_netUri(k) = uri Then
                    If LenB(mth) = 0 Or m_netMethod(k) = mth Then
                        NetMatchRequest = k
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function


' ============================================================
' NetAttachRespHeaders (N-3, Private) - .
'
'   ____________________________.
'   *_____________* (___________________).
' ============================================================
Private Sub NetAttachRespHeaders(ByVal seqNo As Long, _
                                 ByVal pResp As LongPtr, _
                                 ByVal st As Long, _
                                 ByVal reason As String)
    Dim i As Long
    Dim k As Long
    Dim pHeaders As LongPtr
    Dim hr As Long

    For i = m_detCount To 1 Step -1
        k = NetDetailRingIndex(i)
        If k >= 0 Then
            If m_detSeq(k) = seqNo Then
                m_detStatus(k) = st
                m_detReason(k) = reason

                ' ResponseView.get_Headers (3) -> ICoreWebView2HttpResponseHeaders
                hr = dcf(pResp, 3, "", VarPtr(pHeaders))
                If hr <> 0 Then
                    m_detRespHeaders(k) = "(__: get_Headers ___ hr=&H" & Hex(hr) & ")"
                ElseIf pHeaders = 0 Then
                    m_detRespHeaders(k) = "(__: get_Headers _ S_OK __ 0 ____)"
                Else
                    ' *____ GetIterator _ 7* (____ 8. _____)
                    m_detRespHeaders(k) = NetReadHeadersIter(pHeaders, 7)
                    ComRelease pHeaders
                End If
                Exit Sub
            End If
        End If
    Next i
End Sub


' ============================================================
' *N-4 : _______*
'
'   N-1_N-3 ___________________.
'   *N-4 __________* __ `GetContent` ______________,
'   ______ `OnResponseContentCompleted` ___.
'
'   _ *N-2 ____ Seek(0) ________*
'     _______________, ___________.
'     ____118 (__________________) ______.
'
'   _ *_________ (__1 _A)*
'     _____ KB, _____ KB _____. .
'     NetDetailOn __________. `NetRespBodyOn = True` ______.
'     *NetDetailOn _ True __________* (____________).
'
'   _ *____________ (__2 _P')*
'     ____ content-encoding: gzip / br / zstd ___. GetContent .
'     _________ SDK ___________. .
'     _content-encoding .
'     _*__ 32 ____ hex _____* (gzip __ 1F 8B _______)
'     _________ UTF-8 .
'     _ 3 ______. 1 __________ (____117).
'
'   _ *ResponseView _________ (__5 __)*
'     _GetContent _____ ResponseView .
'     ___________. __________*.
'     ____ = _____* _, ____20 ____________.
'     ___ AddRef __ m_respPending ___, ____ Release __.
'
'   _ ____ Public .
'     NetRespBodyOn        As Boolean  *__ False*
'     NetRespBodyMaxBytes  As Long     _____ (__ 65536)
'     NetRespBodyOrphan    As Long     .
'     NetDetailRespBody(i) As String   __ (___________)
'     NetDetailRespLine(i) As String   ____ / ___ / __ / encoding / __hex
'
'   _ vtable index (*tools/vtable.py ___*)
'       ICoreWebView2WebResourceResponseView::GetContent(handler)   6
'       ______ Invoke(HRESULT, IStream*)                       3
'       IStream::Read                                                3
' ============================================================


' ============================================================
' NetRespBodyOn (N-4) - .
'   *__ False* ___________ (__ KB _).
'   *NetDetailOn _ True __________* __, _________.
' ============================================================
Public Property Get NetRespBodyOn() As Boolean
    NetRespBodyOn = m_respBodyOn
End Property

Public Property Let NetRespBodyOn(ByVal onOff As Boolean)
    m_respBodyOn = onOff
    If onOff And Not m_detOn Then
        Wv2Log.LogW "Wv2Pane.NetRespBodyOn = True __ *NetDetailOn _ False* ___ " & _
                    "_________, 1 ______. NetDetailOn = True ___. "
    Else
        Wv2Log.LogD "Wv2Pane.NetRespBodyOn = " & onOff
    End If
End Property


' ============================================================
' NetRespBodyMaxBytes (N-4) - __________ (___)
'   *___________ (__3 _S)* _______.
' ============================================================
Public Property Get NetRespBodyMaxBytes() As Long
    NetRespBodyMaxBytes = m_respBodyMax
End Property

Public Property Let NetRespBodyMaxBytes(ByVal n As Long)
    If n < 0 Then n = 0
    If n > NET_BODY_MAX_LIMIT Then n = NET_BODY_MAX_LIMIT
    m_respBodyMax = n
    Wv2Log.LogD "Wv2Pane.NetRespBodyMaxBytes = " & m_respBodyMax
End Property


' ============================================================
' NetRespBodyOrphan (N-4) - .
'   *____________________* .
'   (________ / NetDetailClear _______). _______.
' ============================================================
Public Property Get NetRespBodyOrphan() As Long
    NetRespBodyOrphan = m_respBodyOrphan
End Property


' ============================================================
' NetDetailRespBody (N-4) - __ 1 .
'   *___________________* (N-2 _________).
'   ____________________ NetDetailRespLine .
'   encoding ____ hex _____.
' ============================================================
Public Property Get NetDetailRespBody(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property
    NetDetailRespBody = m_detRespBody(k)
End Property


' ============================================================
' NetDetailRespLine (N-4) - .
'   vbTab ____ 6 __:
'     ____ / ____ / _____ / content-encoding / __ 32 ____ hex /
'     *____* (0 ___________________)
'   *encoding _ hex _____* __ N-4 ___ (__2 _P').
' ============================================================
Public Property Get NetDetailRespLine(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailRespLine = m_detRespBytes(k) & vbTab & CStr(m_detRespTrunc(k)) & _
                        vbTab & CStr(m_detRespIsText(k)) & vbTab & m_detRespEnc(k) & _
                        vbTab & m_detRespHex(k) & vbTab & CStr(m_detRespGot(k))
End Property


' ============================================================
' NetStartRespContent (N-4, Private) - .
'
'   *WebResourceResponseReceived ___________*
'   ResponseView _ args __________, ___ GetContent _______.
'   GetContent ________, _______________.
'
'   *ResponseView _ AddRef ______ (__5 __)*
'     _____________ SDK ________. .
'     ______________. ___ / ______ Release __.
' ============================================================
Private Sub NetStartRespContent(ByVal seqNo As Long, ByVal pResp As LongPtr)
    Dim h  As ComCallbackHandler
    Dim hr As Long
    Dim ky As String

    If pResp = 0 Then Exit Sub

    ' *____________* ______________,
    '   ___________________________.
    If NetFindDetailBySeq(seqNo) < 0 Then Exit Sub

    ky = "r" & seqNo
    ' __ seq ________ (1 __ 1 ___________)
    On Error Resume Next
    Dim dummy As Variant
    dummy = m_respPending(ky)
    If Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    Set h = AcquireHandlerFor(HK_ResponseGetContent, Me)
    If h Is Nothing Then Exit Sub
    h.callbackId = seqNo

    ' *____ ResponseView ____*
    ComAddRef pResp
    m_respPending.Add pResp, ky

    hr = dcf(pResp, 6, "", h.Slot + VTABLE_OBJ_OFFSET)   ' GetContent
    If hr <> 0 Then
        ' ______. ______.
        ComRelease pResp
        m_respPending.Remove ky
        Set h = Nothing
    End If
End Sub


' ============================================================
' OnResponseContentCompleted (N-4)
'   ComCallbackHandler __________________. Public Sub __.
'   1 _______________ (_9.6a __).
'
'   *IStream _ Invoke _________* __________.
'   *N-1 ____________________*
' ============================================================
Public Sub OnResponseContentCompleted(ByVal cbId As Long, _
                                      ByVal errorCode As Long, _
                                      ByVal pStream As LongPtr)
    Dim errNum As Long

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    OnResponseContentCompleted_Core cbId, errorCode, pStream

GuardExit:
    errNum = Err.Number
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        On Error Resume Next
        m_respBodyOrphan = m_respBodyOrphan + 1
    End If
End Sub


Private Sub OnResponseContentCompleted_Core(ByVal cbId As Long, _
                                            ByVal errorCode As Long, _
                                            ByVal pStream As LongPtr)
    Dim k       As Long
    Dim nBytes  As Long
    Dim trunc   As Boolean
    Dim buf()   As Byte
    Dim enc     As String
    Dim cType   As String
    Dim ky      As String

    ' --- ______ ResponseView ____ (__5 __ ____) ---
    ky = "r" & cbId
    On Error Resume Next
    Dim pKept As LongPtr
    pKept = CLngPtr(m_respPending(ky))
    If pKept <> 0 Then
        ComRelease pKept
        m_respPending.Remove ky
    End If
    Err.Clear
    On Error GoTo 0

    ' --- ______ ---
    k = NetFindDetailBySeq(cbId)
    If k < 0 Then
        ' *_________________* (___ / Clear __)
        m_respBodyOrphan = m_respBodyOrphan + 1
        Exit Sub
    End If

    ' *____________* 0 _________, .
    m_detRespGot(k) = True

    If errorCode = E_NO_DATA Then
        ' *_______ (204 __)* ______. 0 ________.
        m_detRespBody(k) = ""
        m_detRespBytes(k) = 0
        Exit Sub
    ElseIf errorCode <> 0 Then
        m_detRespBody(k) = ""
        m_detRespEnc(k) = "(____ hr=&H" & Hex(errorCode) & ")"
        Exit Sub
    End If

    If pStream = 0 Then
        ' *_______* (204 / 304 __). 0 ________.
        m_detRespBytes(k) = 0
        Exit Sub
    End If

    ' --- __ (*Seek ________* _________) ---
    If Not NetReadStream(pStream, m_respBodyMax, nBytes, trunc, buf) Then Exit Sub

    m_detRespBytes(k) = nBytes
    m_detRespTrunc(k) = trunc
    If nBytes <= 0 Then Exit Sub

    ' --- *____________, ____ (__2 _P')* ---
    enc = NetHeaderValue(m_detRespHeaders(k), "content-encoding")
    cType = NetHeaderValue(m_detRespHeaders(k), "content-type")
    m_detRespEnc(k) = enc
    m_detRespHex(k) = NetHexHead(buf, nBytes)     ' *____*
    m_detRespIsText(k) = NetIsTextType(cType)
    If m_detRespIsText(k) Then
        m_detRespBody(k) = Utf8BytesToString(buf, nBytes)
    End If
End Sub


' ============================================================
' NetFindDetailBySeq (N-4, Private) - .
'   ________ -1. N-3 _ NetAttachRespHeaders ______.
' ============================================================
Private Function NetFindDetailBySeq(ByVal seqNo As Long) As Long
    Dim i As Long
    Dim k As Long

    NetFindDetailBySeq = -1
    For i = m_detCount To 1 Step -1
        k = NetDetailRingIndex(i)
        If k >= 0 Then
            If m_detSeq(k) = seqNo Then
                NetFindDetailBySeq = k
                Exit Function
            End If
        End If
    Next i
End Function


' ============================================================
' NetRespPendingClear (N-4, Private) - ____ ResponseView .
'   *________________* (______ / Pane ____).
'   _______________ Release __.
' ============================================================
Private Sub NetRespPendingClear()
    Dim v As Variant

    If m_respPending Is Nothing Then Exit Sub

    On Error Resume Next
    For Each v In m_respPending
        If CLngPtr(v) <> 0 Then ComRelease CLngPtr(v)
    Next v
    Err.Clear
    On Error GoTo 0

    Set m_respPending = New Collection
End Sub


Public Function View_RemoveHostObjectFromScript(ByVal objName As String) As Long
    If m_pView = 0 Then
        View_RemoveHostObjectFromScript = &H80004003
        Exit Function
    End If
    View_RemoveHostObjectFromScript = dcf(m_pView, 50, "View.RemoveHostObjectFromScript", StrPtr(objName))
End Function

Public Function View_OpenDevToolsWindow() As Long
    If m_pView = 0 Then
        View_OpenDevToolsWindow = &H80004003
        Exit Function
    End If
    View_OpenDevToolsWindow = dcf(m_pView, 51, "View.OpenDevToolsWindow")
End Function

