Attribute VB_Name = "Wv2Element"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' --- Wv2Element.cls  D-3 __ (_______) ---
'
'   D-3 __________ 4 _:
'
'     Value = v                  Property Let  ______ (*______*)
'     Click()             As Boolean           .
'     SetAttribute(n, v)  As Boolean           .
'     LastInfo            As String            .
'
'   _ _____ (D-2 _ LastOk / LastError _________, __8)
'     Click / SetAttribute ... ___ (Boolean) ____. ________ LastError.
'     Value =              ... *Property Let _________* __,
'                            __________*__ LastOk ___*.
'
'       el.Value = "abc"
'       If Not el.LastOk Then Debug.Print el.LastError
'
'     LastError _____ D-2 ___ (stale / no-pane / in-callback / timeout(n s) /
'     JS ________).
'
'   _ *____JS ____________________________*
'     disabled _____ Click ___________ True ___. readonly .
'     input _________. ________*_________*__.
'
'   _ Value = ________ (__6: *SPA _____*)
'     ___ e.value = x ____, React / Vue .
'     _____*_____*. .
'     value setter (value tracker) ______, .
'     ________________________.
'
'     ___ D-3 _______:
'       (1) e.constructor.prototype __ value _*_____ setter* _____,
'           ______ e ______ (= tracker _____). LastInfo = 'setter'
'           __________________.             LastInfo = 'direct'
'       (2) input _______ (React _______)
'       (3) change _______ (jQuery _______)
'     *____*____6 ___. __ HTML __ change _ 1 _______,
'     ________________________________.
'
'   _ Click __ (__7 _A)
'     e.click() _________ (LastInfo = 'click'). .
'     MouseEvent('click') _ bubbles ___ dispatch __ (LastInfo = 'dispatch').
'     *mousedown / mouseup _____* _____ UI ______________.
'     ___________ isTrusted _ false ___ (JS .
'     ___________). isTrusted ______________.
'
'   _ .
'     Dim p As Wv2Pane, el As Wv2Element
'     Set p = UserForm1.GetActivePane
'     Set el = p.QuerySelector("input[name='q']")
'     el.Value = "___"
'     If el.LastOk Then p.QuerySelector("button[type='submit']").Click
'
'   *D-3 ________________________*
'     Click ___ DOM __________ WaitFor / WaitGone _,
'     __3 _____________ (Wv2Tests.Test_D3_Probe_Promise).
'''''''''''''''''''''''''''''''''''
' --- Wv2Element.cls  D-2 __ (_____________) ---
'
'   *D _ (DOM / _______) ______. D-2 _____. *
'
'   _ .
'     JS _______ (window.__wv2.els) ______ 1 __ DOM ___,
'     VBA ______________. _____ 3 ______:
'
'       m_pane   ... __ Wv2Pane (= ____) .
'       m_handle ... JS _ window.__wv2.els[] .
'       m_gen    ... ______________ (JS ________)
'
'     *DOM _________. * _______ JS _______.
'     _____ 1 _____________ ExecuteScript _ 1 ____.
'     ______________, .
'     Wv2Pane.EvalSync _ JS __ 1 ________ (D-2 _________).
'
'   _ .
'     *New _____. * Wv2Pane.GetElementById / Wv2Pane.QuerySelector ___.
'     __ New ___ m_pane _ Nothing _________.
'
'   _ _____ (D-2 _*______*. _____ Click _ D-3)
'     TagName            As String   ______ (___. _: INPUT)
'     InnerText          As String   .
'     InnerHTML          As String   ___ HTML (____30 _ \u003C _____)
'     Value              As String   input / textarea / select __. .
'     GetAttribute(name) As String   ___. .
'     IsStale            As Boolean  _______________ (__ 1 _)
'     Handle             As Long     JS ____ (_____)
'     Generation         As String   ______ (_____)
'     LastOk             As Boolean  .
'     LastError          As String   .
'
'   _ _____________ (*__*)
'     ___________________, _______ LastOk ____:
'
'       LastOk = True   ... *__* _______ (_______ / _____)
'       LastOk = False  ... *__* ___ LastError:
'                           stale        ____________ (______)
'                           no-pane      Pane _____ (New ______)
'                           in-callback  _______________ (____14)
'                           timeout(n s) JS _______ (____32)
'                           ___        EvalSync _ LastEvalError .
'
'   _ __ (Generation) ____ (D-2 __2 _b)
'     ___*JS _____* (window.__wv2.gen). ______ JS .
'     _____, __________________, _________.
'     VBA ___________________ (_a) _, SPA _ pushState .
'     DOM ____________, ______ JS .
'     ___________, ______. ___*________ JS _*___.
'
'     Wv2Pane __________ (CurrentDomGen) ________ stale .
'     _____________________. .
'     __ JS ______________, ________.
'
'   _ .
'     *Class_Terminate _______. * JS ____________ (D-2 __3).
'     Terminate _____________________, ____ EvalSync .
'     ___ D-1 _ in-callback ____ Ctrl_Close ___ Pane ______.
'     _____________________, .
'     ________ JS ___________. .
'     Wv2Pane.ClearElementRegistry ___. _________ D-4 ___.
''''''''''''''''''''''''''''''''''
Option Explicit


' --- ____ ---
'   m_pane   : ____ Wv2Pane ____ (Wv2Pane __ Wv2Element .
'              __________)
'   m_handle : JS _ window.__wv2.els[] ___ (0 ___)
'   m_gen    : ___________ (JS ________)
Private m_pane   As Wv2Pane
Private m_handle As Long
Private m_gen    As String

' --- _________ (D-1 _ LastEvalOk / LastEvalError _____) ---
Private m_lastOk    As Boolean
Private m_lastError As String

' --- ___*__*_________ (D-3. 'setter' / 'direct' / 'click' _) ---
Private m_lastInfo  As String


Private Sub Class_Initialize()
    Set m_pane = Nothing
    m_handle = -1
    m_gen = ""
    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""
End Sub


Private Sub Class_Terminate()
    ' *_________* (D-2 __3 _a). _____________.
    Set m_pane = Nothing
End Sub


' ============================================================
' InitElement (Friend, Wv2Pane ________)
'   Wv2Pane.GetElementById / QuerySelector _, JS .
'   ____________________.
' ============================================================
Friend Sub InitElement(ByVal ownerPane As Wv2Pane, _
                       ByVal elemHandle As Long, _
                       ByVal elemGen As String)
    Set m_pane = ownerPane
    m_handle = elemHandle
    m_gen = elemGen
    m_lastOk = True
    m_lastError = ""
    m_lastInfo = ""
End Sub


' ============================================================
' Handle / Generation (_________)
' ============================================================
Public Property Get Handle() As Long
    Handle = m_handle
End Property

Public Property Get Generation() As String
    Generation = m_gen
End Property


' ============================================================
' LastOk / LastError (__________)
'
'   *_________________*
'     LastOk = True  -> _______ (__)
'     LastOk = False -> __. LastError .
' ============================================================
Public Property Get LastOk() As Boolean
    LastOk = m_lastOk
End Property

Public Property Get LastError() As String
    LastError = m_lastError
End Property


' ============================================================
' LastInfo (D-3) - ___*__*.
'
'   Value =      : 'setter' (_____ setter __) / 'direct' (____)
'   Click        : 'click'  (e.click())             / 'dispatch' (______)
'   SetAttribute : 'ok'
'
'   *_____ (TagName _) __________* _____________,
'   ______________*______*____.
'   _______, ___________ (___ LastOk / ______).
' ============================================================
Public Property Get LastInfo() As String
    LastInfo = m_lastInfo
End Property


' ============================================================
' TagName / InnerText / InnerHTML / Value (D-2 _____ 4 _)
'
'   ____ 1 _________ ExecuteScript _ 1 ____.
'
'   *InnerHTML _____30 (< _ \u003C ___) _____*
'     Wv2Json.JsonUnescape _______, VBA _____ HTML _____.
'
'   *Value _ input / textarea / select _______*
'     _________ value ________, JS __ undefined _____.
'     ____ LastOk _ True (_______).
' ============================================================
Public Property Get TagName() As String
    TagName = ReadProp("e.tagName")
End Property

Public Property Get InnerText() As String
    InnerText = ReadProp("e.innerText")
End Property

Public Property Get InnerHTML() As String
    InnerHTML = ReadProp("e.innerHTML")
End Property

Public Property Get value() As String
    value = ReadProp("e.value")
End Property


' ============================================================
' GetAttribute (D-2)
'   ________. __________ + LastOk = True.
'
'   ___:
'     el.GetAttribute("href")
'     el.GetAttribute("data-id")
'
'   *getAttribute _ .value ___*
'     <input value="___"> _____________,
'       .Value               -> ____ (_____)
'       .GetAttribute("value") -> ____ (______)
'     ___. __________.
' ============================================================
Public Function GetAttribute(ByVal attrName As String) As String
    If m_pane Is Nothing Then
        m_lastOk = False
        m_lastError = "no-pane"
        GetAttribute = ""
        Exit Function
    End If
    GetAttribute = ReadProp("e.getAttribute(" & m_pane.JsQuote(attrName) & ")")
End Function


' ============================================================
' Value (D-3, Property Let) - .
'
'   *Property Let _________* _____ LastOk _____ (__8).
'   *________________ 3 __* (______):
'     _____ setter <- input ____ <- change ____ .
'
'   ___:
'     el.Value = "___"
'     If Not el.LastOk Then Debug.Print "______: " & el.LastError
'
'   *input / textarea / select ___________*
'     div ___ value ____________, _____ setter .
'     ________ (LastInfo = 'direct'). JS ___________ 1 .
'     ________________, *LastOk _ True ___*.
'     ____________________.
' ============================================================
Public Property Let value(ByVal newValue As String)
    Dim jsVal As String
    Dim js As String

    If m_pane Is Nothing Then
        m_lastOk = False
        m_lastError = "no-pane"
        m_lastInfo = ""
        Exit Property
    End If

    jsVal = m_pane.JsQuote(newValue)

    js = "var d=Object.getOwnPropertyDescriptor(e.constructor.prototype,'value');" & _
         "if(d&&d.set){d.set.call(e," & jsVal & ");rv='setter';}" & _
         "else{e.value=" & jsVal & ";rv='direct';}" & _
         "e.dispatchEvent(new Event('input',{bubbles:true}));" & _
         "e.dispatchEvent(new Event('change',{bubbles:true}));"

    ExecAction js, "Value"
End Property


' ============================================================
' Click (D-3) - .
'
'   ___: True = JS _________ / False = __ (___ LastError)
'
'   e.click() _________ (LastInfo = 'click'), .
'   MouseEvent('click') _ bubbles + cancelable ___ dispatch .
'   (LastInfo = 'dispatch'). ____*_________ isTrusted _ false*.
'
'   *___ True ____________________*
'     disabled ______________ True ___. ______ DOM .
'     ______________, ________ D-3 _________.
'
'   ___:
'     If Not el.Click() Then Debug.Print el.LastError
' ============================================================
Public Function Click() As Boolean
    Dim js As String

    js = "if(typeof e.click==='function'){e.click();rv='click';}" & _
         "else{e.dispatchEvent(new MouseEvent('click'," & _
         "{bubbles:true,cancelable:true,view:window}));rv='dispatch';}"

    Click = ExecAction(js, "Click")
End Function


' ============================================================
' SetAttribute (D-3) - .
'
'   ___: True = _____ / False = __ (___ LastError)
'
'   *GetAttribute _ .Value ____ D-2 _______*
'     SetAttribute("value", ...) _*__*_______, .
'     input _*___*______. ________ .Value = _____.
'
'   _________ (_____) _ JS _ InvalidCharacterError ______,
'   False + LastError ______________.
'
'   ___:
'     el.SetAttribute "data-id", "42"
' ============================================================
Public Function SetAttribute(ByVal attrName As String, _
                             ByVal attrValue As String) As Boolean
    Dim js As String

    SetAttribute = False

    If m_pane Is Nothing Then
        m_lastOk = False
        m_lastError = "no-pane"
        m_lastInfo = ""
        Exit Function
    End If

    js = "e.setAttribute(" & m_pane.JsQuote(attrName) & "," & _
         m_pane.JsQuote(attrValue) & ");rv='ok';"

    SetAttribute = ExecAction(js, "SetAttribute")
End Function

' ============================================================
' IsStale (D-2)
'   _______________________.
'
'   True _____:
'     _________ JS ______________ (_______)
'     ________________ (________)
'   __________ (______) ________ True ___.
'   ____ LastOk = False ___, _______ LastError _____.
'
'   *DOM ___________ stale _____*
'     _______________ False ___. detached .
'     JS ______________, ___________.
' ============================================================
Public Property Get IsStale() As Boolean
    Dim st As String
    Dim stale As Boolean

    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""

    If m_pane Is Nothing Then
        m_lastError = "no-pane"
        IsStale = True
        Exit Property
    End If

    stale = m_pane.Elem_CheckStale(m_handle, m_gen, st)

    If Len(st) = 0 Then
        m_lastOk = True
    Else
        m_lastError = st
    End If

    IsStale = stale
End Property


' ============================================================
' ReadProp (____: _____ 1 ____)
'   valueExpr _ JS ____ e (= ____) _____.
' ============================================================
Private Function ReadProp(ByVal valueExpr As String) As String
    Dim v As String
    Dim st As String
    Dim ok As Boolean

    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""
    ReadProp = ""

    If m_pane Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    ok = m_pane.Elem_Read(m_handle, m_gen, valueExpr, v, st)

    If ok Then
        m_lastOk = True
        ReadProp = v
    Else
        m_lastError = st
    End If
End Function


' ============================================================
' ExecAction (D-3, ____: ___ 1 ___)
'   actionStmts _ JS ____ e (____) ____*_*.
'   _____________ JS ____ rv _____.
'   ReadProp ______________, ______________.
' ============================================================
Private Function ExecAction(ByVal actionStmts As String, _
                            ByVal apiName As String) As Boolean
    Dim v As String
    Dim st As String

    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""
    ExecAction = False

    If m_pane Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    If m_pane.Elem_Exec(m_handle, m_gen, actionStmts, v, st) Then
        m_lastOk = True
        m_lastInfo = v
        ExecAction = True

        ' --- D-4b: opt-in _____ (__ False _________) ---
        '   *___________* ___________,
        '   ____________________ (D-3 ______).
        '   ______ m_pane.LastSettleInfo ___.
        If m_pane.AutoWaitAfterAction Then m_pane.WaitSettled
    Else
        m_lastError = st
        Wv2Log.LogW "Wv2Element." & apiName & ": __ (handle=" & m_handle & _
                    ") err=" & st
    End If
End Function


