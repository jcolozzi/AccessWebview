Attribute VB_Name = "Wv2Json"
''''''''''''''''''''''''''''''''''
' --- Wv2Json.bas  D-2 __ (JSON __________) ---
'
'   D-2 _____ (__6 _b):
'     *JsonUnescapeAt / JsonUnescape / JsonPickStr _____*
'       D-1 _ Wv2Pane _ Private ______ JsonPickStr _, .
'       ____ Public ___. D-2 __ Wv2Pane (EvalSync ______) .
'       Wv2Element (___________) _ 2 _______________,
'       ______________________ (____70).
'
'     JsonUnescapeAt(json, startPos, outVal) As Long
'       JSON ________*__*_ startPos _____, .
'       outVal ____. ____________ (________ 0).
'       startPos ________*_*_________.
'
'     JsonUnescape(raw) As String
'       ____ 1 __ JSON ______, .
'       ________. ________ (__ / true / null / _______)
'       *_________*. EvalSync __________________.
'
'     JsonPickStr(json, key, outVal) As Boolean
'       "key":"value" _ value _, _____________.
'
'   *JsonGetStr ______*
'     JsonGetStr _*__________*_____, __ JS ___ cmd / index .
'     ___________________________ (____64).
'     _________ (innerText / innerHTML / ___) _____________,
'     _____ JsonPickStr / JsonUnescape _____. JsonGetStr .
'     __________*_________*.
'
'   *__________ (____30)*
'     __ JSON __ < _ \u003C ___________. D-2 _ innerHTML .
'     ____________, \uXXXX ____________.
'
'   *_______*
'     \uD83D\uDE00 ____ 2 ___, ChrW$ _ 2 .
'     (VBA _ String _ UTF-16 ___________, ______________).
'     *Python ______________* Python _ str ____________,
'     _____________________________ NG ___.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Json.bas  _9.13 __ (__: JSON ________) ---
'
'   *VBA_JS _________ JSON ____*
'
'   _9.11c_9.12d _ Wv2TabBar _ Wv2NavBar ___ Private .
'   ______ (JsonEscape / GetJsonStr / GetJsonNum / BoolToJson). 3 ___ UI .
'   ________________ (____________).
'
'   *__ Wv2Thunks.bas ______*
'     Wv2Thunks.bas ______________ vtable_SAFEARRAY .
'     _________________, __ 2600 __. JSON .
'     ________, __________________.
'
'   *____ 64* __ HTML _ JS _________, __ JSON ________.
'     __________ JS _________________, .
'     __________________. __________________.
'     (VBA ___ JSON ______, __ COM ________________)
'
'   *____* ________ Public _____, .
'     _____ Json ____________:
'       JsonEscape  ... ____ JSON .
'       JsonGetStr  ... "key":"value" _ value .
'       JsonGetNum  ... "key":123     _ 123 .
'       JsonBool    ... Boolean _ true/false .
'
'   ______ (9.13 ___):
'       JsonEscape   <- JsonEscape  (____)
'       JsonGetStr   <- GetJsonStr
'       JsonGetNum   <- GetJsonNum
'       JsonBool     <- BoolToJson
''''''''''''''''''''''''''''''''''

Option Explicit


' ============================================================
' JsonEscape ? ____ JSON .
'
'   " _ \ _____ (0x00-0x1F) ________.
'   _______ URL _ JSON ___________.
'
'   *_9.13 ___: _______________ (____ 16)*
'     VBA _ AscW _ 0x8000 __________ Integer _____.
'     _: ___U+9032 = 36914 -> AscW _ -28622 ___.
'     ___________ Case Is < 32 ________, .
'     ____________, \uXXXX _____________.
'     *___ JSON ______, JS ___________________*_,
'     __________, _______________________.
'     -> __________ ch >= 0 And ch < 32 ___, __________.
'       __ (JS ______) ______.
' ============================================================
Public Function JsonEscape(ByVal s As String) As String
    Dim out As String
    Dim i As Long
    Dim ch As Long
    Dim c As String

    out = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        ch = AscW(c)
        Select Case ch
            Case 34            ' " (_____)
                out = out & "\"""
            Case 92            ' \ (________)
                out = out & "\\"
            Case 8             ' backspace
                out = out & "\b"
            Case 9             ' tab
                out = out & "\t"
            Case 10            ' LF
                out = out & "\n"
            Case 12            ' FF
                out = out & "\f"
            Case 13            ' CR
                out = out & "\r"
            Case Else
                ' *_9.13 __* _______ \uXXXX ___.
                '   ch >= 0 _____ (AscW _____________. ____ 16).
                '   ____________ \uXXXX ____.
                If ch >= 0 And ch < 32 Then
                    out = out & "\u" & Right$("0000" & Hex$(ch), 4)
                Else
                    out = out & c
                End If
        End Select
    Next i

    JsonEscape = out
End Function


' ============================================================
' JsonGetStr ? "key":"value" _ value .
'
'   _: JsonGetStr("{""cmd"":""activate"",""index"":3}", "cmd") -> "activate"
'   ______________.
'
'   *__________* _______ JS _, cmd _ url .
'     ___________. _________________ (____ 64).
' ============================================================
Public Function JsonGetStr(ByVal json As String, ByVal key As String) As String
    Dim pat As String
    pat = """" & key & """:"""
    ' pat ____  "key":"  ___ 7 __ + .

    Dim p As Long
    p = InStr(json, pat)
    If p = 0 Then
        JsonGetStr = ""
        Exit Function
    End If

    Dim st As Long
    st = p + Len(pat)                ' .

    Dim en As Long
    en = InStr(st, json, """")       ' .
    If en = 0 Then
        JsonGetStr = ""
        Exit Function
    End If

    JsonGetStr = Mid$(json, st, en - st)
End Function


' ============================================================
' JsonGetNum ? "key":123 _ 123 .
'
'   _: JsonGetNum("{""cmd"":""close"",""index"":2}", "index") -> 2
'   ________ 0 ___.
'
'   *__________* __ index (1-based ____) _________.
'     ___________ (____ 64).
' ============================================================
Public Function JsonGetNum(ByVal json As String, ByVal key As String) As Long
    Dim pat As String
    pat = """" & key & """:"
    ' pat ____  "key":  ___ 4 __ + .

    Dim p As Long
    p = InStr(json, pat)
    If p = 0 Then
        JsonGetNum = 0
        Exit Function
    End If

    Dim st As Long
    st = p + Len(pat)

    ' _______ (JSON.stringify __________)
    Do While st <= Len(json)
        If Mid$(json, st, 1) <> " " Then Exit Do
        st = st + 1
    Loop

    ' .
    Dim buf As String
    Dim c As String
    Do While st <= Len(json)
        c = Mid$(json, st, 1)
        If c < "0" Or c > "9" Then Exit Do
        buf = buf & c
        st = st + 1
    Loop

    If Len(buf) = 0 Then
        JsonGetNum = 0
    Else
        JsonGetNum = CLng(buf)
    End If
End Function


' ============================================================
' JsonBool ? Boolean _ JSON _ true/false .
'
'   *CStr(True) _______* CStr(True) _ "True" (_____) _____,
'     ____________. JSON _____ true/false .
'     ________.
' ============================================================
Public Function JsonBool(ByVal b As Boolean) As String
    If b Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
End Function


' ============================================================
' JsonUnescapeAt (D-2) - JSON .
'
'   __:
'     json     : ___ JSON ___ (__)
'     startPos : ______. *__________*_____ (1 ___)
'     outVal   : ______ (___________)
'   ___:
'     ________. __________ 0 (____ outVal ____)
'
'   _________: ___ / ________ / _____ / b / f / n / r / t / uXXXX
'   (_______________ 1 __. uXXXX __ 4 __ 16 ____)
'   _____ \x _ x _________ (JSON ________, _____).
' ============================================================
Public Function JsonUnescapeAt(ByVal json As String, _
                               ByVal startPos As Long, _
                               ByRef outVal As String) As Long
    Dim i As Long
    Dim buf As String
    Dim c As String
    Dim d As String

    outVal = ""
    JsonUnescapeAt = 0

    If startPos < 1 Then Exit Function

    i = startPos
    Do While i <= Len(json)
        c = Mid$(json, i, 1)

        If c = "\" Then
            d = Mid$(json, i + 1, 1)
            Select Case d
                Case """"
                    buf = buf & """"
                Case "\"
                    buf = buf & "\"
                Case "/"
                    buf = buf & "/"
                Case "b"
                    buf = buf & Chr$(8)
                Case "f"
                    buf = buf & Chr$(12)
                Case "n"
                    buf = buf & vbLf
                Case "r"
                    buf = buf & vbCr
                Case "t"
                    buf = buf & vbTab
                Case "u"
                    ' *4 ____________* CLng("&HD83D") _ 16 .
                    '   _______ -10179 ____, ChrW$ ____ 65536 + n .
                    '   ______ 55357 (= D83D) ________.
                    '   __________________________.
                    buf = buf & ChrW$(CLng("&H" & Mid$(json, i + 2, 4)))
                    i = i + 4
                Case Else
                    buf = buf & d
            End Select
            i = i + 2

        ElseIf c = """" Then
            outVal = buf
            JsonUnescapeAt = i
            Exit Function

        Else
            buf = buf & c
            i = i + 1
        End If
    Loop
End Function


' ============================================================
' JsonUnescape (D-2) - JSON _ 1 .
'
'   _______ (______) _________________.
'   ____ (__ / true / false / null / undefined / ______ / __) .
'   *______*. EvalSync ____________.
'
'   __________________, _________ (_____,
'   __________________).
' ============================================================
Public Function JsonUnescape(ByVal raw As String) As String
    Dim v As String
    Dim en As Long

    If Len(raw) = 0 Then
        JsonUnescape = ""
        Exit Function
    End If

    If Left$(raw, 1) <> """" Then
        JsonUnescape = raw
        Exit Function
    End If

    en = JsonUnescapeAt(raw, 2, v)
    If en > 0 Then
        JsonUnescape = v
    Else
        JsonUnescape = raw
    End If
End Function


' ============================================================
' JsonPickStr (D-2, D-1 _ Wv2Pane _________)
'   "key":"value" _ value _, _____________.
'
'   __:
'     json   : ___ JSON .
'     key    : ______ (__________)
'     outVal : ______ (___________)
'   ___:
'     True  = _______, .
'     False = _____, .
' ============================================================
Public Function JsonPickStr(ByVal json As String, _
                            ByVal key As String, _
                            ByRef outVal As String) As Boolean
    Dim pat As String
    Dim p As Long
    Dim en As Long

    outVal = ""
    JsonPickStr = False

    pat = """" & key & """:"""
    p = InStr(1, json, pat)
    If p = 0 Then Exit Function

    en = JsonUnescapeAt(json, p + Len(pat), outVal)
    JsonPickStr = (en > 0)
    If Not JsonPickStr Then outVal = ""
End Function


