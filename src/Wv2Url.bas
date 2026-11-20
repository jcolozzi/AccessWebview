Attribute VB_Name = "Wv2Url"
''''''''''''''''''''''''''''''''''
' --- Wv2Url.bas  _9.16 __ (__: URL ______________) ---
'
'   URL ___________________________ Wv2Browser .
'   ____________. Wv2Json.bas (JSON __) .
'   ___________. _______ (9.13 _ JSON ________).
'
'   *_____ (_____ = _____)*
'     _IsAlphaOnly(s)          : _________ (__________)
'     _UrlEncodeComponent(s)   : encodeURIComponent ___ percent-encoding
'     _Utf8Bytes(s)            : String(UTF-16LE) -> UTF-8 ____ (ADODB.Stream)
'
'   *_______ (____ = _____ Wv2Browser ___)*
'     NormalizeUrl / BuildSearchUrl _________ (m_searchTemplate) .
'     ___________________________ Wv2Browser ___,
'     __ Wv2Url _________. _____________.
'
'   *ADODB.Stream ___* Utf8Bytes _ ADODB.Stream ____, ___ WebView2
'     __ (____vtable__________) _________________,
'     __ HTML _________. _________WebView2 ___ COM .
'     ____________ (9.14 ______).
''''''''''''''''''''''''''''''''''

Option Explicit


' ============================================================
' IsAlphaOnly (_9.16 _ Wv2Browser ____, _______)
'
'   ______ (A-Z / a-z) ____________.
'   NormalizeUrl _ "about:" .
'   ("example.com:8080" _______________, ________ False).
' ============================================================
Public Function IsAlphaOnly(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    Dim c As String
    For i = 1 To Len(s)
        c = LCase$(Mid$(s, i, 1))
        If c < "a" Or c > "z" Then Exit Function
    Next i
    IsAlphaOnly = True
End Function


' ============================================================
' UrlEncodeComponent (_9.16 _ Wv2Browser ____, _______)
'
'   ____ RFC3986 ____ percent-encoding .
'   (JS _ encodeURIComponent __. __________, __ %XX).
'
'   *_____ (unreserved)* A-Z a-z 0-9 _ "-" "_" "." "~" _____.
'     ________ UTF-8 ____________ %XX ___.
'     (___ %20. encodeURIComponent ____ "+" _____)
'
'   *UTF-8 __ Utf8Bytes (ADODB.Stream __) ___* .
'     _______________. 1 ____ AscW .
'     _______ (____) _____, _________ UTF-8 .
'     _____________.
' ============================================================
Public Function UrlEncodeComponent(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function

    Dim bytes() As Byte
    bytes = Utf8Bytes(s)

    ' Utf8Bytes __ (Len(s)=0 ________________) .
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(bytes)
    ub = UBound(bytes)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Dim sb As String
    Dim i As Long
    Dim b As Integer
    Dim c As String
    For i = lb To ub
        b = bytes(i)            ' 0..255
        c = Chr$(b)             ' ASCII ______ (b<128 __________)
        If (b >= 65 And b <= 90) _
        Or (b >= 97 And b <= 122) _
        Or (b >= 48 And b <= 57) _
        Or b = 45 Or b = 95 Or b = 46 Or b = 126 Then
            ' unreserved: A-Z a-z 0-9 - _ . ~
            sb = sb & c
        Else
            sb = sb & "%" & Right$("0" & Hex$(b), 2)
        End If
    Next i

    UrlEncodeComponent = sb
End Function


' ============================================================
' Utf8Bytes (_9.16 _ Wv2Browser ____, _______)
'
'   VBA _ String (UTF-16LE) _ UTF-8 __________.
'   ADODB.Stream _ Charset="utf-8" _____, Type=adTypeBinary _____.
'
'   *__ BOM _____* Charset="utf-8" ___________ BOM .
'     ____ (adTypeBinary ___________ EF BB BF ______).
'     __________, __ 3 ____ BOM ________________.
' ============================================================
Public Function Utf8Bytes(ByVal s As String) As Byte()
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1                 ' adTypeBinary
    ' _________________ Position ______.
    Dim raw() As Byte
    raw = st.Read                ' .
    st.Close
    Set st = Nothing

    ' BOM __ (________, ______ EF BB BF ____)
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(raw)
    ub = UBound(raw)
    If Err.Number <> 0 Then
        ' _______________. ______.
        On Error GoTo 0
        Utf8Bytes = raw
        Exit Function
    End If
    On Error GoTo 0

    If (ub - lb) >= 2 Then
        If raw(lb) = &HEF And raw(lb + 1) = &HBB And raw(lb + 2) = &HBF Then
            Dim out() As Byte
            ReDim out(0 To (ub - lb) - 3)
            Dim j As Long
            For j = lb + 3 To ub
                out(j - (lb + 3)) = raw(j)
            Next j
            Utf8Bytes = out
            Exit Function
        End If
    End If

    Utf8Bytes = raw
End Function


