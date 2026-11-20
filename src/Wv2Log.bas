Attribute VB_Name = "Wv2Log"
''''''''''''''''''''''''''''''''''
' --- Wv2Log.bas  K-4b __ (___ GetLocalTime 1 ____) ---
'
'   *K-4a ______________*
'     000030  22:45:23.992
'     000031  22:45:23.000   <- *___________* (____ 24.000)
'     000032  22:45:24.008
'
'   ___ *Now _ Timer _________________*.
'   Timer _________________________ 000____.
'   (___________. ______________)
'
'   *__* ______________1008 ms .
'   _______. *___________________________. *
'
'   ____ GetLocalTime _ 1 _____. .
'   *________*.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Log.bas  K-1 __ (_____________) ---
'
'   *_____*
'     D-2 ____, .
'     ____________. ___________________,
'     *___________*________. D __ 1 .
'     ______. ____________.
'
'   *_____*
'     _LogE / LogW / LogI / LogD : _____ 1 .
'     _LogPath                   : ______________ (Property Get)
'     _LogLevel                  : ____ (Property Get/Let). ___ LOG_DEBUG
'     _LogEcho                   : Debug.Print ____ (Property Get/Let). ___ True
'     _LogStart / LogStop / LogFlush : .
'
'   *_____ (K-1 _ 9 _______)*
'     _____ %APPDATA%\Wv2Browser\logs\. settings.txt ________.
'       ____________ OneDrive .
'     _*_____ 1 _* (wv2_yyyymmdd_hhnnss.log). _____ 20 ________.
'       D ___1 ___ = ________, ___ 1 ___ 1 .
'     _*LogPrint ____ Debug.Print ___* ______ 1 ___.
'       ___ Debug.Print "..." -> Wv2Log.LogD "..." _____ 1 _ 1 _______,
'       _________________. LogEcho = False .
'     _*___________* ____. ERROR / WARN .
'       _____ Shared ___. *Lock Write __ ADODB .
'       ______ (K-1 ________)* _______________,
'       _____________________.
'       (_______). ________________, ____20 .
'       .
'     _*__________ + _____* (____78 ____). .
'       ____________________ K-1 _____________,
'       .
'     _*_______* __________________. K-1 .
'
'   *_____*
'     _______ UTF-8 / BOM __ / CRLF. UTF-8 .
'     _______ (ADODB.Stream __________). ____ String .
'     Web _________ Unicode _______, _____________.
'     _ VBA _______ CP932 _______ (____35) _____.
'
'   *Wv2Tests.WriteUtf8NoBom _______*
'     Wv2Tests ______. ____________ (____ 12 ________).
'     __________________.
''''''''''''''''''''''''''''''''''

Option Explicit

' ===== _____ =====
Public Const LOG_ERROR As Long = 1
Public Const LOG_WARN As Long = 2
Public Const LOG_INFO As Long = 3
Public Const LOG_DEBUG As Long = 4

' --- K-4b: ___ GetLocalTime _ 1 ____ ---
'   *Now _ Timer _____* _______________ (______).
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" ( _
    ByRef lpSystemTime As SYSTEMTIME)

Private Const LOG_KEEP_FILES As Long = 20
Private Const LOG_PREFIX As String = "wv2_"

Private m_logNum As Long
Private m_logPath As String
Private m_logLevel As Long
Private m_logEcho As Boolean
Private m_logDepth As Long
Private m_logPending As String
Private m_logSeq As Long
Private m_logInit As Boolean


' ============================================================
' LogE / LogW / LogI / LogD
'
'   _____ 1 ___. _______________.
'   ERROR _ WARN ________________.
' ============================================================
Public Sub LogE(ByVal logMsg As String)
    LogWrite LOG_ERROR, logMsg
End Sub

Public Sub LogW(ByVal logMsg As String)
    LogWrite LOG_WARN, logMsg
End Sub

Public Sub LogI(ByVal logMsg As String)
    LogWrite LOG_INFO, logMsg
End Sub

Public Sub LogD(ByVal logMsg As String)
    LogWrite LOG_DEBUG, logMsg
End Sub


' ============================================================
' LogPath / LogLevel / LogEcho
' ============================================================
Public Property Get LogPath() As String
    LogPath = m_logPath
End Property

Public Property Get LogLevel() As Long
    If Not m_logInit Then InitDefaults
    LogLevel = m_logLevel
End Property

Public Property Let LogLevel(ByVal logArg As Long)
    If Not m_logInit Then InitDefaults
    If logArg < LOG_ERROR Then
        m_logLevel = LOG_ERROR
    ElseIf logArg > LOG_DEBUG Then
        m_logLevel = LOG_DEBUG
    Else
        m_logLevel = logArg
    End If
End Property

Public Property Get LogEcho() As Boolean
    If Not m_logInit Then InitDefaults
    LogEcho = m_logEcho
End Property

Public Property Let LogEcho(ByVal logArg As Boolean)
    If Not m_logInit Then InitDefaults
    m_logEcho = logArg
End Property


' ============================================================
' InitDefaults
'
'   ___________. _________ Class_Initialize _____,
'   _____________ 1 _____.
' ============================================================
Private Sub InitDefaults()
    If m_logInit Then Exit Sub
    m_logInit = True
    m_logLevel = LOG_DEBUG
    m_logEcho = True
    m_logSeq = 0
    m_logDepth = 0
    m_logPending = ""
End Sub


' ============================================================
' LogStart
'
'   ____________. _________________.
'   _____________ (___ LogD _______).
'   _______________.
' ============================================================
Public Sub LogStart()
    On Error Resume Next
    LogStop
    On Error GoTo 0

    InitDefaults
    m_logSeq = 0

    Dim logFolder As String
    logFolder = LogFolderPath()
    If Len(logFolder) = 0 Then Exit Sub

    RotateOldLogs logFolder

    m_logPath = logFolder & LOG_PREFIX & Format$(Now, "yyyymmdd_hhnnss") & ".log"

    On Error GoTo eh
    m_logNum = FreeFile
    Open m_logPath For Binary Access Write Shared As #m_logNum
    Seek #m_logNum, LOF(m_logNum) + 1

    Dim logHead As String
    logHead = "=== Wv2Log __ " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ===" & vbCrLf & _
              "=== path = " & m_logPath & " ===" & vbCrLf & _
              "=== ___ 6 ____. ______________ ===" & vbCrLf
    PutBytes Utf8FromString(logHead)

    Debug.Print "Wv2Log.LogStart: " & m_logPath
    Exit Sub
eh:
    Debug.Print "Wv2Log.LogStart: __ (" & Err.Number & ") " & Err.Description
    m_logNum = 0
    m_logPath = ""
End Sub


' ============================================================
' LogStop
' ============================================================
Public Sub LogStop()
    If m_logNum = 0 Then Exit Sub
    On Error Resume Next
    Close #m_logNum
    On Error GoTo 0
    m_logNum = 0
End Sub


' ============================================================
' LogFlush
'
'   __________, OS ______________.
'   ERROR / WARN ____ LogWrite ______.
' ============================================================
Public Sub LogFlush()
    If m_logNum = 0 Then Exit Sub
    If Len(m_logPath) = 0 Then Exit Sub

    On Error GoTo eh
    Close #m_logNum
    m_logNum = FreeFile
    Open m_logPath For Binary Access Write Shared As #m_logNum
    Seek #m_logNum, LOF(m_logNum) + 1
    Exit Sub
eh:
    Debug.Print "Wv2Log.LogFlush: __ (" & Err.Number & ") " & Err.Description
    m_logNum = 0
End Sub


' ============================================================
' LogWrite  (__)
'
'   ________ return
'   _Debug.Print ____ (LogEcho)
'   _*_____* ___________________, .
' ============================================================
Private Sub LogWrite(ByVal logArg As Long, ByVal logMsg As String)
    If Not m_logInit Then InitDefaults
    If logArg > m_logLevel Then Exit Sub

    m_logSeq = m_logSeq + 1

    Dim logLine As String
    logLine = Right$("00000" & CStr(m_logSeq), 6) & " " & NowStamp() & _
              " [" & LevelTag(logArg) & "] " & logMsg

    If m_logEcho Then Debug.Print logLine

    ' --- ___________ (____78 ____) ---
    If m_logDepth > 0 Then
        m_logPending = m_logPending & logLine & vbCrLf
        Exit Sub
    End If

    m_logDepth = m_logDepth + 1
    On Error Resume Next

    If m_logNum = 0 Then LogStart
    If m_logNum <> 0 Then
        PutBytes Utf8FromString(logLine & vbCrLf)

        ' --- ____________ ---
        If Len(m_logPending) > 0 Then
            Dim logDrain As String
            logDrain = m_logPending
            m_logPending = ""
            PutBytes Utf8FromString(logDrain)
        End If
    End If

    On Error GoTo 0
    m_logDepth = m_logDepth - 1

    If logArg <= LOG_WARN Then LogFlush
End Sub


' ============================================================
' Debug_SetLogDepth  (____)
'
'   __________________. Wv2Pane.Debug_SetInCallback ____.
'   1 ______ LogWrite ______________, 0 .
'   __ 1 _________. *____________*
' ============================================================
Public Sub Debug_SetLogDepth(ByVal logArg As Long)
    If Not m_logInit Then InitDefaults
    If logArg < 0 Then logArg = 0
    m_logDepth = logArg
End Sub

' ============================================================
' PutBytes
' ============================================================
Private Sub PutBytes(ByRef outBytes() As Byte)
    On Error Resume Next
    If m_logNum = 0 Then Exit Sub
    Put #m_logNum, , outBytes
End Sub


' ============================================================
' LevelTag
' ============================================================
Private Function LevelTag(ByVal logArg As Long) As String
    Select Case logArg
        Case LOG_ERROR: LevelTag = "ERR "
        Case LOG_WARN:  LevelTag = "WARN"
        Case LOG_INFO:  LevelTag = "INFO"
        Case Else:      LevelTag = "DBG "
    End Select
End Function


' ============================================================
' NowStamp (K-4b ______)
'
'   yyyy-mm-dd hh:nn:ss.mmm
'
'   *GetLocalTime _ 1 ___* _____ ______________,
'   ___________________.
'   ___ Now (___) _ Timer (___) ___________,
'   *________________ 000__________* (K-4a ___).
' ============================================================
Private Function NowStamp() As String
    Dim st As SYSTEMTIME

    GetLocalTime st

    NowStamp = Format$(st.wYear, "0000") & "-" & _
               Format$(st.wMonth, "00") & "-" & _
               Format$(st.wDay, "00") & " " & _
               Format$(st.wHour, "00") & ":" & _
               Format$(st.wMinute, "00") & ":" & _
               Format$(st.wSecond, "00") & "." & _
               Format$(st.wMilliseconds, "000")
End Function


' ============================================================
' LogFolderPath
'
'   %APPDATA%\Wv2Browser\logs\ ______. ________.
' ============================================================
Private Function LogFolderPath() As String
    On Error GoTo eh
    Dim logFolder As String
    logFolder = Environ$("APPDATA")
    If Len(logFolder) = 0 Then Exit Function
    If Right$(logFolder, 1) <> "\" Then logFolder = logFolder & "\"

    logFolder = logFolder & "Wv2Browser\"
    If Dir(logFolder, vbDirectory) = "" Then MkDir logFolder

    logFolder = logFolder & "logs\"
    If Dir(logFolder, vbDirectory) = "" Then MkDir logFolder

    LogFolderPath = logFolder
    Exit Function
eh:
    Debug.Print "Wv2Log.LogFolderPath: __ (" & Err.Number & ") " & Err.Description
End Function


' ============================================================
' RotateOldLogs
'
'   wv2_*.log ______ LOG_KEEP_FILES ____, ___________.
'   ________________, ___ = ______.
'   *Dir ________, ___________*
' ============================================================
Private Sub RotateOldLogs(ByVal logFolder As String)
    On Error GoTo eh

    Dim logNames() As String
    Dim logCount As Long
    ReDim logNames(0 To 255)
    logCount = 0

    Dim logName As String
    logName = Dir(logFolder & LOG_PREFIX & "*.log")
    Do While Len(logName) > 0
        If logCount > UBound(logNames) Then ReDim Preserve logNames(0 To logCount + 255)
        logNames(logCount) = logName
        logCount = logCount + 1
        logName = Dir
    Loop

    If logCount <= LOG_KEEP_FILES Then Exit Sub

    ' --- _______ (________. __________) ---
    Dim logIdx As Long
    Dim logLow As Long
    Dim logHigh As String
    For logIdx = 1 To logCount - 1
        logHigh = logNames(logIdx)
        logLow = logIdx - 1
        Do While logLow >= 0
            If logNames(logLow) <= logHigh Then Exit Do
            logNames(logLow + 1) = logNames(logLow)
            logLow = logLow - 1
        Loop
        logNames(logLow + 1) = logHigh
    Next logIdx

    Dim logKept As Long
    logKept = logCount - LOG_KEEP_FILES
    For logIdx = 0 To logKept - 1
        On Error Resume Next
        Kill logFolder & logNames(logIdx)
        On Error GoTo 0
    Next logIdx

    Debug.Print "Wv2Log.RotateOldLogs: " & logKept & " ____ (" & LOG_KEEP_FILES & " ____)"
    Exit Sub
eh:
    Debug.Print "Wv2Log.RotateOldLogs: __ (" & Err.Number & ") " & Err.Description
End Sub


' ============================================================
' Utf8FromString
'
'   String (UTF-16 ________) _ UTF-8 _______.
'   *________ 4 ______* ____ String __ Web .
'   ___ Unicode _______.
'   ADODB.Stream _________________.
' ============================================================
Private Function Utf8FromString(ByVal logText As String) As Byte()
    Dim outBytes() As Byte
    Dim logLen As Long
    Dim logIdx As Long
    Dim logCount As Long
    Dim logCode As Long
    Dim logLow As Long

    logLen = Len(logText)
    If logLen = 0 Then
        ReDim outBytes(0 To 0)
        Utf8FromString = outBytes
        Exit Function
    End If

    ' __ 1 __ 4 .
    ReDim outBytes(0 To logLen * 4 - 1)
    logCount = 0
    logIdx = 1

    Do While logIdx <= logLen
        logCode = AscW(Mid$(logText, logIdx, 1))
        ' *AscW _ 0x8000 _____ Integer ___ (____16)*
        If logCode < 0 Then logCode = logCode + 65536

        ' --- _______ (__ D800-DBFF + __ DC00-DFFF) ---
        If logCode >= &HD800& And logCode <= &HDBFF& And logIdx < logLen Then
            logLow = AscW(Mid$(logText, logIdx + 1, 1))
            If logLow < 0 Then logLow = logLow + 65536
            If logLow >= &HDC00& And logLow <= &HDFFF& Then
                logCode = &H10000 + (logCode - &HD800&) * &H400 + (logLow - &HDC00&)
                logIdx = logIdx + 1
            End If
        End If

        If logCode < &H80 Then
            outBytes(logCount) = CByte(logCode)
            logCount = logCount + 1
        ElseIf logCode < &H800 Then
            outBytes(logCount) = CByte(&HC0 + (logCode \ &H40))
            outBytes(logCount + 1) = CByte(&H80 + (logCode Mod &H40))
            logCount = logCount + 2
        ElseIf logCode < &H10000 Then
            outBytes(logCount) = CByte(&HE0 + (logCode \ &H1000))
            outBytes(logCount + 1) = CByte(&H80 + ((logCode \ &H40) Mod &H40))
            outBytes(logCount + 2) = CByte(&H80 + (logCode Mod &H40))
            logCount = logCount + 3
        Else
            outBytes(logCount) = CByte(&HF0 + (logCode \ &H40000))
            outBytes(logCount + 1) = CByte(&H80 + ((logCode \ &H1000) Mod &H40))
            outBytes(logCount + 2) = CByte(&H80 + ((logCode \ &H40) Mod &H40))
            outBytes(logCount + 3) = CByte(&H80 + (logCode Mod &H40))
            logCount = logCount + 4
        End If

        logIdx = logIdx + 1
    Loop

    ReDim Preserve outBytes(0 To logCount - 1)
    Utf8FromString = outBytes
End Function


