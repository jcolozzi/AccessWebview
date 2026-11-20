Attribute VB_Name = "clsWv2AccessBridge"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

' ============================================================================
' clsWv2AccessBridge - Direct JS <-> Access VBA/DAO Host Object Bridge
'
' Registered via:
'   m_pane.View_AddHostObjectToScript "vba", m_bridge
'
' Callable from JavaScript as:
'   const json = await window.chrome.webview.hostObjects.vba.GetTasksJson();
'   const syncJson = window.chrome.webview.hostObjects.sync.vba.GetHostInfoJson();
' ============================================================================

Public Function Ping(ByVal echoText As String) As String
    Ping = "Access Pong: " & echoText
End Function

Public Function GetHostInfoJson() As String
    Dim arch As String
    #If Win64 Then
        arch = "x64"
    #Else
        arch = "x86"
    #End If

    GetHostInfoJson = "{" & _
        """host"":""Microsoft Access""," & _
        """version"":""" & Wv2Json.JsonEscape(Application.Version & "." & CStr(Application.Build)) & """," & _
        """arch"":""" & arch & """," & _
        """dbPath"":""" & Wv2Json.JsonEscape(CurrentProject.FullName) & """" & _
        "}"
End Function

Public Function GetTasksJson() As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sb As String
    Dim isFirst As Boolean

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT TaskID, TaskTitle, CategoryName, PriorityLevel, IsCompleted " & _
        "FROM tblWv2Tasks ORDER BY TaskID;", dbOpenSnapshot)

    sb = "["
    isFirst = True
    Do While Not rs.EOF
        If Not isFirst Then sb = sb & ","
        isFirst = False
        sb = sb & "{" & _
            """id"":" & CStr(Nz(rs!taskId, 0)) & "," & _
            """title"":""" & Wv2Json.JsonEscape(CStr(Nz(rs!taskTitle, ""))) & """," & _
            """category"":""" & Wv2Json.JsonEscape(CStr(Nz(rs!categoryName, ""))) & """," & _
            """priority"":""" & Wv2Json.JsonEscape(CStr(Nz(rs!priorityLevel, "Medium"))) & """," & _
            """done"":" & IIf(CBool(Nz(rs!IsCompleted, False)), "true", "false") & _
            "}"
        rs.MoveNext
    Loop
    sb = sb & "]"

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    GetTasksJson = sb
    Exit Function

ErrHandler:
    Dim errDesc As String
    errDesc = Err.Description
    On Error Resume Next
    If Not (rs Is Nothing) Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    On Error GoTo 0
    GetTasksJson = "{""error"":""" & Wv2Json.JsonEscape(errDesc) & """}"
End Function

Public Function AddTask(ByVal taskTitle As String, ByVal categoryName As String, ByVal priorityLevel As String) As String
    On Error GoTo ErrHandler
    Dim cleanTitle As String
    Dim cleanCat As String
    Dim cleanPrio As String

    cleanTitle = Trim$(taskTitle)
    If Len(cleanTitle) = 0 Then
        AddTask = "{""ok"":false,""error"":""Task title is required""}"
        Exit Function
    End If

    cleanCat = Trim$(categoryName)
    If Len(cleanCat) = 0 Then cleanCat = "General"

    cleanPrio = Trim$(priorityLevel)
    If Len(cleanPrio) = 0 Then cleanPrio = "Medium"

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Set db = CurrentDb
    Set qdf = db.CreateQueryDef("", _
        "PARAMETERS pTitle Text(255), pCat Text(100), pPrio Text(50); " & _
        "INSERT INTO tblWv2Tasks (TaskTitle, CategoryName, PriorityLevel, IsCompleted) " & _
        "VALUES ([pTitle], [pCat], [pPrio], False);")
    qdf.Parameters("pTitle").value = cleanTitle
    qdf.Parameters("pCat").value = cleanCat
    qdf.Parameters("pPrio").value = cleanPrio
    qdf.Execute dbFailOnError

    Set qdf = Nothing
    Set db = Nothing
    AddTask = GetTasksJson()
    Exit Function

ErrHandler:
    Set qdf = Nothing
    Set db = Nothing
    AddTask = "{""ok"":false,""error"":""" & Wv2Json.JsonEscape(Err.Description) & """}"
End Function

Public Function ToggleTask(ByVal taskId As Long) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Set db = CurrentDb
    Set qdf = db.CreateQueryDef("", _
        "PARAMETERS pId Long; " & _
        "UPDATE tblWv2Tasks SET IsCompleted = Not IsCompleted WHERE TaskID = [pId];")
    qdf.Parameters("pId").value = taskId
    qdf.Execute dbFailOnError

    Set qdf = Nothing
    Set db = Nothing
    ToggleTask = GetTasksJson()
    Exit Function

ErrHandler:
    Set qdf = Nothing
    Set db = Nothing
    ToggleTask = "{""ok"":false,""error"":""" & Wv2Json.JsonEscape(Err.Description) & """}"
End Function

