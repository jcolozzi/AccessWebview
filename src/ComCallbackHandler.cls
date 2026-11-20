Attribute VB_Name = "ComCallbackHandler"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' ============================================================================
' ComCallbackHandler - Dual x86/x64 COM Event & Completion Callback Receiver
'
' IMPORTANT: Handler_Invoke MUST remain the FIRST member in this class so that
' its vtable index is slot 7 (immediately after IUnknown 0..2 + IDispatch 3..6)
' on both 32-bit (x86) and 64-bit (x64) VBA!
' ============================================================================

Private m_kind       As HandlerKind
Private m_owner      As Object
Private m_pSlot      As LongPtr
Private m_refCount   As Long
Private m_callbackId As Long

Public Function Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal arg1 As LongPtr, _
    ByVal arg2 As LongPtr) As Long

    Select Case m_kind
        Case HK_EnvironmentCompleted
            Dim errorCode As Long
            errorCode = LongPtrLowDword(arg1)
            If Not (m_owner Is Nothing) Then
                m_owner.OnEnvironmentCompleted errorCode, arg2
            End If

        Case HK_ControllerCompleted
            Dim errorCode2 As Long
            errorCode2 = LongPtrLowDword(arg1)
            If Not (m_owner Is Nothing) Then
                m_owner.OnControllerCompleted errorCode2, arg2
            End If

        Case HK_NavigationStarting
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnNavigationStarting arg1, arg2
            End If

        Case HK_NavigationCompleted
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnNavigationCompleted arg1, arg2
            End If

        Case HK_WebMessageReceived
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnWebMessageReceived arg1, arg2
            End If

        Case HK_DocumentTitleChanged
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnDocumentTitleChanged arg1, arg2
            End If

        Case HK_NewWindowRequested
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnNewWindowRequested arg1, arg2
            End If

        Case HK_HistoryChanged
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnHistoryChanged arg1, arg2
            End If

        Case HK_DOMContentLoaded
            If Not (m_owner Is Nothing) Then
                m_owner.View2_OnDOMContentLoaded arg1, arg2
            End If

        Case HK_WebResourceRequested
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnWebResourceRequested arg1, arg2
            End If

        Case HK_WebResourceResponseReceived
            If Not (m_owner Is Nothing) Then
                m_owner.View2_OnWebResourceResponseReceived arg1, arg2
            End If

        Case HK_ResponseGetContent
            Dim errorCode4 As Long
            errorCode4 = LongPtrLowDword(arg1)
            If Not (m_owner Is Nothing) Then
                m_owner.OnResponseContentCompleted m_callbackId, errorCode4, arg2
            End If

        Case HK_ExecuteScriptCompleted
            Dim errorCode3 As Long
            errorCode3 = LongPtrLowDword(arg1)
            Dim resultStr As String
            If arg2 <> 0 Then
                resultStr = PtrToString(arg2)
            End If
            If Not (m_owner Is Nothing) Then
                m_owner.OnExecuteScriptCompleted m_callbackId, errorCode3, resultStr
            End If

        Case Else
            Debug.Print "ComCallbackHandler.Invoke[unknown kind=" & m_kind & "]"
    End Select

    Handler_Invoke = 0
End Function

Private Function LongPtrLowDword(ByVal v As LongPtr) As Long
#If Win64 Then
    Dim u As LongLong
    u = v And &HFFFFFFFF^
    If u >= &H80000000^ Then
        LongPtrLowDword = CLng(u - &H100000000^)
    Else
        LongPtrLowDword = CLng(u)
    End If
#Else
    LongPtrLowDword = v
#End If
End Function

Friend Sub Init( _
    ByVal kind As HandlerKind, _
    ByVal owner As Object, _
    ByVal pSlot As LongPtr)

    m_kind = kind
    Set m_owner = owner
    m_pSlot = pSlot
    m_refCount = 0
End Sub

Friend Property Get RefCount() As Long
    RefCount = m_refCount
End Property

Friend Property Let RefCount(ByVal v As Long)
    m_refCount = v
End Property

Friend Property Get Slot() As LongPtr
    Slot = m_pSlot
End Property

Friend Property Get kind() As HandlerKind
    kind = m_kind
End Property

Friend Property Get callbackId() As Long
    callbackId = m_callbackId
End Property

Friend Property Let callbackId(ByVal v As Long)
    m_callbackId = v
End Property

Friend Sub ClearOwner()
    Set m_owner = Nothing
End Sub
