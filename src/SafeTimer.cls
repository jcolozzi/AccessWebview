Attribute VB_Name = "SafeTimer"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

#Const Windows = (Mac = 0)
#If Windows Then
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As LongPtr) As LongPtr

    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal uIDEvent As LongPtr) As Long

    Private Declare PtrSafe Function CreateWindowExW Lib "user32" ( _
        ByVal dwExStyle As Long, _
        ByVal lpClassName As LongPtr, _
        ByVal lpWindowName As LongPtr, _
        ByVal dwStyle As Long, _
        ByVal X As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hWndParent As LongPtr, _
        ByVal hMenu As LongPtr, _
        ByVal hInstance As LongPtr, _
        ByVal lpParam As LongPtr) As LongPtr

    Private Declare PtrSafe Function DestroyWindow Lib "user32" ( _
        ByVal hwnd As LongPtr) As Long

    Private Const HWND_MESSAGE As LongPtr = -3
#End If

Private m_hwnd As LongPtr
Private m_uElapse As Long
Private m_data As Variant

Public Event TimerCall(ByVal wTime As Long, ByRef varData As Variant)

' IMPORTANT: Do not add any Public/Friend/Private Sub/Function/Property BEFORE StartTimer and TimerProc!
' LibTimers.GetTimerProc calls SafeTimer.TimerProc via fixed vtable slot 8 (7=StartTimer, 8=TimerProc).
Public Sub StartTimer(ByVal uElapse As Long, Optional ByRef varData As Variant)
    If uElapse = m_uElapse Then Exit Sub
#If Windows Then
    If m_hwnd <> 0 Then
        KillTimer m_hwnd, ObjPtr(Me)
        DestroyWindow m_hwnd
        m_hwnd = 0
    End If
#End If
    m_uElapse = 0
    If uElapse = 0 Then Exit Sub

    m_uElapse = uElapse
    If IsObject(varData) Then Set m_data = varData Else m_data = varData
#If Windows Then
    m_hwnd = CreateWindowExW(0, StrPtr("STATIC"), 0, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, 0, 0)
    If m_hwnd <> 0 Then
        SetTimer m_hwnd, ObjPtr(Me), uElapse, GetTimerProc(Me)
    End If
#End If
End Sub

'Do not move. This must be the 8th position in virtual table (0..6 IDispatch, 7 StartTimer, 8 TimerProc)
Private Sub TimerProc(ByVal wMsg As Long, ByVal nIDEvent As LongPtr, ByVal wTime As Long)
    On Error Resume Next
    RaiseEvent TimerCall(wTime, m_data)
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
#If Windows Then
    If m_hwnd <> 0 Then
        KillTimer m_hwnd, ObjPtr(Me)
        DestroyWindow m_hwnd
        m_hwnd = 0
    End If
#End If
End Sub


