Attribute VB_Name = "mdNineButton"
'=========================================================================
'
' Nine Patch PNGs for VB6 (c) 2018 by wqweto@gmail.com
' FireOnceTimer/PushParamThunk (c) M. Curland
'
' mdNineButton.bas -- fire-once timers and redirectors
'
'=========================================================================
Option Explicit

#Const ImplUseShared = NPPNG_USE_SHARED <> 0

#If Not ImplUseShared Then

'==============================================================================
' API
'==============================================================================

'--- for VirtualProtect
Private Const PAGE_EXECUTE_READWRITE            As Long = &H40

Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private Type ThunkBytes
    Thunk(5)            As Long
End Type

Public Type PushParamThunk
    pfn                 As Long
    Code                As ThunkBytes
End Type

Public Type FireOnceTimerData
    TimerID             As Long
    TimerProcThunkData  As PushParamThunk
    TimerProcThunkThis  As PushParamThunk
End Type

'==============================================================================
' Functions
'==============================================================================

'= push-param thunk ======================================================

Public Sub InitPushParamThunk(Thunk As PushParamThunk, ByVal ParamValue As Long, ByVal pfnDest As Long)
'push [esp]
'mov eax, 16h // Dummy value for parameter value
'mov [esp + 4], eax
'nop // Adjustment so the next long is nicely aligned
'nop
'nop
'mov eax, 1234h // Dummy value for function
'jmp eax
'nop
'nop
    Dim dwDummy         As Long

    With Thunk.Code
        .Thunk(0) = &HB82434FF
        .Thunk(1) = ParamValue
        .Thunk(2) = &H4244489
        .Thunk(3) = &HB8909090
        .Thunk(4) = pfnDest
        .Thunk(5) = &H9090E0FF
        Call VirtualProtect(.Thunk(0), Len(Thunk), PAGE_EXECUTE_READWRITE, dwDummy)
    End With
    Thunk.pfn = VarPtr(Thunk.Code)
End Sub

'= fire-once timers ======================================================

Public Sub InitFireOnceTimer(Data As FireOnceTimerData, ByVal ThisPtr As Long, ByVal pfnRedirect As Long, Optional ByVal Delay As Long)
    With Data
        InitPushParamThunk .TimerProcThunkData, VarPtr(Data), pfnRedirect
        InitPushParamThunk .TimerProcThunkThis, ThisPtr, .TimerProcThunkData.pfn
        .TimerID = SetTimer(0, 0, Delay, .TimerProcThunkThis.pfn)
    End With
End Sub

Public Sub TerminateFireOnceTimer(Data As FireOnceTimerData)
    With Data
        If .TimerID <> 0 Then
            Call KillTimer(0, .TimerID)
            .TimerID = 0
        End If
    End With
End Sub

#End If ' ImplUseShared

'==============================================================================
' Redirectors
'==============================================================================

Public Sub RedirectNineButtonTimerProc( _
            Data As FireOnceTimerData, _
            ByVal This As ctxNineButton, _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal idEvent As Long, _
            ByVal dwTime As Long)
    #If hWnd And wMsg And dwTime Then '--- touch
    #End If
    Data.TimerID = idEvent
    TerminateFireOnceTimer Data
    This.frTimer
End Sub
