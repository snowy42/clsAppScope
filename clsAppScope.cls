Option Explicit

' ==============================================================================================
' clsAppScope — Scoped Suspension of Excel Application Settings
' ==============================================================================================
' Author:        Matthew Snow / Your VB Tutor
' File:          clsAppScope.cls
' Version:       1.0.0
' Created:       2025-11-01
'
' Purpose:
'   Provides a safe and concise way to temporarily modify Excel's global application settings
'   (EnableEvents, ScreenUpdating, DisplayAlerts, Calculation mode, and StatusBar text) within
'   a defined execution scope — and automatically restore the original values afterward.
'
' Problem This Solves:
'   Most VBA examples use this pattern:
'
'       Application.EnableEvents = False
'       Application.ScreenUpdating = False
'       Application.Calculation = xlCalculationManual
'       ' ... work ...
'       Application.EnableEvents = True
'       Application.ScreenUpdating = True
'       Application.Calculation = xlCalculationAutomatic
'
'   This is both **unsafe** and **incorrect** because:
'     - If an error occurs, the restore step is skipped ? Excel is left in a broken state.
'     - More importantly: this forces values to TRUE even if they were already FALSE.
'       Example:
'           - A higher-level procedure has already disabled ScreenUpdating for performance.
'           - A nested routine runs and restores ScreenUpdating to TRUE when it finishes.
'           > The nested routine **overrides the outer routine’s intention**, causing flicker
'             and potentially breaking UI logic.
'
' Why This Class Is Better:
'   - We *capture the previous state*, and only restore that state.
'   - We *only restore values we actually changed*.
'   - Restoration occurs automatically via Class_Terminate, meaning:
'       - Works on normal exit
'       - Works on errors (exception-safe)
'       - Allows proper nested usage — inner scopes cannot "undo" outer scopes.
'
'   Example:
'       With AppScopeF(sEvents Or sScreen Or sCalc, , "Processing...")
'           ' ... fast & clean code ...
'       End With   ' < original state restored properly
'
' Flags Available:
'       sEvents     > Disable events
'       sScreen     > Disable screen updating
'       sAlerts     > Disable alerts
'       sCalc       > Switch to Manual calculation (restored afterwards)
'       sStatus     > Show custom StatusBar text (restored afterwards)
'       sAll        > Apply all above
'
' Notes:
'   - If the VBA "Reset" button is used, Class_Terminate does *not* fire — call AppRestoreDefaults.
'   - sCalc defaults  to xlCalculationManual (override in AppScopeF if needed).
'   - Designed to be used via the factory:
'
'           With AppScopeF(sEvents Or sScreen Or sCalc Or sStatus, , "Working…")
'               ' work
'           End With
'
' ==============================================================================================

' ------------------------------------------------------------------------------
' Captured originals
' ------------------------------------------------------------------------------
Private mPrevEvents As Boolean
Private mPrevScreen As Boolean
Private mPrevAlerts As Boolean
Private mPrevCalc As XlCalculation
Private mPrevStatus As Variant

' Which properties we actually changed (so we only restore those)
Private mSetEvents As Boolean
Private mSetScreen As Boolean
Private mSetAlerts As Boolean
Private mSetCalc As Boolean
Private mSetStatus As Boolean

Private mDisposed As Boolean

' ------------------------------------------------------------------------------
' Bitmask flags (combine with 'Or' or '+')
' ------------------------------------------------------------------------------
Public Enum ScopeFlags
    sNone = &H0&
    sEvents = &H1&      ' Application.EnableEvents = False
    sScreen = &H2&      ' Application.ScreenUpdating = False
    sAlerts = &H4&      ' Application.DisplayAlerts = False
    sCalc = &H8&        ' Application.Calculation = <calc> (Manual by default)
    sStatus = &H10&     ' Application.StatusBar = <status text>
    sAll = &H1F&        ' events + screen + alerts + calc + status
End Enum

' ------------------------------------------------------------------------------
' Enter scope: apply selected flags.
' - calc   is only used if sCalc is present. Defaults to Manual.
' - status is only used if sStatus is present and non-empty.
' ------------------------------------------------------------------------------
Public Sub SuspendFlags( _
    ByVal flags As ScopeFlags, _
    Optional ByVal calc As XlCalculation = xlCalculationManual, _
    Optional ByVal status As String = vbNullString _
)
    On Error Resume Next

    If (flags And sEvents) <> 0 Then
        If Application.EnableEvents <> False Then
            mPrevEvents = Application.EnableEvents
            Application.EnableEvents = False
            mSetEvents = True                  ' only mark changed when we actually changed it
        End If
    End If

    If (flags And sScreen) <> 0 Then
        If Application.ScreenUpdating <> False Then
            mPrevScreen = Application.ScreenUpdating
            Application.ScreenUpdating = False
            mSetScreen = True
        End If
    End If

    If (flags And sAlerts) <> 0 Then
        If Application.DisplayAlerts <> False Then
            mPrevAlerts = Application.DisplayAlerts
            Application.DisplayAlerts = False
            mSetAlerts = True
        End If
    End If

    If (flags And sCalc) <> 0 Then
        If Application.Calculation <> calc Then
            mPrevCalc = Application.Calculation
            Application.Calculation = calc      ' default = Manual, override via AppScopeF if desired
            mSetCalc = True
        End If
    End If

    If (flags And sStatus) <> 0 Then
        If LenB(status) > 0 Then
            mPrevStatus = Application.StatusBar
            Application.StatusBar = status
            mSetStatus = True
        End If
    End If
End Sub

' ------------------------------------------------------------------------------
' Optional manual restore (usually not needed; destructor handles it)
' ------------------------------------------------------------------------------
Public Sub ResumeOriginals()
    If mDisposed Then Exit Sub
    On Error Resume Next

    If mSetEvents Then Application.EnableEvents = mPrevEvents
    If mSetScreen Then Application.ScreenUpdating = mPrevScreen
    If mSetAlerts Then Application.DisplayAlerts = mPrevAlerts
    If mSetCalc Then Application.Calculation = mPrevCalc
    If mSetStatus Then Application.StatusBar = mPrevStatus

    mDisposed = True
End Sub

' ------------------------------------------------------------------------------
' Leave scope: auto-restore on normal exit and on error unwind
' (Does NOT run if the IDE "Reset" button is used)
' ------------------------------------------------------------------------------
Private Sub Class_Terminate()
    ResumeOriginals
End Sub


