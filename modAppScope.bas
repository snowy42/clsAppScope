Option Explicit

' ==============================================================================================
' modAppScope — Factory and Recovery Helpers for clsAppScope
' Author:        Matthew Snow / Your VB Tutor
' File:          modAppScope.bas
' Version:       1.0.0
'
' Usage:
'   With AppScopeF(sEvents Or sScreen Or sCalc Or sStatus, , "Working…")
'       ' ... your code ...
'   End With  ' ? all settings restored to their original values
'
' Recovery (for IDE Stop/Reset):
'   Call AppRestoreDefaults
' ==============================================================================================

' Factory for concise syntax
Public Function AppScopeF( _
    ByVal flags As ScopeFlags, _
    Optional ByVal calc As XlCalculation = xlCalculationManual, _
    Optional ByVal status As String = vbNullString _
) As clsAppScope
    Dim s As New clsAppScope
    s.SuspendFlags flags, calc, status
    Set AppScopeF = s
End Function

' Panic reset if the IDE "Reset" button was pressed mid-run (Class_Terminate won’t fire)
Public Sub AppRestoreDefaults()
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
End Sub

' ------------------------------------------------------------------------------
' Quick examples
' ------------------------------------------------------------------------------
' Sub Example_EventsOnly()
'     With AppScopeF(sEvents)
'         ' events disabled here only
'         [A1].Value = "Hello"
'     End With
' End Sub
'
' Sub Example_AllCommon()
'     With AppScopeF(sAll, , "Updating data…")
'         Range("A1:B100000").Value = 1
'     End With
' End Sub
'
' Sub Example_Nested()
'     With AppScopeF(sEvents Or sScreen, , "Outer…")
'         ' outer scope
'         With AppScopeF(sCalc, Status:="Inner…")
'             ' inner scope suspends calc only
'         End With
'         ' outer still active here
'     End With
' End Sub


