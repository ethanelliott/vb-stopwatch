Attribute VB_Name = "moduleMain"
'==================================
'        Stopwatch Program
'     (c) Ethan Elliott 2014
'==================================
Option Explicit
Public curTimStr As String      'Current time represented as a string
Public lapnum As Integer        'current lap number
Public timelap As String        'current time matching the lap

'Function   : recordLap()
'Input      : None
'Return     : None
'Description: Records the current time to the lap form
Public Sub recordLap()
    frmLap.txtLap.Text = frmLap.txtLap.Text & "Lap " & lapnum & " = " & timelap & vbNewLine
End Sub
