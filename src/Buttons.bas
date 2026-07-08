Attribute VB_Name = "Buttons"
Public Sub ScrollForward()

    On Error Resume Next
    Dim UB As Long
    UB = UBound(dBeam)

    If Err.Number <> 0 Then
        MsgBox "Please solve a beam first."
        Exit Sub
    End If
    On Error GoTo 0
    
    SCROLL = SCROLL + 1
    If SCROLL > DEFLECTION_COL Then
        SCROLL = SHEAR_COL
    End If
    
    
    Call MakeScatterPlot(dBeam, X_COL, SCROLL)
    
End Sub

Public Sub ScrollBackward()
    On Error Resume Next
    Dim UB As Long
    UB = UBound(dBeam)

    If Err.Number <> 0 Then
        MsgBox "Please solve a beam first."
        Exit Sub
    End If
    On Error GoTo 0
    
    SCROLL = SCROLL - 1
    If SCROLL < SHEAR_COL Then
        SCROLL = DEFLECTION_COL
    End If
    
     Call MakeScatterPlot(dBeam, X_COL, SCROLL)
End Sub

Public Sub LaunchBeamSolver()
    BeamANHEILyzerV1.Show vbModeless
End Sub
