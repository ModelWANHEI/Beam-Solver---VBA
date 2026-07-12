Attribute VB_Name = "Results"
Public Sub MakeScatterPlot(dB() As BeamSample, X As Long, Y As Long)
    
    Dim Index As Long
    Dim Chart As ChartObject
    Dim WS As Worksheet
    Dim LastRow As Long
    Dim c As ChartObject
    
    LastRow = 25 + UBound(dB)
    
    On Error Resume Next
    Set WS = Worksheets("Results")
    On Error GoTo 0

    If WS Is Nothing Then
        Set WS = Worksheets.Add
        WS.Name = "Results"
    End If
    
    If X < 1 Or Y < 1 Then
        MsgBox "Go home.", vbOKOnly, "Please be a responsible user."
        Exit Sub
    End If
    
    WS.Cells.ClearContents
     
    WS.Cells(24, X_COL).Value = "x(m)"
    WS.Cells(24, SHEAR_COL).Value = "Shear(kN)"
    WS.Cells(24, MOMENT_COL).Value = "Moment(kNm)"
    WS.Cells(24, STRESS_COL).Value = "Stress (MPa)"
    WS.Cells(24, SLOPE_COL).Value = "Slope"
    WS.Cells(24, DEFLECTION_COL).Value = "Deflection (m)"
    
    For Index = LBound(dB) To UBound(dB)
        WS.Cells(Index + 25, X_COL).Value = dB(Index).X
        'X should only ever be X.
        WS.Cells(Index + 25, SHEAR_COL).Value = dB(Index).Shear
        WS.Cells(Index + 25, MOMENT_COL).Value = dB(Index).Moment
        WS.Cells(Index + 25, STRESS_COL).Value = dB(Index).Stress
        WS.Cells(Index + 25, SLOPE_COL).Value = dB(Index).Slope
        WS.Cells(Index + 25, DEFLECTION_COL).Value = dB(Index).Deflection
    Next Index
    
    For Each c In WS.ChartObjects
        c.Delete
    Next c
    
    Set Chart = WS.ChartObjects.Add(Left:=100, Top:=50, Width:=400, Height:=250)
    
    With Chart.Chart
        .ChartType = xlXYScatterLines
        .HasTitle = True
        Select Case Y
            Case SHEAR_COL
                Debug.Print "Writing Shear"
                .ChartTitle.Text = "Shear Along Beam"
            Case MOMENT_COL
                .ChartTitle.Text = "Moment Along Beam"
            Case STRESS_COL
                .ChartTitle.Text = "Stress Along Beam"
            Case SLOPE_COL
                .ChartTitle.Text = "Slope Along Beam"
            Case DEFLECTION_COL
                .ChartTitle.Text = "Deflection Along Beam"
        End Select
        
        With .SeriesCollection.NewSeries
            Select Case Y
                Case SHEAR_COL
                    .Name = "Shear (kN)"
                Case MOMENT_COL
                    .Name = "Moment (kNm)"
                Case STRESS_COL
                    .Name = "Stress (MPa)"
                Case SLOPE_COL
                    .Name = "Slope"
                Case DEFLECTION_COL
                    .Name = "Deflection (m)"
            End Select
            .XValues = WS.Range(WS.Cells(25, X), WS.Cells(LastRow, X))
            .Values = WS.Range(WS.Cells(25, Y), WS.Cells(LastRow, Y))
        End With
    End With
    
End Sub

