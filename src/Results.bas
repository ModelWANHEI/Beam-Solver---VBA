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

Public Sub ExportCSV(dB() As BeamSample, ProjectName As String)
    Dim FileNum As Integer
    FileNum = FreeFile
    Dim FileName As String
    Dim SaveFolder As String
    Dim ProjectFolder As String
    Dim RowString As String
    Dim Index As Long
    
    If ProjectName = "" Then
        MsgBox "Please enter a project name."
        Exit Sub
    End If
    
    SaveFolder = ThisWorkbook.Path & "\Projects"
    If Dir(SaveFolder, vbDirectory) = "" Then
        MkDir SaveFolder
    End If
    
    ProjectFolder = SaveFolder & "\" & ProjectName
    
    If Dir(ProjectFolder, vbDirectory) = "" Then
        MkDir ProjectFolder
    End If
    
    FileName = ProjectFolder & "\" & ProjectName & "Results.csv"
    Open FileName For Output As #FileNum
    
    'Header
    RowString = "x(m),Shear(kN),Moment(kNm),Stress(MPa),Slope(rad),Deflection(m)"
    Print #FileNum, RowString
    
    For Index = LBound(dB) To UBound(dB)
        RowString = dB(Index).X & "," & dB(Index).Shear & "," & dB(Index).Moment & "," & dB(Index).Stress & "," & dB(Index).Slope & "," & dB(Index).Deflection
        Print #FileNum, RowString
    Next Index
    
    Close #FileNum
End Sub

Public Function Q(ByVal S As String) As String
    Q = """" & S & """"
End Function

Public Sub ExportJSON(ProjectName As String, B As Beam, Loads() As PointLoad, Moments() As PointMoment, DLs() As DistributedLoad)
    Dim FileNum As Integer
    FileNum = FreeFile
    Dim FileName As String
    Dim SaveFolder As String
    Dim ProjectFolder As String
    Dim RowString As String
    Dim Index As Long
    
    If ProjectName = "" Then
        MsgBox "Please enter a project name."
        Exit Sub
    End If
    
    SaveFolder = ThisWorkbook.Path & "\Projects"
    If Dir(SaveFolder, vbDirectory) = "" Then
        MkDir SaveFolder
    End If
    
    ProjectFolder = SaveFolder & "\" & ProjectName
    
    If Dir(ProjectFolder, vbDirectory) = "" Then
        MkDir ProjectFolder
    End If
    
    FileName = ProjectFolder & "\" & ProjectName & "Config.json"
    Open FileName For Output As #FileNum
    
    Print #FileNum, "{" 'This is the starting JSON curl
    
    'Project Name
    RowString = "   " & Q("ProjectName") & ": " & Q(ProjectName) & ","
    Print #FileNum, RowString
    
    Print #FileNum, "" 'Separator
    
    'Beam Specifications
    Print #FileNum, "   " & Q("Beam") & ": {"
    RowString = "    " & Q("Length_m") & ": " & B.Length & ","
    Print #FileNum, RowString
    RowString = "    " & Q("YoungsModulus_GPa") & ": " & B.E & ","
    Print #FileNum, RowString
    RowString = "    " & Q("Width_m") & ": " & B.B & ","
    Print #FileNum, RowString
    RowString = "    " & Q("Height_m") & ": " & B.H & ","
    Print #FileNum, RowString
    RowString = "    " & Q("MomentOfInertia_m_pow_4") & ": " & B.i & ","
    Print #FileNum, RowString
    RowString = "    " & Q("Resolution_m") & ": " & B.Resolution
    Print #FileNum, RowString
    Print #FileNum, "   },"
    
    Print #FileNum, ""
    
    'Point Loads
    Print #FileNum, "   " & Q("PointLoads") & ": ["
        For Index = LBound(Loads) To UBound(Loads)
            Print #FileNum, "    {"
            RowString = "    " & Q("DistanceFromA_m") & ": " & Loads(Index).PositionFromA & ","
            Print #FileNum, RowString
            RowString = "    " & Q("Magnitude_kN") & ": " & Loads(Index).Magnitude
            Print #FileNum, RowString
            If Index <> UBound(Loads) Then
                Print #FileNum, "    },"
            Else
                Print #FileNum, "    }"
            End If
        Next Index
    Print #FileNum, "   ],"
    
    Print #FileNum, ""
    
    'Applied Moments
    Print #FileNum, "   " & Q("AppliedMoments") & ": ["
        For Index = LBound(Moments) To UBound(Moments)
            Print #FileNum, "    {"
            RowString = "    " & Q("DistanceFromA_m") & ": " & Moments(Index).PositionFromA & ","
            Print #FileNum, RowString
            RowString = "    " & Q("Magnitude_kNm") & ": " & Moments(Index).Magnitude
            Print #FileNum, RowString
            If Index <> UBound(Moments) Then
                Print #FileNum, "    },"
            Else
                Print #FileNum, "    }"
            End If
        Next Index
    Print #FileNum, "   ],"
    
    Print #FileNum, ""
    
    'Dist Loads
    Print #FileNum, "   " & Q("DistributedLoads") & ": ["
        For Index = LBound(DLs) To UBound(DLs)
            Print #FileNum, "    {"
            RowString = "    " & Q("StartDistanceFromA_m") & ": " & DLs(Index).StartPosition & ","
            Print #FileNum, RowString
            RowString = "    " & Q("EndDistanceFromA_m") & ": " & DLs(Index).EndPosition & ","
            Print #FileNum, RowString
            RowString = "    " & Q("StartIntensity_kN/m") & ": " & DLs(Index).StartIntensity & ","
            Print #FileNum, RowString
            RowString = "    " & Q("EndIntensity_kN/m") & ": " & DLs(Index).EndIntensity
            Print #FileNum, RowString
            If Index <> UBound(DLs) Then
                Print #FileNum, "    },"
            Else
                Print #FileNum, "    }"
            End If
        Next Index
    Print #FileNum, "   ]"
    
    Print #FileNum, "}" 'Closing curl
    
    Close #FileNum
End Sub

