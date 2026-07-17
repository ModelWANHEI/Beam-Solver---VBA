VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SummaryForm 
   Caption         =   "Summary"
   ClientHeight    =   8988.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5124
   OleObjectBlob   =   "SummaryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SummaryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub LoadSummary(POI() As PointOfInterest, ProjectName As String)
    Dim Index As Long
    
    ResultsBox.CLEAR
    
    ResultsBox.ColumnCount = 4
    ResultsBox.ColumnWidths = "80 pt;70 pt;70 pt;90 pt"
    
    ResultsBox.AddItem
    ResultsBox.List(0, 0) = "Variable"
    ResultsBox.List(0, 1) = "Extrema"
    ResultsBox.List(0, 2) = "Position (m)"
    ResultsBox.List(0, 3) = "Magnitude"
    
    For Index = 0 To POI(SHEAR_COL).NumberOfPoints(MINIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Shear"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Minimum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(SHEAR_COL).MinimaPositionFromA(Index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(SHEAR_COL).Magnitude(MINIMA), "0.000")
    Next Index
    
    For Index = 0 To POI(SHEAR_COL).NumberOfPoints(MAXIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Shear"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Maximum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(SHEAR_COL).MaximaPositionFromA(Index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(SHEAR_COL).Magnitude(MAXIMA), "0.000")
    Next Index
    
    For Index = 0 To POI(MOMENT_COL).NumberOfPoints(MINIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Moment"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Minimum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(MOMENT_COL).MinimaPositionFromA(Index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(MOMENT_COL).Magnitude(MINIMA), "0.000")
    Next Index
    
    For Index = 0 To POI(MOMENT_COL).NumberOfPoints(MAXIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Moment"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Maximum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(MOMENT_COL).MaximaPositionFromA(Index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(MOMENT_COL).Magnitude(MAXIMA), "0.000")
    Next Index
    
    For Index = 0 To POI(STRESS_COL).NumberOfPoints(MINIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Stress"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Minimum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(STRESS_COL).MinimaPositionFromA(Index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(STRESS_COL).Magnitude(MINIMA), "0.000")
    Next Index
    
    For Index = 0 To POI(STRESS_COL).NumberOfPoints(MAXIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Stress"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Maximum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(STRESS_COL).MaximaPositionFromA(Index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(STRESS_COL).Magnitude(MAXIMA), "0.000")
    Next Index
    
    For Index = 0 To POI(SLOPE_COL).NumberOfPoints(MINIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Slope"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Minimum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(SLOPE_COL).MinimaPositionFromA(Index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(SLOPE_COL).Magnitude(MINIMA), "0.000")
    Next Index
    
    For Index = 0 To POI(SLOPE_COL).NumberOfPoints(MAXIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Slope"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Maximum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(SLOPE_COL).MaximaPositionFromA(Index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(SLOPE_COL).Magnitude(MAXIMA), "0.000")
    Next Index
    
    For Index = 0 To POI(DEFLECTION_COL).NumberOfPoints(MINIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Deflection"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Minimum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(DEFLECTION_COL).MinimaPositionFromA(Index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(DEFLECTION_COL).Magnitude(MINIMA), "0.000")
    Next Index
    
    For Index = 0 To POI(DEFLECTION_COL).NumberOfPoints(MAXIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Deflection"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Maximum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(DEFLECTION_COL).MaximaPositionFromA(Index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(DEFLECTION_COL).Magnitude(MAXIMA), "0.000")
    Next Index
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Dim FileName As String
    Dim SaveFolder As String
    Dim ProjectFolder As String
    Dim RowString As String
    
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
    
    FileName = ProjectFolder & "\" & ProjectName & "Summary.csv"
    Open FileName For Output As #FileNum
    
    For Index = 0 To ResultsBox.ListCount - 1
        RowString = ResultsBox.List(Index, 0) & "," & ResultsBox.List(Index, 1) & "," & ResultsBox.List(Index, 2) & "," & ResultsBox.List(Index, 3)
        Print #FileNum, RowString
    Next Index
    
    Close #FileNum
    
End Sub
