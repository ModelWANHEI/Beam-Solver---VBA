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
Public Sub LoadSummary(POI() As PointOfInterest)
    Dim index As Long
    
    ResultsBox.CLEAR
    
    ResultsBox.ColumnCount = 4
    ResultsBox.ColumnWidths = "80 pt;70 pt;70 pt;90 pt"
    
    ResultsBox.AddItem
    ResultsBox.List(0, 0) = "Variable"
    ResultsBox.List(0, 1) = "Extrema"
    ResultsBox.List(0, 2) = "Position (m)"
    ResultsBox.List(0, 3) = "Magnitude"
    
    For index = 0 To POI(SHEAR_COL).NumberOfPoints(MINIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Shear"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Minimum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(SHEAR_COL).MinimaPositionFromA(index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(SHEAR_COL).Magnitude(MINIMA), "0.000")
    Next index
    
    For index = 0 To POI(SHEAR_COL).NumberOfPoints(MAXIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Shear"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Maximum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(SHEAR_COL).MaximaPositionFromA(index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(SHEAR_COL).Magnitude(MAXIMA), "0.000")
    Next index
    
    For index = 0 To POI(MOMENT_COL).NumberOfPoints(MINIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Moment"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Minimum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(MOMENT_COL).MinimaPositionFromA(index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(MOMENT_COL).Magnitude(MINIMA), "0.000")
    Next index
    
    For index = 0 To POI(MOMENT_COL).NumberOfPoints(MAXIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Moment"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Maximum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(MOMENT_COL).MaximaPositionFromA(index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(MOMENT_COL).Magnitude(MAXIMA), "0.000")
    Next index
    
    For index = 0 To POI(STRESS_COL).NumberOfPoints(MINIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Stress"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Minimum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(STRESS_COL).MinimaPositionFromA(index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(STRESS_COL).Magnitude(MINIMA), "0.000")
    Next index
    
    For index = 0 To POI(STRESS_COL).NumberOfPoints(MAXIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Stress"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Maximum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(STRESS_COL).MaximaPositionFromA(index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(STRESS_COL).Magnitude(MAXIMA), "0.000")
    Next index
    
    For index = 0 To POI(SLOPE_COL).NumberOfPoints(MINIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Slope"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Minimum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(SLOPE_COL).MinimaPositionFromA(index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(SLOPE_COL).Magnitude(MINIMA), "0.000")
    Next index
    
    For index = 0 To POI(SLOPE_COL).NumberOfPoints(MAXIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Slope"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Maximum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(SLOPE_COL).MaximaPositionFromA(index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(SLOPE_COL).Magnitude(MAXIMA), "0.000")
    Next index
    
    For index = 0 To POI(DEFLECTION_COL).NumberOfPoints(MINIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Deflection"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Minimum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(DEFLECTION_COL).MinimaPositionFromA(index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(DEFLECTION_COL).Magnitude(MINIMA), "0.000")
    Next index
    
    For index = 0 To POI(DEFLECTION_COL).NumberOfPoints(MAXIMA) - 1
        ResultsBox.AddItem
        ResultsBox.List(ResultsBox.ListCount - 1, 0) = "Deflection"
        ResultsBox.List(ResultsBox.ListCount - 1, 1) = "Maximum"
        ResultsBox.List(ResultsBox.ListCount - 1, 2) = Format(POI(DEFLECTION_COL).MaximaPositionFromA(index), "0.000")
        ResultsBox.List(ResultsBox.ListCount - 1, 3) = Format(POI(DEFLECTION_COL).Magnitude(MAXIMA), "0.000")
    Next index
    
End Sub
