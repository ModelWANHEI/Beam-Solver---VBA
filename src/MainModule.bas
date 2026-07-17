Attribute VB_Name = "MainModule"
Option Explicit

'Thank you for viewing my work. Last maintained 2026.07.10. If the GitHub says otherwise, I forgot to change this.
'=================================================================================================================
'Conventions
'SI Units
'Clockwise Moment Positive
'Up, Right Forces Positive
'If Shear is Up on the rightmost side of the slice, it's positive.

       '^
'--|Beam|------ Positive Shear
  'v
  
'Maxima is defined as positive extrema (most positive value)
'Minima is defined as negative extrema (most negative value)

'If you dunno what that means, my bad. But I get it, so YOU are gonna have to live with it.
'=================================================================================================================

Public Sub main(B As Beam, Loads() As PointLoad, Moments() As PointMoment, ProjectName As String)

    Dim R As Reactions
    Dim dxPoints As Long
    Dim POI() As PointOfInterest
    
    'Solve Reactions
    R = SolveReactionForces(B, Loads, Moments)
    
    'Generate Shear, Generate Moment
    dxPoints = CLng(B.Length / B.Resolution)
    ReDim dBeam(0 To dxPoints)
    dBeam = ItemizeInternalReactions(B, Loads, Moments, R)
    
    'Generate Stress
    Call CalculateStressReactions(B, dBeam)
    
    'Generate Deflection
    Call CalculateDeflection(B, dBeam)
    
    SCROLL = SHEAR_COL
    Call MakeScatterPlot(dBeam, X_COL, SCROLL)
    
    POI = SummarizeGlobality(dBeam)
    
    Call SummaryForm.LoadSummary(POI, ProjectName)
    
    Call ExportCSV(dBeam, ProjectName)
    
End Sub
