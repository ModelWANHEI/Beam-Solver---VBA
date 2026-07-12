Attribute VB_Name = "solvers"
Public Function SolveReactionForces(B As Beam, Loads() As PointLoad, Moments() As PointMoment) As Reactions
    Dim R As Reactions
    Dim SumFy As Double
    Dim SumMA As Double
    Dim Counter As Long
    SumFy = 0
    SumMA = 0
    
    'Ay + By = -Sum Forces
    'By*-L + Pi*-xi + Pj*-xj + ... Pn*-xn = 0; Pi*-xi + Pj*-xj + ... Pn*-xn = By*L
    
    For Counter = LBound(Loads) To UBound(Loads)
      SumFy = SumFy + Loads(Counter).Magnitude
      SumMA = SumMA - (Loads(Counter).Magnitude * Loads(Counter).PositionFromA)
    Next Counter
    
    For Counter = LBound(Moments) To UBound(Moments)
        SumMA = SumMA + Moments(Counter).Magnitude
    Next Counter
    
    R.By = SumMA / B.Length
    R.Ay = -(SumFy + R.By)
    
    SolveReactionForces = R
    
End Function

Public Function ItemizeInternalReactions(B As Beam, Loads() As PointLoad, Moments() As PointMoment, R As Reactions) As BeamSample()
    Dim BeamResults() As BeamSample
    Dim CumulativeShear As Double
    Dim CumulativeMoment As Double
    Dim NumPoints As Long
    Dim i As Long
    Dim PointIndexer As Long
    
    NumPoints = CLng(B.Length / B.Resolution)
    ReDim BeamResults(0 To NumPoints)
    
    For i = 0 To NumPoints
        CumulativeShear = R.Ay
        BeamResults(i).X = i * B.Resolution
        CumulativeMoment = CumulativeShear * BeamResults(i).X
        
        For PointIndexer = LBound(Loads) To UBound(Loads)
            If Loads(PointIndexer).PositionFromA <= BeamResults(i).X Then
                CumulativeShear = CumulativeShear + Loads(PointIndexer).Magnitude
                'They should add += and ++ to VBA.
                CumulativeMoment = CumulativeMoment + ((BeamResults(i).X - Loads(PointIndexer).PositionFromA) * (Loads(PointIndexer).Magnitude))
            End If
        Next PointIndexer
        
        For PointIndexer = LBound(Moments) To UBound(Moments)
            If Moments(PointIndexer).PositionFromA <= BeamResults(i).X Then
                CumulativeMoment = CumulativeMoment + Moments(PointIndexer).Magnitude
            End If
        Next PointIndexer
        
        If i = NumPoints Then
            CumulativeShear = CumulativeShear + R.By
        End If
        
        BeamResults(i).Shear = -CumulativeShear
        BeamResults(i).Moment = -CumulativeMoment
        
        If Abs(BeamResults(i).Shear) < Tolerance Then
            BeamResults(i).Shear = 0
        End If
        
        If Abs(BeamResults(i).Moment) < Tolerance Then
            BeamResults(i).Moment = 0
        End If
          
    Next i
    
    ItemizeInternalReactions = BeamResults
     
End Function

Public Function PrettyForce(F As Double) As String
    If Abs(F) < Tolerance Then
        PrettyForce = "0"
    Else
        PrettyForce = Format(F, "0.000")
    End If
End Function
