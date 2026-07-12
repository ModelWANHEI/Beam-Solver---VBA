Attribute VB_Name = "constructors"
Public Function ConstructBeam(Length As Double, E As Double, W As Double, H As Double, Res As Double) As Beam
    Dim B As Beam
    
    If Length <= 0 Then
        Err.Raise vbObjectError + 1, , "Length must be larger than zero."
        Exit Function
    End If
    
    If E <= 0 Then
        Err.Raise vbObjectError + 1, , "Young's Modulus must be larger than zero."
        Exit Function
    End If
    
    If W <= 0 Then
        Err.Raise vbObjectError + 1, , "Width must be be larger than zero."
        Exit Function
    End If
    
    If H <= 0 Then
        Err.Raise vbObjectError + 1, , "Height must be be larger than zero."
        Exit Function
    End If
    
    If Res <= 0 Then
        Err.Raise vbObjectError + 1, , "Resolution must be larger than zero."
        Exit Function
    End If
    
    B.Length = Length
    B.E = E
    B.B = W
    B.H = H
    B.i = W * (H ^ 3) / 12
    B.c = H / 2
    B.Resolution = Res
    
    ConstructBeam = B
    
End Function

Public Function ConstructLoad(B As Beam, Position As Double, Magnitude As Double) As PointLoad
    Dim F As PointLoad
    
    If Position > B.Length Or Position < 0 Then
        Err.Raise vbObjectError + 1, , "Position does not exist on beam."
        Exit Function
    End If
    
    F.PositionFromA = Position
    F.Magnitude = Magnitude
    
    ConstructLoad = F
    
End Function

Public Function ConstructMoment(B As Beam, Position As Double, Magnitude As Double) As PointMoment
    Dim M As PointMoment
    
    If Position > B.Length Or Position < 0 Then
        Err.Raise vbObjectError + 1, , "Position does not exist on beam."
        Exit Function
    End If
    
    M.PositionFromA = Position
    M.Magnitude = Magnitude
    
    ConstructMoment = M
    
End Function

'I just realized I made these, and forgot to use them...

Public Function ConstructDistributedLoad(B As Beam, SP As Double, EP As Double, SI As Double, EI As Double) As DistributedLoad
    Dim DL As DistributedLoad
    
    If SP < 0 Or SP > B.Length Or EP < 0 Or EP > B.Length Or EP <= SP Then
        Err.Raise vbObjectError + 1, , "One or more of your positions are invalid!"
    End If
    
    DL.StartPosition = SP
    DL.EndPosition = EP
    DL.StartIntensity = SI
    DL.EndIntensity = EI
    
    ConstructDistributedLoad = DL
End Function

Public Function GenerateEquivalentPointLoads(B As Beam, DL As DistributedLoad) As PointLoad()
    Dim Slope As Double
    Dim Base As Double
    Dim TotalRelativeDistance As Double
    Dim Midpoint As Double
    Dim PositionFromA As Double
    Dim Magnitude As Double
    Dim Index As Long
    Dim NumPoints As Long
    Dim DistLoads() As PointLoad
    
    TotalRelativeDistance = DL.EndPosition - DL.StartPosition
    Base = DL.StartIntensity
    Slope = (DL.EndIntensity - DL.StartIntensity) / (TotalRelativeDistance)
    'Using w(x) = ax + b
    
    NumPoints = CLng(TotalRelativeDistance / B.Resolution)
    ReDim DistLoads(0 To NumPoints - 1)
    For Index = 0 To NumPoints - 1
        Midpoint = Index * B.Resolution + B.Resolution / 2
        
        Magnitude = (Slope * Midpoint + Base) * B.Resolution
        PositionFromA = Midpoint + DL.StartPosition
        
        DistLoads(Index) = ConstructLoad(B, PositionFromA, Magnitude)
    Next Index
    
    GenerateEquivalentPointLoads = DistLoads
    
End Function

Public Function AppendPointLoads(Loads() As PointLoad, NewLoads() As PointLoad) As PointLoad()
    Dim Combined() As PointLoad
    Dim Index As Long
    Dim Offset As Long
    
    ReDim Combined(LBound(Loads) To UBound(Loads) + UBound(NewLoads) + 1)
    
    For Index = LBound(Loads) To UBound(Loads)
        Combined(Index) = Loads(Index)
    Next Index
    
    Offset = UBound(Loads) + 1
    
    For Index = LBound(NewLoads) To UBound(NewLoads)
        Combined(Offset + Index) = NewLoads(Index)
    Next Index
    
    AppendPointLoads = Combined
End Function
