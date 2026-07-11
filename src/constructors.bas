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

Public Function ConstructMoment(B As Beam, Position As Double, Magnitude As Double)
    Dim M As PointMoment
    
    If Position > B.Length Or Position < 0 Then
        Err.Raise vbObjectError + 1, , "Position does not exist on beam."
        Exit Function
    End If
    
    M.PositionFromA = Position
    M.Magnitude = Magnitude
    
    ConstructMoment = M
    
End Function
