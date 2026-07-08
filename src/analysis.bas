Attribute VB_Name = "analysis"
Public Sub CalculateStressReactions(B As Beam, Samples() As BeamSample)
    Dim i As Long
    
    For i = LBound(Samples) To UBound(Samples)
        Samples(i).Stress = Samples(i).Moment * B.c / (1000 * B.i) 'MPa conversion
    Next i
     
End Sub

Public Sub CalculateDeflection(B As Beam, Samples() As BeamSample)
    Dim index As Long
    Dim Curvature As Double
    Dim Correction As Double
    Dim dx As Double
    
    dx = B.Resolution
    
    For index = LBound(Samples) To UBound(Samples) - 1
        Curvature = (Samples(index).Moment * KN_TO_N) / (B.E * GPA_TO_PA * B.i)
        Samples(index + 1).Slope = Samples(index).Slope + (Curvature * dx)
    Next index
    
    For index = LBound(Samples) To UBound(Samples) - 1
        Samples(index + 1).Deflection = Samples(index).Deflection + (Samples(index).Slope * dx)
    Next index
    
    Correction = Samples(UBound(Samples)).Deflection / B.Length
    'v = Int(Theta dx) ; v = vo - Cx ; v(L) = vo(L) - CL ; C = vo(L)/L when v(L) = 0
    
    For index = LBound(Samples) To UBound(Samples)
        Samples(index).Slope = Samples(index).Slope - Correction
    Next index
    
    For index = LBound(Samples) To UBound(Samples)
        Samples(index).Deflection = 0
    Next index
    
    For index = LBound(Samples) To UBound(Samples) - 1
        Samples(index + 1).Deflection = Samples(index).Deflection + (Samples(index).Slope * dx)
    Next index
    
End Sub

Public Function SummarizeGlobality(dB() As BeamSample) As PointOfInterest()
    Dim GlobalPoints(SHEAR_COL To DEFLECTION_COL) As PointOfInterest
    'Warning you now, this is gonna stop being VBA and start being J with all these arrays. Sorry, not sorry.
    
    Dim NumPoints As Long
    Dim i As Long
    Dim Positional(SHEAR_COL To DEFLECTION_COL) As Long 'psuedo index pointer. You'll see what I mean later.
    
    NumPoints = CLng(UBound(dB))
    
    'Initialization
    Call InitializeGlobalPoints(GlobalPoints)
    
    'Maxima
    Call FindExtremaMagnitude(dB, GlobalPoints, MAXIMA)
    
    'Minima
    Call FindExtremaMagnitude(dB, GlobalPoints, MINIMA)
    
    'I already miss my Common Lisp. I miss my far more efficient wife.
    
    'Finding maxima points
    Call CountExtrema(dB, GlobalPoints, MAXIMA)
    
    'Finding minima points
    Call CountExtrema(dB, GlobalPoints, MINIMA)
    
    'Updating points
    Call UpdatePoints(GlobalPoints, Positional)
   
    Call InformPOIPositions(dB, GlobalPoints(), Positional(), MAXIMA)
    
    Call UpdatePoints(GlobalPoints, Positional)
    
    Call InformPOIPositions(dB, GlobalPoints(), Positional(), MINIMA)
    
    'Wanna see this thing absolutely fry your CPU for some 92 seconds?
    
    SummarizeGlobality = GlobalPoints
    
    
End Function

Private Sub InitializeGlobalPoints(GlobalPoints() As PointOfInterest)
    Dim i As Long
    For i = SHEAR_COL To DEFLECTION_COL
        ReDim GlobalPoints(i).Magnitude(MINIMA To MAXIMA)
        ReDim GlobalPoints(i).NumberOfPoints(MINIMA To MAXIMA)
        GlobalPoints(i).Magnitude(MINIMA) = 0
        GlobalPoints(i).Magnitude(MAXIMA) = 0
        GlobalPoints(i).NumberOfPoints(MINIMA) = 0
        GlobalPoints(i).NumberOfPoints(MAXIMA) = 0
    Next i
End Sub

Private Sub FindExtremaMagnitude(dB() As BeamSample, GlobalPoints() As PointOfInterest, EXTREMA As ExtremaType)
    Dim NumPoints As Long
    Dim i As Long
    
    NumPoints = CLng(UBound(dB))
    
    If EXTREMA = MAXIMA Then
    
        For i = 0 To NumPoints
            If dB(i).Shear >= GlobalPoints(SHEAR_COL).Magnitude(MAXIMA) Then
                GlobalPoints(SHEAR_COL).Magnitude(MAXIMA) = dB(i).Shear
            End If
            If dB(i).Moment >= GlobalPoints(MOMENT_COL).Magnitude(MAXIMA) Then
                GlobalPoints(MOMENT_COL).Magnitude(MAXIMA) = dB(i).Moment
            End If
            If dB(i).Stress >= GlobalPoints(STRESS_COL).Magnitude(MAXIMA) Then
                GlobalPoints(STRESS_COL).Magnitude(MAXIMA) = dB(i).Stress
            End If
            If dB(i).Slope >= GlobalPoints(SLOPE_COL).Magnitude(MAXIMA) Then
                GlobalPoints(SLOPE_COL).Magnitude(MAXIMA) = dB(i).Slope
            End If
            If dB(i).Deflection >= GlobalPoints(DEFLECTION_COL).Magnitude(MAXIMA) Then
                GlobalPoints(DEFLECTION_COL).Magnitude(MAXIMA) = dB(i).Deflection
            End If
        Next i
    
    ElseIf EXTREMA = MINIMA Then
        For i = 0 To NumPoints
            If dB(i).Shear <= GlobalPoints(SHEAR_COL).Magnitude(MINIMA) Then
                GlobalPoints(SHEAR_COL).Magnitude(MINIMA) = dB(i).Shear
            End If
            If dB(i).Moment <= GlobalPoints(MOMENT_COL).Magnitude(MINIMA) Then
                GlobalPoints(MOMENT_COL).Magnitude(MINIMA) = dB(i).Moment
            End If
            If dB(i).Stress <= GlobalPoints(STRESS_COL).Magnitude(MINIMA) Then
                GlobalPoints(STRESS_COL).Magnitude(MINIMA) = dB(i).Stress
            End If
            If dB(i).Slope <= GlobalPoints(SLOPE_COL).Magnitude(MINIMA) Then
                GlobalPoints(SLOPE_COL).Magnitude(MINIMA) = dB(i).Slope
            End If
            If dB(i).Deflection <= GlobalPoints(DEFLECTION_COL).Magnitude(MINIMA) Then
                GlobalPoints(DEFLECTION_COL).Magnitude(MINIMA) = dB(i).Deflection
            End If
        Next i
    End If
    
End Sub

Private Sub CountExtrema(dB() As BeamSample, GlobalPoints() As PointOfInterest, EXTREMA As ExtremaType)
    Dim NumPoints As Long
    Dim i As Long
    NumPoints = CLng(UBound(dB))
    
     For i = 0 To NumPoints
        If Abs(dB(i).Shear - GlobalPoints(SHEAR_COL).Magnitude(EXTREMA)) < Tolerance Then
            GlobalPoints(SHEAR_COL).NumberOfPoints(EXTREMA) = GlobalPoints(SHEAR_COL).NumberOfPoints(EXTREMA) + 1
            'God help me.
        End If
        If Abs(dB(i).Moment - GlobalPoints(MOMENT_COL).Magnitude(EXTREMA)) < Tolerance Then
            GlobalPoints(MOMENT_COL).NumberOfPoints(EXTREMA) = GlobalPoints(MOMENT_COL).NumberOfPoints(EXTREMA) + 1
        End If
        If Abs(dB(i).Stress - GlobalPoints(STRESS_COL).Magnitude(EXTREMA)) < Tolerance Then
            GlobalPoints(STRESS_COL).NumberOfPoints(EXTREMA) = GlobalPoints(STRESS_COL).NumberOfPoints(EXTREMA) + 1
        End If
        If Abs(dB(i).Slope - GlobalPoints(SLOPE_COL).Magnitude(EXTREMA)) < Tolerance Then
            GlobalPoints(SLOPE_COL).NumberOfPoints(EXTREMA) = GlobalPoints(SLOPE_COL).NumberOfPoints(EXTREMA) + 1
        End If
        If Abs(dB(i).Deflection - GlobalPoints(DEFLECTION_COL).Magnitude(EXTREMA)) < Tolerance Then
            GlobalPoints(DEFLECTION_COL).NumberOfPoints(EXTREMA) = GlobalPoints(DEFLECTION_COL).NumberOfPoints(EXTREMA) + 1
        End If
    Next i
    
End Sub

Private Sub UpdatePoints(GlobalPoints() As PointOfInterest, Positional() As Long)
    Dim i As Long
    
    For i = SHEAR_COL To DEFLECTION_COL
        ReDim Preserve GlobalPoints(i).MaximaPositionFromA(0 To GlobalPoints(i).NumberOfPoints(MAXIMA) - 1)
        ReDim Preserve GlobalPoints(i).MinimaPositionFromA(0 To GlobalPoints(i).NumberOfPoints(MINIMA) - 1)
        'Preserve so I don't have to deal with flags of whether this is the first or second update.
        Positional(i) = 0
    Next i
End Sub

Private Sub InformPOIPositions(dB() As BeamSample, GlobalPoints() As PointOfInterest, Positional() As Long, EXTREMA As ExtremaType)
    Dim NumPoints As Long
    Dim i As Long
    NumPoints = CLng(UBound(dB))
    
    If EXTREMA = MAXIMA Then
        For i = 0 To NumPoints
            If Abs(dB(i).Shear - GlobalPoints(SHEAR_COL).Magnitude(MAXIMA)) < Tolerance Then
                GlobalPoints(SHEAR_COL).MaximaPositionFromA(Positional(SHEAR_COL)) = dB(i).X
                Positional(SHEAR_COL) = Positional(SHEAR_COL) + 1
            End If
            If Abs(dB(i).Moment - GlobalPoints(MOMENT_COL).Magnitude(MAXIMA)) < Tolerance Then
                GlobalPoints(MOMENT_COL).MaximaPositionFromA(Positional(MOMENT_COL)) = dB(i).X
                Positional(MOMENT_COL) = Positional(MOMENT_COL) + 1
            End If
            If Abs(dB(i).Stress - GlobalPoints(STRESS_COL).Magnitude(MAXIMA)) < Tolerance Then
                GlobalPoints(STRESS_COL).MaximaPositionFromA(Positional(STRESS_COL)) = dB(i).X
                Positional(STRESS_COL) = Positional(STRESS_COL) + 1
            End If
            If Abs(dB(i).Slope - GlobalPoints(SLOPE_COL).Magnitude(MAXIMA)) < Tolerance Then
                GlobalPoints(SLOPE_COL).MaximaPositionFromA(Positional(SLOPE_COL)) = dB(i).X
                Positional(SLOPE_COL) = Positional(SLOPE_COL) + 1
            End If
            If Abs(dB(i).Deflection - GlobalPoints(DEFLECTION_COL).Magnitude(MAXIMA)) < Tolerance Then
                GlobalPoints(DEFLECTION_COL).MaximaPositionFromA(Positional(DEFLECTION_COL)) = dB(i).X
                Positional(DEFLECTION_COL) = Positional(DEFLECTION_COL) + 1
            End If
        Next i
    
    ElseIf EXTREMA = MINIMA Then
        For i = 0 To NumPoints
            If Abs(dB(i).Shear - GlobalPoints(SHEAR_COL).Magnitude(MINIMA)) < Tolerance Then
                GlobalPoints(SHEAR_COL).MinimaPositionFromA(Positional(SHEAR_COL)) = dB(i).X
                Positional(SHEAR_COL) = Positional(SHEAR_COL) + 1
            End If
            If Abs(dB(i).Moment - GlobalPoints(MOMENT_COL).Magnitude(MINIMA)) < Tolerance Then
                GlobalPoints(MOMENT_COL).MinimaPositionFromA(Positional(MOMENT_COL)) = dB(i).X
                Positional(MOMENT_COL) = Positional(MOMENT_COL) + 1
            End If
            If Abs(dB(i).Stress - GlobalPoints(STRESS_COL).Magnitude(MINIMA)) < Tolerance Then
                GlobalPoints(STRESS_COL).MinimaPositionFromA(Positional(STRESS_COL)) = dB(i).X
                Positional(STRESS_COL) = Positional(STRESS_COL) + 1
            End If
            If Abs(dB(i).Slope - GlobalPoints(SLOPE_COL).Magnitude(MINIMA)) < Tolerance Then
                GlobalPoints(SLOPE_COL).MinimaPositionFromA(Positional(SLOPE_COL)) = dB(i).X
                Positional(SLOPE_COL) = Positional(SLOPE_COL) + 1
            End If
            If Abs(dB(i).Deflection - GlobalPoints(DEFLECTION_COL).Magnitude(MINIMA)) < Tolerance Then
                GlobalPoints(DEFLECTION_COL).MinimaPositionFromA(Positional(DEFLECTION_COL)) = dB(i).X
                Positional(DEFLECTION_COL) = Positional(DEFLECTION_COL) + 1
            End If
        Next i
    End If
    
End Sub
