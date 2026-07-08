Attribute VB_Name = "globals"
Option Explicit
Public Const Tolerance As Double = 0.00001
Public Const KN_TO_N As Double = 1000#
Public Const GPA_TO_PA As Double = 1000000000#
Public SCROLL As Long
Public dBeam() As BeamSample
Public Const CLEAR As Long = 0
Public Const X_COL As Long = 1
Public Const SHEAR_COL As Long = 2
Public Const MOMENT_COL As Long = 3
Public Const STRESS_COL As Long = 4
Public Const SLOPE_COL As Long = 5
Public Const DEFLECTION_COL As Long = 6

Public Enum ExtremaType
    MINIMA = 0
    MAXIMA = 1
End Enum

Public Type Reactions
    Ay As Double 'kN
    By As Double 'kN
End Type

Public Type PointLoad
    PositionFromA As Double 'x, m
    Magnitude As Double 'kN
End Type

Public Type PointOfInterest
    MaximaPositionFromA() As Double 'm, array in the event of multiple points
    MinimaPositionFromA() As Double 'Because do YOU want to deal with a potentially non-rectangular matrix? Hmm? No? Thought so.
    Magnitude() As Double 'Any SI units, depending on what we're looking at
    NumberOfPoints() As Integer 'Metadata. NOP(MINIMA) for minima, NOP(MAXIMA) for maxima.
End Type

Public Type Beam
    Length As Double 'm
    E As Double 'GPa
    B As Double 'm
    H As Double 'm
    c As Double 'm
    i As Double 'm^4
    Resolution As Double 'm
End Type

Public Type BeamSample
    X As Double 'm
    Shear As Double 'kN
    Moment As Double 'kN*m
    Stress As Double 'MPa
    Slope As Double 'Radians
    Deflection As Double 'm
End Type

