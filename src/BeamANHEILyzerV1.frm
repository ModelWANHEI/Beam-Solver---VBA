VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BeamANHEILyzerV1 
   Caption         =   "Beam Analyzer"
   ClientHeight    =   7716
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7860
   OleObjectBlob   =   "BeamANHEILyzerV1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BeamANHEILyzerV1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddPLoad_Click()

    If PositionText.Value = "" Or MagnitudeText.Value = "" Then
        MsgBox "Enter both position and magnitude."
        Exit Sub
    End If
    
    lstLoads.AddItem PositionText.Value
    lstLoads.List(lstLoads.ListCount - 1, 1) = MagnitudeText.Value
    
    PositionText.Value = ""
    MagnitudeText.Value = ""
End Sub

Private Sub AddMoment_Click()
    If PositionText.Value = "" Or MagnitudeText.Value = "" Then
        MsgBox "Enter both position and magnitude."
        Exit Sub
    End If
    
    lstMoments.AddItem PositionText.Value
    lstMoments.List(lstMoments.ListCount - 1, 1) = MagnitudeText.Value
    
    PositionText.Value = ""
    MagnitudeText.Value = ""
End Sub

Private Sub RemovePLoad_Click()
    Debug.Print "ListIndex =", lstLoads.ListIndex
    Debug.Print "ListCount =", lstLoads.ListCount

    If lstLoads.ListIndex >= 0 Then
        lstLoads.RemoveItem lstLoads.ListIndex
    End If

End Sub

Private Sub RemoveMoment_Click()
    If lstMoments.ListIndex >= 0 Then
        lstMoments.RemoveItem lstMoments.ListIndex
    End If
End Sub

Private Sub SolveButton_Click()
    Dim i As Long
    Dim B As Beam
    Dim R As Reactions
    Dim Loads() As PointLoad
    Dim Moments() As PointMoment
    
    Dim L As Double
    Dim Y As Double
    Dim W As Double
    Dim H As Double
    Dim Res As Double
    
    Dim Position As Double
    Dim Magnitude As Double
    
    If lstLoads.ListCount = 0 Then
        ReDim Loads(0)
        Position = 0 'Because you should be able to have nothing, and I shouldn't have to do any more branch casing than this.
        Magnitude = 0
    Else
        ReDim Loads(lstLoads.ListCount - 1)
        For i = 0 To lstLoads.ListCount - 1
            Position = CDbl(lstLoads.List(i, 0))
            Magnitude = CDbl(lstLoads.List(i, 1))
        
            If Position < 0 Then
                MsgBox "Invalid Position of Load!", vbOKOnly, "Invalid Parameter"
                Exit Sub
            End If
        
            Loads(i).PositionFromA = Position
            Loads(i).Magnitude = Magnitude
        Next i
    End If
    
    If lstMoments.ListCount = 0 Then
        ReDim Moments(0)
        Position = 0
        Magnitude = 0
    Else
        ReDim Moments(lstMoments.ListCount - 1)
        For i = 0 To lstMoments.ListCount - 1
            Position = CDbl(lstMoments.List(i, 0))
            Magnitude = CDbl(lstMoments.List(i, 1))
        
            If Position < 0 Then
                MsgBox "Invalid Position of Moment!", vbOKOnly, "Invalid Parameter"
                Exit Sub
            End If
        
            Moments(i).PositionFromA = Position
            Moments(i).Magnitude = Magnitude
        Next i
    End If
    
    
    L = CDbl(LengthText.Value)
    Y = CDbl(YMText.Value)
    W = CDbl(WidthText.Value)
    H = CDbl(HeightText.Value)
    Res = CDbl(ResolutionText.Value)
    
    B = ConstructBeam(L, Y, W, H, Res)
    
    R = SolveReactionForces(B, Loads, Moments)
    
    Ay.Caption = R.Ay
    By.Caption = R.By
    
    Call main(B, Loads, Moments) 'Needs moments
    
End Sub
