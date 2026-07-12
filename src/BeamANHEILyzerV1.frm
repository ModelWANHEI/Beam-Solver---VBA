VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BeamANHEILyzerV1 
   Caption         =   "Beam Analyzer"
   ClientHeight    =   11340
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
    
    PositionText.SetFocus
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
    
    PositionText.SetFocus
End Sub

Private Sub AddDistLoad_Click()
    If StartPositionText.Value = "" Or EndPositionText.Value = "" _
    Or StartIntensityText.Value = "" Or EndIntensityText.Value = "" Then
        MsgBox "Please complete all distributed load fields."
        Exit Sub
    End If
    
    lstDistLoads.AddItem StartPositionText.Value
    lstDistLoads.List(lstDistLoads.ListCount - 1, 1) = EndPositionText.Value
    lstDistLoads.List(lstDistLoads.ListCount - 1, 2) = StartIntensityText.Value
    lstDistLoads.List(lstDistLoads.ListCount - 1, 3) = EndIntensityText.Value
    
    StartPositionText.Value = ""
    EndPositionText.Value = ""
    StartIntensityText.Value = ""
    EndIntensityText.Value = ""
    
    StartPositionText.SetFocus
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

Private Sub RemoveDistLoad_Click()
    If lstDistLoads.ListIndex >= 0 Then
        lstDistLoads.RemoveItem lstDistLoads.ListIndex
    End If
End Sub

Private Sub SolveButton_Click()
    Dim i As Long
    Dim B As Beam
    Dim R As Reactions
    Dim Loads() As PointLoad
    Dim Moments() As PointMoment
    Dim DL As DistributedLoad
    Dim EquivalentPointLoads() As PointLoad
    
    Dim L As Double
    Dim Y As Double
    Dim W As Double
    Dim H As Double
    Dim Res As Double
    
    Dim Position As Double
    Dim Magnitude As Double
    
    Dim StartPos As Double
    Dim EndPos As Double
    Dim StartInt As Double
    Dim EndInt As Double
    'These are distload params
    
    L = CDbl(LengthText.Value)
    Y = CDbl(YMText.Value)
    W = CDbl(WidthText.Value)
    H = CDbl(HeightText.Value)
    Res = CDbl(ResolutionText.Value)
    
    B = ConstructBeam(L, Y, W, H, Res)
    
    If lstLoads.ListCount = 0 Then
        ReDim Loads(0)
        Position = 0 'Because you should be able to have nothing, and I shouldn't have to do any more branch casing than this.
        Magnitude = 0
        Loads(0) = ConstructLoad(B, 0, 0)
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
        Moments(0) = ConstructMoment(B, 0, 0)
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
    'Again, I COULD replace the manual set with the construct functions I forgot about, now. But this works.
    'And 'works' is better than 'pretty.' So, I'm not gonna touch it...
    
    If lstDistLoads.ListCount > 0 Then
        For i = 0 To lstDistLoads.ListCount - 1
            StartPos = CDbl(lstDistLoads.List(i, 0))
            EndPos = CDbl(lstDistLoads.List(i, 1))
            StartInt = CDbl(lstDistLoads.List(i, 2))
            EndInt = CDbl(lstDistLoads.List(i, 3))
            DL = ConstructDistributedLoad(B, StartPos, EndPos, StartInt, EndInt)
            EquivalentPointLoads = GenerateEquivalentPointLoads(B, DL)
            Loads = AppendPointLoads(Loads, EquivalentPointLoads)
        Next i
    End If
    
    R = SolveReactionForces(B, Loads, Moments)
    
    Ay.Caption = PrettyForce(R.Ay)
    By.Caption = PrettyForce(R.By)
    
    Call main(B, Loads, Moments)
    
End Sub
