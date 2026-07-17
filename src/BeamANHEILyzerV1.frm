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
    
    Call AddPointLoad(PositionText.Value, MagnitudeText.Value)
    
    PositionText.Value = ""
    MagnitudeText.Value = ""
    
    PositionText.SetFocus
End Sub

Private Sub AddPointLoad(Position As Double, Magnitude As Double)
    If Magnitude = 0 Then
        Exit Sub
    End If
    lstLoads.AddItem Position
    lstLoads.List(lstLoads.ListCount - 1, 1) = Magnitude

End Sub

Private Sub AddMoment_Click()
    If PositionText.Value = "" Or MagnitudeText.Value = "" Then
        MsgBox "Enter both position and magnitude."
        Exit Sub
    End If
    
    Call AddAppliedMoment(PositionText.Value, MagnitudeText.Value)
    
    PositionText.Value = ""
    MagnitudeText.Value = ""
    
    PositionText.SetFocus
End Sub

Private Sub AddAppliedMoment(Position As Double, Magnitude As Double)
    If Magnitude = 0 Then
        Exit Sub
    End If
    lstMoments.AddItem Position
    lstMoments.List(lstMoments.ListCount - 1, 1) = Magnitude
    'Lowkey forgot why it's ListCount - 1. I needa figure out why I'm doing that later.
End Sub

Private Sub AddDistLoad_Click()
    If StartPositionText.Value = "" Or EndPositionText.Value = "" _
    Or StartIntensityText.Value = "" Or EndIntensityText.Value = "" Then
        MsgBox "Please complete all distributed load fields."
        Exit Sub
    End If
    
    Call AddDist(StartPositionText.Value, EndPositionText.Value, StartIntensityText.Value, EndIntensityText.Value)
    
    'lstDistLoads.AddItem StartPositionText.Value
    'lstDistLoads.List(lstDistLoads.ListCount - 1, 1) = EndPositionText.Value
    'lstDistLoads.List(lstDistLoads.ListCount - 1, 2) = StartIntensityText.Value
    'lstDistLoads.List(lstDistLoads.ListCount - 1, 3) = EndIntensityText.Value
    'For debugging if I need it
    
    StartPositionText.Value = ""
    EndPositionText.Value = ""
    StartIntensityText.Value = ""
    EndIntensityText.Value = ""
    
    StartPositionText.SetFocus
End Sub

Private Sub AddDist(StartPos As Double, EndPos As Double, StartInt As Double, EndInt As Double)
    If StartPos = EndPos Or (StartInt = 0 And EndInt = 0) Then
        Exit Sub
    End If
    
    lstDistLoads.AddItem StartPos
    lstDistLoads.List(lstDistLoads.ListCount - 1, 1) = EndPos
    lstDistLoads.List(lstDistLoads.ListCount - 1, 2) = StartInt
    lstDistLoads.List(lstDistLoads.ListCount - 1, 3) = EndInt
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
    Dim PName As String 'Project Name
    Dim i As Long
    Dim B As Beam
    Dim R As Reactions
    Dim Loads() As PointLoad
    Dim PassLoads() As PointLoad 'For JSON
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
    Dim DistLoads() As DistributedLoad
    'These are distload params
    
    PName = ProjectNameText.Value
    
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
        ReDim PassLoads(0)
        PassLoads(0) = ConstructLoad(B, 0, 0)
    Else
        ReDim Loads(lstLoads.ListCount - 1)
        ReDim PassLoads(lstLoads.ListCount - 1)
        For i = 0 To lstLoads.ListCount - 1
            Position = CDbl(lstLoads.List(i, 0))
            Magnitude = CDbl(lstLoads.List(i, 1))
            
            If Position < 0 Then
                MsgBox "Invalid Position of Load!", vbOKOnly, "Invalid Parameter"
                Exit Sub
            End If
        
            Loads(i).PositionFromA = Position
            Loads(i).Magnitude = Magnitude
            
            PassLoads(i) = Loads(i)
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
    
    ReDim DistLoads(0)
    DistLoads(0) = ConstructDistributedLoad(B, 0, Tolerance, 0, 0)
    
    If lstDistLoads.ListCount > 0 Then
        ReDim DistLoads(lstDistLoads.ListCount - 1)
        For i = 0 To lstDistLoads.ListCount - 1
            StartPos = CDbl(lstDistLoads.List(i, 0))
            EndPos = CDbl(lstDistLoads.List(i, 1))
            StartInt = CDbl(lstDistLoads.List(i, 2))
            EndInt = CDbl(lstDistLoads.List(i, 3))
            DL = ConstructDistributedLoad(B, StartPos, EndPos, StartInt, EndInt)
            DistLoads(i) = DL
            EquivalentPointLoads = GenerateEquivalentPointLoads(B, DL)
            Loads = AppendPointLoads(Loads, EquivalentPointLoads)
        Next i
    End If
    
    R = SolveReactionForces(B, Loads, Moments)
    
    Ay.Caption = PrettyForce(R.Ay)
    By.Caption = PrettyForce(R.By)
    
    Call ExportJSON(PName, B, PassLoads, Moments, DistLoads)
    
    Call main(B, Loads, Moments, PName)
    
End Sub

Private Sub jsonimport_Click()
    Dim FileNum As Integer
    FileNum = FreeFile
    Dim ProjectName As String
    Dim FileName As String
    Dim SaveFolder As String
    Dim ProjectFolder As String
    Dim RowString As String
    Dim Value(0 To 5) As String
    Dim PointValue(0 To 1) As String
    Dim DistValue(0 To 3) As String
    Dim Index As Long
    
    lstLoads.CLEAR
    lstMoments.CLEAR
    lstDistLoads.CLEAR

    ProjectName = ProjectNameText.Value

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
    
    FileName = ProjectFolder & "\" & ProjectName & "Config.json"
    
    If Dir(FileName) = "" Then
        MsgBox "Project configuration not found."
        Exit Sub
    End If
    
    Open FileName For Input As #FileNum
    RowString = ""
    
    Do Until EOF(FileNum)
        Line Input #FileNum, RowString
        If InStr(RowString, Q("Beam") & ": {") > 0 Then Exit Do
    Loop 'This is just a simple way to make sure you behave as a user.

    If EOF(FileNum) Then
        MsgBox "Invalid project configuration."
        Close #FileNum
        Exit Sub
    End If
    
    For Index = 0 To 5 '1 to 6; Length to Resolution
        Line Input #FileNum, RowString
    
        Value(Index) = PrepString(RowString)
    Next Index
    
    LengthText.Value = Value(BEAMLENGTH)
    YMText.Value = Value(YOUNGMOD)
    WidthText.Value = Value(BEAMWIDTH)
    HeightText.Value = Value(BEAMHEIGHT)
    ResolutionText.Value = Value(BEAMRESOLUTION)
    
    'We are now at the blankspace before point load
    
    Do Until EOF(FileNum)
        Line Input #FileNum, RowString
        If InStr(RowString, Q("PointLoads") & ": [") > 0 Then Exit Do
    Loop

    If EOF(FileNum) Then
        MsgBox "Invalid project configuration."
        Close #FileNum
        Exit Sub
    End If
    
    'We are now inside the pointloads.
    
    Do
    Line Input #FileNum, RowString

    Select Case Trim(RowString)
        Case "],", "]"
            Exit Do
        Case "{"
            ' start of an object
        Case Else
            Err.Raise vbObjectError + 2, , _
                "Expected '{' or end of array, got: " & RowString
    End Select

    Line Input #FileNum, RowString
    PointValue(DISTANCEFROMA) = PrepString(RowString)
    Line Input #FileNum, RowString
    PointValue(APPLIEDMAGN) = PrepString(RowString)

    Line Input #FileNum, RowString

    If Trim(RowString) <> "}" Then
        Err.Raise vbObjectError + 3, , _
            "Expected '}', got: " & RowString
    End If

    Call AddPointLoad(CDbl(PointValue(DISTANCEFROMA)), CDbl(PointValue(APPLIEDMAGN)))

    Loop
    
    'We are now at the blankspace prior to the applied moments.
    Do Until EOF(FileNum)
        Line Input #FileNum, RowString
        If InStr(RowString, Q("AppliedMoments") & ": [") > 0 Then Exit Do
    Loop
    
    If EOF(FileNum) Then
        MsgBox "Invalid project configuration."
        Close #FileNum
        Exit Sub
    End If
    
    Do
    Line Input #FileNum, RowString

    Select Case Trim(RowString)
        Case "],", "]"
            Exit Do
        Case "{"
            ' start of an object
        Case Else
            Err.Raise vbObjectError + 2, , _
                "Expected '{' or end of array, got: " & RowString
    End Select

    Line Input #FileNum, RowString
    PointValue(DISTANCEFROMA) = PrepString(RowString)
    Line Input #FileNum, RowString
    PointValue(APPLIEDMAGN) = PrepString(RowString)

    Line Input #FileNum, RowString

    If Trim(RowString) <> "}" Then
        Err.Raise vbObjectError + 3, , _
            "Expected '}', got: " & RowString
    End If

    Call AddAppliedMoment(CDbl(PointValue(DISTANCEFROMA)), CDbl(PointValue(APPLIEDMAGN)))

    Loop
    
    Do Until EOF(FileNum)
    Line Input #FileNum, RowString
    If InStr(RowString, Q("DistributedLoads") & ": [") > 0 Then Exit Do
    Loop

    If EOF(FileNum) Then
        MsgBox "Invalid project configuration."
        Close #FileNum
        Exit Sub
    End If
    
    Do
    Line Input #FileNum, RowString 'Should be {

    Select Case Trim(RowString)
        Case "]", "],"
            Exit Do
        Case "{"
            'Do nada
        Case Else
            Err.Raise vbObjectError + 4, , _
                "Expected '{' or end of array, got: " & RowString
    End Select

    Line Input #FileNum, RowString 'Start Position
    DistValue(DISTLOADSTARTP) = PrepString(RowString)
    Line Input #FileNum, RowString 'End Position
    DistValue(DISTLOADENDP) = PrepString(RowString)
    Line Input #FileNum, RowString 'Start Intensity
    DistValue(DISTLOADSTARTI) = PrepString(RowString)
    Line Input #FileNum, RowString 'End Intensity
    DistValue(DISTLOADENDI) = PrepString(RowString)

    Line Input #FileNum, RowString 'Should be }

    If Trim(RowString) <> "}" Then
        Err.Raise vbObjectError + 5, , _
            "Expected '}', got: " & RowString
    End If 'These are debuggers.

    Call AddDist(CDbl(DistValue(DISTLOADSTARTP)), CDbl(DistValue(DISTLOADENDP)), CDbl(DistValue(DISTLOADSTARTI)), CDbl(DistValue(DISTLOADENDI)))

Loop
    
    'We don't need anything else.
    Close #FileNum
End Sub

Private Function PrepString(TargetLine As String) As String

    If Trim(TargetLine) = "" Then
        PrepString = ""
        Exit Function
    End If

    Dim Parts() As String
    Parts = Split(TargetLine, ":")

    If UBound(Parts) < 1 Then
        Err.Raise vbObjectError + 1, , "Invalid JSON line: " & TargetLine
    End If

    PrepString = Trim(Replace(Parts(1), ",", ""))

End Function
