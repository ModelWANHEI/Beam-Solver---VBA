Attribute VB_Name = "GitTools"
Option Explicit

Public Sub ExportProject()
    Dim Project As VBIDE.VBProject
    Dim Component As VBIDE.VBComponent
    Dim ExportFolder As String
    
    Set Project = ThisWorkbook.VBProject
    ExportFolder = ThisWorkbook.Path & "\src\"
    
    If Dir(ExportFolder, vbDirectory) = "" Then
        MkDir ExportFolder
    End If
    
    For Each Component In Project.VBComponents
        Select Case Component.Type
        
            Case vbext_ct_StdModule
                Component.Export ExportFolder & Component.Name & ".bas"
            Case vbext_ct_ClassModule
                Component.Export ExportFolder & Component.Name & ".cls"
            Case vbext_ct_MSForm
                Component.Export ExportFolder & Component.Name & ".frm"
        End Select
    Next Component
    
End Sub
