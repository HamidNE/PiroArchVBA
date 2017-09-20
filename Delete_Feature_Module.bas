Attribute VB_Name = "Delete_Feature_Module"
Sub Delete_Feature()

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oOcc As ComponentOccurrence

    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        For Each Feature In oOcc.Definition.Features
            If Feature.Suppressed = True Then
                Feature.Delete
            End If
        Next
    Next
    
End Sub
