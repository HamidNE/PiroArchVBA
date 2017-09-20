Attribute VB_Name = "PertPartNumber"
Sub PertPartNumber()

    Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Dim userParams As UserParameters
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim isPert As Boolean
    Dim pertType As Integer
    Dim pertCountArray(1 To 3) As Integer
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        Set userParams = oOcc.Definition.Parameters.UserParameters
        For Each param In userParams
            If param.Name = "Pert" Then
                isPert = CBool(param.Value)
            ElseIf param.Name = "PertType" Then
                pertType = param.Value
            End If
        Next
        
        If isPert = True And pertType > 0 And pertType < 4 Then
            For Each childOcc In oOcc.Definition.Occurrences
                
                pertCountArray(pertType) = pertCountArray(pertType) + 1
                
                Set iProperty = childOcc.Definition.Document.PropertySets
                iProperty.Item(1).Item(2).Value = "F" & pertType
                iProperty.Item(2).Item(2).Value = pertCountArray(pertType)
                iProperty.Item(3).Item(2).Expression = "=<Subject>.<Manager>"
            
            Next
        ElseIf isPert = True And pertType = 4 Then
        
            pertCountArray(pertType) = pertCountArray(pertType) + 1
            
            Set iProperty = oOcc.Definition.Document.PropertySets
            iProperty.Item(1).Item(2).Value = "F" & pertType
            iProperty.Item(2).Item(2).Value = pertCountArray(pertType)
            iProperty.Item(3).Item(2).Expression = "=<Subject>.<Manager>"
        
        End If
    
    Next
    
End Sub

