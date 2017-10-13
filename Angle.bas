Attribute VB_Name = "Angle"
Sub a()

    'Dim oOccs As ComponentOccurrences
    'Set oOccs = ThisApplication.ActiveDocument.ComponentDefinition.Occurrences
    
    'Dim oOcc1, oOcc2 As ComponentOccurrence
    'Set oOcc1 = oOccs.Item(2)
    'Set oOcc2 = oOccs.Item(3)
    
    'Dist = ThisApplication.MeasureTools.GetAngle(oOcc1.Definition.WorkPlanes.Item(3).Plane, oOcc2.Definition.WorkPlanes.Item(3).Plane)
    
    'MsgBox (CStr(Dist))
    
    Dim oPartDoc As PartDocument
    Set oPartDoc = ThisApplication.ActiveDocument
    
    Dim face1, face2 As Face
    Set face1 = oPartDoc.SelectSet.Item(1)
    Set face2 = oPartDoc.SelectSet.Item(2)
    
    Dist = ThisApplication.MeasureTools.GetAngle(face1, face2)
    MsgBox (CStr(Dist))
    
    
End Sub
