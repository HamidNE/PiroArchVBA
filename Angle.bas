Attribute VB_Name = "Angle"
Sub a()

    Dim oOccs As ComponentOccurrences
    Set oOccs = ThisApplication.ActiveDocument.ComponentDefinition.Occurrences
    
    Dist = ThisApplication.MeasureTools.GetAngle(oOccs.Item(2).Definition.WorkPlanes.Item(3).Plane, oOccs.Item(3).Definition.WorkPlanes.Item(3).Plane)
    
    MsgBox (CStr(Dist))
End Sub
