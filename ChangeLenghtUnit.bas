Attribute VB_Name = "ChangeLenghtUnit"
Sub ChangeLenghtUnit()
    
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    oDoc.UnitsOfMeasure.LengthUnits = kCentimeterLengthUnits
    
    For Each oPart In oDoc.AllReferencedDocuments
        oPart.UnitsOfMeasure.LengthUnits = kCentimeterLengthUnits
    Next
    
End Sub
