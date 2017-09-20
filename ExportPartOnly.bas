Attribute VB_Name = "ExportPartOnly"
Public Sub BOMExport()
    ' Set a reference to the assembly document.
    ' This assumes an assembly document is active.
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument

    ' Set a reference to the BOM
    Dim oBOM As BOM
    Set oBOM = oDoc.ComponentDefinition.BOM
  
    ' Make sure that the parts only view is enabled.
    oBOM.PartsOnlyViewEnabled = True

    ' Set a reference to the "Parts Only" BOMView
    Dim oPartsOnlyBOMView As BOMView
    Set oPartsOnlyBOMView = oBOM.BOMViews.Item("Parts Only")

    ' Export the BOM view to an Excel file
    oPartsOnlyBOMView.Export "C:\Users\HamidNE\Desktop\Temp\BOM-PartsOnly.accdb", kMicrosoftAccessFormat
End Sub

