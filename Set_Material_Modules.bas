Attribute VB_Name = "Set_Material_Modules"
Sub RotateAllDoor()

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim material As MaterialAsset
    Set material = oDoc.Assets.Application.AssetLibraries.Item(4).MaterialAssets.Item(69)
    
    Dim oOcc As ComponentOccurrence
    Set oOcc = oDoc.ComponentDefinition.Occurrences.Item(1)
    
    Dim partDoc As PartDocument
    Set partDoc = oOcc.Definition.Document
    
    partDoc.ActiveMaterial = material
    
    ThisApplication.ActiveDocument.Update
    
End Sub


