Attribute VB_Name = "RotateDoor_Module"
Sub RotateDoor()

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oOcc, part As ComponentOccurrence
    Set oOcc = ThisApplication.CommandManager.Pick(kAssemblyOccurrenceFilter, "Pick Door To Rotate Appearance")

    Dim oAppearance As Asset
    Set oAppearance = oOcc.Definition.Document.ActiveAppearance
    
    Dim oValue As AssetValue

    For Each oValue In oAppearance
        If oValue.ValueType = AssetValueTypeEnum.kAssetValueTextureType Then
        
            Dim oTextureAssetValue As TextureAssetValue
            Set oTextureAssetValue = oValue
            Dim oTexture As AssetTexture
            Set oTexture = oTextureAssetValue.Value
            
            If oTexture.Item("unifiedbitmap_Bitmap").Value <> "" Then
        
                If oTexture.Item("texture_WAngle").Value = 0 Then
                    oTexture.Item("texture_WAngle").Value = 90
                ElseIf oTexture.Item("texture_WAngle").Value = 90 Then
                    oTexture.Item("texture_WAngle").Value = 0
                End If
                
                Exit For
            End If
        End If
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub
