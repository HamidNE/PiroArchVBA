Attribute VB_Name = "RotateAllDoor_Module"
Sub RotateAllDoor()

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oOcc, part As ComponentOccurrence
    Dim oAppearance As Asset
    
    Dim DoorAppearanceDeg As Integer
    DoorAppearanceDeg = 0
    DoorAppearanceDeg = InputBox(DoorAppearanceDeg, "Enter Degre")
    
    Dim Unit_Name, Part_Name As String
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        Unit_Name = oOcc.Name
        For Each part In oOcc.Definition.Occurrences
            
            Part_Name = part.Name
            
            If Left(part.Name, 4) = "Door" Or Left(part.Name, 1) = "6" Then
            
                Set oAppearance = part.Definition.Document.ActiveAppearance
                
                Dim oValue As AssetValue
            
                For Each oValue In oAppearance
                    If oValue.ValueType = AssetValueTypeEnum.kAssetValueTextureType Then
                    
                        Dim oTextureAssetValue As TextureAssetValue
                        Set oTextureAssetValue = oValue
                        Dim oTexture As AssetTexture
                        Set oTexture = oTextureAssetValue.Value
                        
                        If oTexture.Item("unifiedbitmap_Bitmap").Value <> "" Then
                    
                            oTexture.Item("texture_WAngle").Value = DoorAppearanceDeg
                            
                            Exit For
                        End If
                    End If
                Next
            End If
        Next
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub

