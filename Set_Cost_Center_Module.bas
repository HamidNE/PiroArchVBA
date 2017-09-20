Attribute VB_Name = "Set_Cost_Center_Module"
Sub Set_Cost_Center()

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oOcc As ComponentOccurrence
    Dim SubOcc As ComponentOccurrenceProxy
    
    Dim oUserPram As userParameters
    Dim oPram As UserParameter
    Dim iProperty As PropertySets
    
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        Dim cost As Integer
        For Each SubOcc In oOcc.SubOccurrences
        
            If Left(SubOcc.Name, 4) = "Door" Then
                Dim CostMateril As StringAssetValue
                Set CostMateril = SubOcc.Definition.Document.ActiveMaterial.Item(4)
                cost = CInt(CostMateril.Value)
            End If
            
        Next
        
        Dim Width, Height, Depth As Double
        Dim existParameters(2) As Boolean
        Set oUserPram = oOcc.Definition.Parameters.userParameters
        
        For Each oPram In oUserPram
        
            If oPram.Name = "width" Then
                Width = oPram.Value
                existParameters(0) = True
            ElseIf oPram.Name = "height" Then
                Height = oPram.Value
                existParameters(1) = True
            ElseIf oPram.Name = "depth" Then
                Depth = oPram.Value
                existParameters(2) = True
            End If
            
        Next
        
        Set iProperty = oOcc.Definition.Document.PropertySets
        
        If existParameters(0) = True And existParameters(0) = True And existParameters(0) = True Then
        
            Dim v, a, b, c As Double
        
            v = Width / 100 * Height / 100 * Depth / 100
            a = 1.2172 * (Depth / 100) ^ (-0.36)
            c = iProperty.Item("Summary Information").Item("Revision Number").Value
            iProperty.Item("Design Tracking Properties").Item("Engineer").Value = a * cost * Height / 100 * Depth / 1000
            iProperty.Item("Design Tracking Properties").Item("Cost center").Value = a * cost * v * c

        End If
        
    Next
    
End Sub

