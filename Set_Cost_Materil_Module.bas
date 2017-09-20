Attribute VB_Name = "Set_Cost_Materil_Module"
Sub Set_Cost_Materil()

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oOcc As ComponentOccurrence
    Dim SubOcc As ComponentOccurrenceProxy
    
    Dim cost As String
    cost = InputBox("Cost Materil")
    
    Do While cost = ""
        MsgBox ("Plase Enter Valid Number")
        cost = InputBox("Cost Materil")
    Loop
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        For Each SubOcc In oOcc.SubOccurrences
        
            If Left(SubOcc.Name, 4) = "Door" Then
                Dim CostMateril As StringAssetValue
                Set CostMateril = SubOcc.Definition.Document.ActiveMaterial.Item(4)
                MsgBox (CostMateril.Value)
                CostMateril.Value = cost
            End If
            
        Next
        
    Next
    
End Sub


