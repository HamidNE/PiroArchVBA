Attribute VB_Name = "FixStyle2"
Sub FixStyle2()

    Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim ExistParamter As Boolean
    Dim userParams As UserParameters
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        Set userParams = oOcc.Definition.Parameters.UserParameters
        
        For Each param In userParams
            If param.Name = "s1_L1" Then
                ExistParamter = True
                Exit For
            End If
        Next
        
        If ExistParamter = False Then
            
            For Each userParam In userParams
            
                If userParam.Name = "L1" Then
                    Set param = userParams.AddByValue("s1_" & userParam.Name, userParam.Value, kTextUnits)
                ElseIf userParam.Name = "L2" Then
                    Set param = userParams.AddByValue("s1_" & userParam.Name, userParam.Value, kTextUnits)
                ElseIf userParam.Name = "W1" Then
                    Set param = userParams.AddByValue("s1_" & userParam.Name, userParam.Value, kTextUnits)
                ElseIf userParam.Name = "W2" Then
                    Set param = userParams.AddByValue("s1_" & userParam.Name, userParam.Value, kTextUnits)
                End If
            
            Next
        
        End If
    
    Next
    
    ThisApplication.ActiveDocument.Update

End Sub
