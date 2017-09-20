Attribute VB_Name = "initStyle"
Private Sub initStyle()
    
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim userParam As UserParameter
    Dim userParams As UserParameters
    Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters
    
    Set param = userParams.AddByExpression("Style", "1", kUnitlessUnits)
    Set param = userParams.AddByExpression("StyleCount", "2", kUnitlessUnits)
    Set param = userParams.AddByValue("Style1_Del", "21", kTextUnits)
    Set param = userParams.AddByValue("Style2_Del", "41", kTextUnits)
    
    For Each userParam In userParams

        If Left(userParam.Name, 2) = "d_" Then
            Set param = userParams.AddByExpression("s2_" + userParam.Name, userParam.Expression, kCentimeterLengthUnits)
        ElseIf Left(userParam.Name, 3) = "wh_" Then
            Set param = userParams.AddByExpression("s2_" + userParam.Name, userParam.Expression, kCentimeterLengthUnits)
        End If
        
    Next
    
    Dim ExistParamter As Boolean
    Dim oOcc As ComponentOccurrence
    
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
