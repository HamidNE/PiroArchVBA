Attribute VB_Name = "initializeUnit"
Sub initializeUnit()

    Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oUserParameters As UserParameters
    Dim partParameters
    partParameters = Array("L1", "L2", "W1", "W2", "s1_L1", "s1_L2", "s1_W1", "s1_W2")

    Dim partParamExist(7) As Boolean
    Dim unitExist As Boolean
    Dim styleExist As Boolean
    Dim SubUnitExist As Boolean
    Dim styleDelExist As Boolean
    Dim styleCountExist As Boolean

    Set oUserParameters = oDoc.ComponentDefinition.Parameters.UserParameters
    For Each param In oUserParameters
                
        If param.Name = "Unit" Then
            unitExist = True
        ElseIf param.Name = "SubUnit" Then
            SubUnitExist = True
        ElseIf param.Name = "Style" Then
            styleExist = True
        ElseIf param.Name = "StyleCount" Then
            styleCountExist = True
        ElseIf param.Name = "StyleDel" Then
            styleDelExist = True
        End If

    Next

    If unitExist = False Then
        Set oParameter = oUserParameters.AddByValue("Unit", True, kBooleanUnits)
    End If
    If SubUnitExist = False Then
        Set oParameter = oUserParameters.AddByValue("SubUnit", False, kBooleanUnits)
    End If
    If styleExist = False Then
        Set oParameter = oUserParameters.AddByExpression("Style", "1", kUnitlessUnits)
    End If
    If styleCountExist = False Then
        Set oParameter = oUserParameters.AddByExpression("StyleCount", "1", kUnitlessUnits)
    End If
    If styleDelExist = False Then
        Set oParameter = oUserParameters.AddByValue("Style1_Del", "", kTextUnits)
    End If

    unitExist = False
    styleExist = False
    SubUnitExist = False
    styleDelExist = False
    styleCountExist = False

    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        Set oUserParameters = oOcc.Definition.Parameters.UserParameters
        If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
            
            For Each param In oUserParameters
                
                If param.Name = "Unit" Then
                    unitExist = True
                ElseIf param.Name = "SubUnit" Then
                    SubUnitExist = True
                ElseIf param.Name = "Style" Then
                    styleExist = True
                ElseIf param.Name = "StyleCount" Then
                    styleCountExist = True
                ElseIf param.Name = "StyleDel" Then
                    styleDelExist = True
                End If

            Next

            If unitExist = False Then
                Set oParameter = oUserParameters.AddByValue("Unit", True, kBooleanUnits)
            End If
            If SubUnitExist = False Then
                Set oParameter = oUserParameters.AddByValue("SubUnit", True, kBooleanUnits)
            End If
            If styleExist = False Then
                Set oParameter = oUserParameters.AddByExpression("Style", "1", kUnitlessUnits)
            End If
            If styleCountExist = False Then
                Set oParameter = oUserParameters.AddByExpression("StyleCount", "1", kUnitlessUnits)
            End If
            If styleDelExist = False Then
                Set oParameter = oUserParameters.AddByValue("Style1_Del", "", kTextUnits)
            End If

            unitExist = False
            styleExist = False
            SubUnitExist = False
            styleDelExist = False
            styleCountExist = False

        ElseIf oOcc.DefinitionDocumentType = kPartDocumentObject Then
        
            For Each param In oUserParameters
                
                If param.Name = partParameters(0) Then
                    partParamExist(0) = True
                ElseIf param.Name = partParameters(1) Then
                    partParamExist(1) = True
                ElseIf param.Name = partParameters(2) Then
                    partParamExist(2) = True
                ElseIf param.Name = partParameters(3) Then
                    partParamExist(3) = True
                ElseIf param.Name = partParameters(4) Then
                    partParamExist(4) = True
                ElseIf param.Name = partParameters(5) Then
                    partParamExist(5) = True
                ElseIf param.Name = partParameters(6) Then
                    partParamExist(6) = True
                ElseIf param.Name = partParameters(7) Then
                    partParamExist(7) = True
                End If

            Next

            If partParamExist(0) = False Then
                Set oParameter = oUserParameters.AddByValue(partParameters(0), "NONE", kTextUnits)
            End If
            If partParamExist(1) = False Then
                Set oParameter = oUserParameters.AddByValue(partParameters(1), "NONE", kTextUnits)
            End If
            If partParamExist(2) = False Then
                Set oParameter = oUserParameters.AddByValue(partParameters(2), "NONE", kTextUnits)
            End If
            If partParamExist(3) = False Then
                Set oParameter = oUserParameters.AddByValue(partParameters(3), "NONE", kTextUnits)
            End If
            If partParamExist(4) = False Then
                Set oParameter = oUserParameters.AddByValue(partParameters(4), "NONE", kTextUnits)
            End If
            If partParamExist(5) = False Then
                Set oParameter = oUserParameters.AddByValue(partParameters(5), "NONE", kTextUnits)
            End If
            If partParamExist(6) = False Then
                Set oParameter = oUserParameters.AddByValue(partParameters(6), "NONE", kTextUnits)
            End If
            If partParamExist(7) = False Then
                Set oParameter = oUserParameters.AddByValue(partParameters(7), "NONE", kTextUnits)
            End If

            For i = 0 To 7
                partParamExist(i) = False
            Next i

        End If
        
    Next
    
End Sub
