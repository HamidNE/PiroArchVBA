Attribute VB_Name = "initStyle"
Private Sub initStyle()
    
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim userParam As UserParameter
    Dim userParams As UserParameters
    Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters
    
    If oDoc.ComponentDefinition.Parameters.IsExpressionValid("1", UnitsTypeEnum.kUnitlessUnits) = True Then
        MsgBox ("A")
    End If
    Var = userParams.IsExpressionValid("1", kUnitlessUnits)
    
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

End Sub
