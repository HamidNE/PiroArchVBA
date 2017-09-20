Attribute VB_Name = "ChangeMMToCM"
Sub ChangeMMToCM()

    Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim ascTemp As Integer
    Dim Params As Parameters
    Set Params = oDoc.ComponentDefinition.Parameters
    
    Dim userParam As UserParameter
    For Each userParam In oDoc.ComponentDefinition.Parameters.UserParameters
        ascTemp = Asc(Left(userParam.Expression, 1))
        If userParam.Units = "mm" Then
            userParam.Units = "cm"
        End If
        If ascTemp < 58 And ascTemp > 47 Then
            userParam.Expression = userParam.Value
        End If
    Next
    
    Dim modelParam As ModelParameter
    For Each modelParam In oDoc.ComponentDefinition.Parameters.ModelParameters
        If modelParam.Units = "mm" Then
            modelParam.Units = "cm"
        End If
        ascTemp = Asc(Left(modelParam.Expression, 1))
        If ascTemp < 58 And ascTemp > 47 Then
            modelParam.Expression = modelParam.Value
        End If
    Next
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        Set Params = oOcc.Definition.Parameters
        
        For Each userParam In Params.UserParameters
            If userParam.Units = "mm" Then
                userParam.Units = "cm"
            End If
            ascTemp = Asc(Left(userParam.Expression, 1))
            If ascTemp < 58 And ascTemp > 47 Then
                userParam.Expression = userParam.Value
            End If
        Next
    
        For Each modelParam In Params.ModelParameters
            If modelParam.Units = "mm" Then
                modelParam.Units = "cm"
            End If
            ascTemp = Asc(Left(modelParam.Expression, 1))
            If ascTemp < 58 And ascTemp > 47 Then
                modelParam.Expression = modelParam.Value
            End If
        Next
        
    Next
    
    ThisApplication.ActiveDocument.Update

End Sub
