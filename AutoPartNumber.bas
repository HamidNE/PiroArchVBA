Attribute VB_Name = "AutoPartNumber"
Sub AutoPartNumber()

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    If oDoc.SelectSet.Count > 0 Then
    
        Dim unitCount As Integer
        unitCount = oDoc.SelectSet.Count
        
        Dim arrayItems() As String
        ReDim arrayItems(1 To unitCount, 1 To 2)
        
        Dim i As Integer
        Dim subject As String
        
        Dim oOcc As ComponentOccurrence
        Dim userParams As UserParameters
        
        subject = InputBox("Plase Enter Subject", "Auto Part Number")
        
        For Each oOcc In oDoc.SelectSet
        
            arrayItems(i + 1, 1) = oOcc.Name
            Set userParams = oOcc.Definition.Parameters.UserParameters
            
            For Each param In userParams
                If param.Name = "width" Then
                    arrayItems(i + 1, 2) = CStr(param.Value)
                End If
            Next
            
            If arrayItems(i + 1, 2) = "" Then
                
                Var = MsgBox("Width Parameter is not Exist !!", vbCritical, "Error")
                Exit Sub
                
            End If
            
            i = i + 1
        Next
        
        For i = 1 To unitCount - 1
            For j = i + 1 To unitCount
            
                If arrayItems(i, 2) < arrayItems(j, 2) Then
                
                    temp1 = arrayItems(i, 1)
                    arrayItems(i, 1) = arrayItems(j, 1)
                    arrayItems(j, 1) = temp1
                    
                    temp2 = arrayItems(i, 2)
                    arrayItems(i, 2) = arrayItems(j, 2)
                    arrayItems(j, 2) = temp2
                
                End If
                
            Next
        Next
        
        
        For i = 1 To unitCount
            
            For Each oOccTemp In oDoc.ComponentDefinition.Occurrences
                If oOccTemp.Name = arrayItems(i, 1) Then
                    Set oOcc = oOccTemp
                    Exit For
                End If
            Next
            
            Set iProperty = oOcc.Definition.Document.PropertySets
            iProperty.Item(1).Item(2).Value = subject
            iProperty.Item(2).Item(2).Value = CStr(i)
        
        Next
        
    End If

End Sub
