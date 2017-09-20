Attribute VB_Name = "AutomaticPartNumber"
Sub AutomaticPartNumber()

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    If oDoc.ComponentDefinition.Occurrences.Count > 0 Then
    
        Dim unitCount As Integer
        unitCount = oDoc.ComponentDefinition.Occurrences.Count
        
        Dim arrayItems() As String
        ReDim arrayItems(1 To unitCount, 1 To 3)
        
        Dim i, Gapped As Integer
        Dim subject As String
        
        Dim oOcc As ComponentOccurrence
        Dim userParams As UserParameters
        
        For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
            arrayItems(i + 1, 1) = oOcc.Name
            Set userParams = oOcc.Definition.Parameters.UserParameters
            
            For Each param In userParams
                If param.Name = "width" Then
                    arrayItems(i + 1, 2) = CStr(param.Value)
                End If
            Next
            
            Set iProperty = oOcc.Definition.Document.PropertySets
            arrayItems(i + 1, 3) = iProperty.Item(1).Item(2).Value
            
            If arrayItems(i + 1, 2) = "" Then
                Gapped = Gapped + 1
                ''Var = MsgBox("Width Parameter For" & oOcc.Name & "is not Exist !!", vbCritical, "Error")
            Else
                i = i + 1
            End If
            
        Next
        
        For i = 1 To unitCount - Gapped - 1
            For j = i + 1 To unitCount - Gapped
            
                If Asc(arrayItems(i, 3)) > Asc(arrayItems(j, 3)) Then
                
                    temp1 = arrayItems(i, 1)
                    arrayItems(i, 1) = arrayItems(j, 1)
                    arrayItems(j, 1) = temp1
                    
                    temp2 = arrayItems(i, 2)
                    arrayItems(i, 2) = arrayItems(j, 2)
                    arrayItems(j, 2) = temp2
                    
                    temp3 = arrayItems(i, 3)
                    arrayItems(i, 3) = arrayItems(j, 3)
                    arrayItems(j, 3) = temp3
                
                End If
                
            Next
        Next
        
        Dim subjects() As Integer
        Dim arraySize  As Integer
        
        ReDim subjects(1)
        subjects(1) = 1

        For i = 1 To unitCount - Gapped
            If Asc(arrayItems(i, 3)) <> Asc(arrayItems(i + 1, 3)) Then

                arraySize = UBound(subjects) - LBound(subjects) + 1

                ReDim Preserve subjects(arraySize)
                subjects(arraySize) = i + 1
            
            End If
        Next
        
        For k = 1 To UBound(subjects) - 1
            For i = subjects(k) To subjects(k + 1) - 2
                For j = i + 1 To subjects(k + 1) - 1
                
                    If CInt(arrayItems(i, 2)) < CInt(arrayItems(j, 2)) Then
                    
                        temp1 = arrayItems(i, 1)
                        arrayItems(i, 1) = arrayItems(j, 1)
                        arrayItems(j, 1) = temp1
                        
                        temp2 = arrayItems(i, 2)
                        arrayItems(i, 2) = arrayItems(j, 2)
                        arrayItems(j, 2) = temp2

                        temp3 = arrayItems(i, 3)
                        arrayItems(i, 3) = arrayItems(j, 3)
                        arrayItems(j, 3) = temp3
                    
                    End If
                    
                Next
            Next
        Next k
        
        Dim Counter As Integer

        For k = 1 To UBound(subjects) - 1

            Counter = 1
            
            For i = subjects(k) To subjects(k + 1) - 1

                For Each oOccTemp In oDoc.ComponentDefinition.Occurrences
                    If oOccTemp.Name = arrayItems(i, 1) Then
                        Set oOcc = oOccTemp
                        Exit For
                    End If
                Next
                
                Set iProperty = oOcc.Definition.Document.PropertySets
                iProperty.Item(2).Item(2).Value = CStr(Counter)
                Counter = Counter + 1

            Next
        Next k
        
    End If

End Sub

