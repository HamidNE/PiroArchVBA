
	MsgBox("Rule Runing")
	
	Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    oDoc = ThisApplication.ActiveDocument

    Dim UnitName As String
	Dim tempStr As String
    UnitName = oDoc.DisplayName
	UnitName = Replace(UnitName, ".iam", "")
	
	If UnitName <> ThisDoc.FileName(False) Then

		For Each oOcc In oDoc.ComponentDefinition.Occurrences
			If Left(oOcc.Name, InStr(1, oOcc.Name, ":") - 1) = ThisDoc.FileName(False) Then
				oDoc = oOcc.Definition.Document
				
				UnitName = oDoc.DisplayName
				UnitName = Replace(UnitName, ".iam", "")
				UnitName = Left(UnitName, InStr(1, UnitName, "-") - 1)
				
				Exit For
			End If
		Next
		
	End If
	
	Dim PartCounter, PartSize As Integer
	Dim AssemblyCounter, AssemblySize As Integer

	''' Get Part and Assembly Size

	For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        If oOcc.DefinitionDocumentType = kPartDocumentObject Then
            PartSize = PartSize + 1
        ElseIf oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
            AssemblySize = AssemblySize + 1
        End If
        
    Next
    
	Dim unitNameArray(AssemblySize) As String	
	
	If InStr(1, UnitName, "-") > 0 Then
		UnitName = Left(UnitName, InStr(1, UnitName, "-")-1)
	End If

    PartCounter = 0
    AssemblyCounter = 0

    Dim width_PramArray(AssemblySize) As String
    Dim depth_PramArray(AssemblySize) As String
    Dim height_PramArray(AssemblySize) As String
    
    Dim UnitPramValue(AssemblySize,3) As Double
    Dim ExistUnit(AssemblySize,3) As Boolean

    Dim partnameArray(PartSize), partnameTemp As String
	Dim d_PramArray(PartSize), wh_PramArray(PartSize) As String
	Dim PartPramValue(PartSize,2) As Double
	Dim ExistPartParameters(PartSize,2) As Boolean
	
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    	If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
			
			tempStr = oOcc.Name
			tempStr = Replace(tempStr, "-", "_")
			unitNameArray(AssemblyCounter) = tempStr
			tempStr = Left(tempStr, InStr(1, tempStr, ":") - 1)

            If Style = 1 Then
                width_PramArray(AssemblyCounter) = "width_" + tempStr
                depth_PramArray(AssemblyCounter) = "depth_" + tempStr
                height_PramArray(AssemblyCounter) = "height_" + tempStr
            ElseIf Style > 1 Then
                width_PramArray(AssemblyCounter) = "s" + CStr(Style) + "_width_" + tempStr
                depth_PramArray(AssemblyCounter) = "s" + CStr(Style) + "_depth_" + tempStr
                height_PramArray(AssemblyCounter) = "s" + CStr(Style) + "_height_" + tempStr
            End If

	        AssemblyCounter = AssemblyCounter + 1

	    ElseIf oOcc.DefinitionDocumentType = kPartDocumentObject Then

	    	partnameArray(PartCounter) = oOcc.Name

            If Style = 1 Then
                d_PramArray(PartCounter) = "d_" + Left(oOcc.Name, 2)			
                wh_PramArray(PartCounter) = "wh_" + Left(oOcc.Name, 2)
            ElseIf Style > 1 Then
                d_PramArray(PartCounter) = "s" + CStr(Style) + "_d_" + Left(oOcc.Name, 2)			
                wh_PramArray(PartCounter) = "s" + CStr(Style) + "_wh_" + Left(oOcc.Name, 2)
            End If
			
	        PartCounter = PartCounter + 1

		End If
    Next
	
	Dim param As Parameter
	Dim userParams As userParameters
    userParams = oDoc.ComponentDefinition.Parameters.userParameters

    If Style = 1 Then
	
        For Each param In userParams

            If Left(param.Name,2) = "d_" Then
                For i = 0 To PartSize    				
                    If param.Name = d_PramArray(i) Then
                        PartPramValue(i, 0) = param.Value
                        ExistPartParameters(i, 0) = True
                        Exit For
                    End If			
                Next            
            ElseIf Left(param.Name,3) = "wh_" Then
                For i = 0 To PartSize		
                    If param.Name = wh_PramArray(i) Then
                        PartPramValue(i, 1) = param.Value
                        ExistPartParameters(i, 1) = True
                        Exit For
                    End If				
                Next            
            ElseIf Left(param.Name,6) = "width_" Then
                For i = 0 To AssemblySize
                    If param.Name = width_PramArray(i) Then
                        UnitPramValue(i, 0) = param.Value
                        ExistUnit(i, 0) = True
                        Exit For
                    End If
                Next
            ElseIf Left(param.Name,6) = "depth_" Then
                For i = 0 To AssemblySize
                    If param.Name = depth_PramArray(i) Then
                        UnitPramValue(i, 1) = param.Value
                        ExistUnit(i, 1) = True
                        Exit For
                    End If
                Next
            ElseIf Left(param.Name,7) = "height_" Then
                For i = 0 To AssemblySize
                    If param.Name = height_PramArray(i) Then
                        UnitPramValue(i, 2) = param.Value
                        ExistUnit(i, 2) = True
                        Exit For
                    End If
                Next
            End If

        Next

    ElseIf Style > 1 Then
		
        For Each param In userParams

            If Left(param.Name,5) = "s" + CStr(Style) + "_d_" Then
                For i = 0 To PartSize    				
                    If param.Name = d_PramArray(i) Then
                        PartPramValue(i, 0) = param.Value
                        ExistPartParameters(i, 0) = True
                        Exit For
                    End If			
                Next            
            ElseIf Left(param.Name,6) = "s" + CStr(Style) + "_wh_" Then
                For i = 0 To PartSize		
                    If param.Name = wh_PramArray(i) Then
                        PartPramValue(i, 1) = param.Value
                        ExistPartParameters(i, 1) = True
                        Exit For
                    End If				
                Next            
            ElseIf Left(param.Name,9) = "s" + CStr(Style) + "_width_" Then
                For i = 0 To AssemblySize
                    If param.Name = width_PramArray(i) Then
                        UnitPramValue(i, 0) = param.Value
                        ExistUnit(i, 0) = True
                        Exit For
                    End If
                Next
            ElseIf Left(param.Name,9) = "s" + CStr(Style) + "_depth_" Then
                For i = 0 To AssemblySize
                    If param.Name = depth_PramArray(i) Then
                        UnitPramValue(i, 1) = param.Value
                        ExistUnit(i, 1) = True
                        Exit For
                    End If
                Next
            ElseIf Left(param.Name,10) = "s" + CStr(Style) + "_height_" Then
                For i = 0 To AssemblySize
                    If param.Name = height_PramArray(i) Then
                        UnitPramValue(i, 2) = param.Value
                        ExistUnit(i, 2) = True
                        Exit For
                    End If
                Next
            End If

        Next
    End If
		
    For i = 0 To AssemblySize
	
		unitNameArray(i) = Replace(unitNameArray(i), "_", "-")
		
    	If ExistUnit(i, 0) = True Then
    		Parameter(unitNameArray(i), "width") = UnitPramValue(i, 0)
		End If
    	If ExistUnit(i, 1) = True Then
    		Parameter(unitNameArray(i), "depth") = UnitPramValue(i, 1)
    	End If
    	If ExistUnit(i, 2) = True Then
    		Parameter(unitNameArray(i), "height") = UnitPramValue(i, 2)
    	End If

    Next

    For i = 0 To PartSize

    	If ExistPartParameters(i, 0) = True Then
			tempStr = CStr(PartPramValue(i, 0))
			tempStr = Replace(tempStr, "/", ".")
    		Parameter(partnameArray(i), "D") = tempStr
		End If		
    	If ExistPartParameters(i, 1) = True Then
			tempStr = CStr(PartPramValue(i, 1))
			tempStr = Replace(tempStr, "/", ".")
    		Parameter(partnameArray(i), "WH") = tempStr
    	End If

    Next

    If Unit = True Then
        oDoc.PropertySets.Item(3).Item(2).Expression = "=<Subject><Manager>"
    End If
	
	For i = 0 To PartCounter-1		

		iProperties.Value(partnameArray(i),"Summary", "Title") = Left(partnameArray(i), 2)			
		iProperties.Value(partnameArray(i),"Project", "Part Number") = "=<Subject><Manager>.<Title>"
		iProperties.Value(partnameArray(i),"Summary", "Manager") = iProperties.Value("Summary", "Manager")		
		iProperties.Value(partnameArray(i),"Summary", "Subject") = iProperties.Value("Summary", "Subject")

	Next
	
	iLogicVb.UpdateWhenDone = True

	''' Key Parameters '''
	Dim tempInt As Integer
	
	tempInt= Style
	
	tempInt = width
	tempInt = width1
	tempInt = width2
	tempInt = width3
	tempInt = width4
	tempInt = width5

    tempInt = depth
    tempInt = depth1
	tempInt = depth2
	tempInt = depth3
	tempInt = depth4
	tempInt = depth5
	
    tempInt = height
    tempInt = height1
    tempInt = height2
	tempInt = height3
	tempInt = height4
	tempInt = height5