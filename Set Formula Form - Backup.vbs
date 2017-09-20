
''' Global farameters '''

Dim oDoc As AssemblyDocument

Dim partnameArray(100) As String
Dim materialArray(150) As String
Dim parametersArray(150) As String
Dim subOccurrenceUnit(50) As String
Dim assemblyNameArray(100) As String
Dim unitParametersArray(100) As String

Private Sub FrameMaterialAssembly_Click()

End Sub

''' Load Form '''

Private Sub UserForm_Activate()
    
    'Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    Dim oOcc As ComponentOccurrence    

    ''' Get Assembly and Part Count and Write To Array '''

    Dim count1, count2 As Integer
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        If oOcc.DefinitionDocumentType = kPartDocumentObject Then
            partnameArray(count1) = oOcc.Name
            count1 = count1 + 1
        ElseIf oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
            assemblyNameArray(count2) = oOcc.Name
            count2 = count2 + 1
        End If
        
    Next

    ''' Get Parameters Count and Write To Array '''
    
    Dim param As Parameter
    Dim parametersLenght As Integer
    parametersLenght = 0
    
    For Each param In oDoc.ComponentDefinition.Parameters.userParameters

        If Left(param.Name, 2) <> "d_" And Left(param.Name, 3) <> "wh_" Then
            If Left(param.Name, 6) <> "width_" And Left(param.Name, 6) <> "depth_" Then
                If Left(param.Name, 7) <> "height_" Then
                    parametersArray(parametersLenght) = param.Name
                    parametersLenght = parametersLenght + 1
                End If
            End If
        End If
        
    Next
    
    ''' Add Table Parameters Count and Write To Array '''

    Dim Tprams As ParameterTable
    Dim Tpram As TableParameter
    
    For Each Tprams In oDoc.ComponentDefinition.Parameters.ParameterTables
    
        For Each Tpram In Tprams.TableParameters
            parametersArray(parametersLenght) = Tpram.Name
            parametersLenght = parametersLenght + 1
        Next
        
    Next
    
    ''' Add Arrays To ComboBox's '''

    ComboBoxPart.List = partnameArray
    ComboBoxAssembly.List = assemblyNameArray
    
    lisPram.List = parametersArray

    selectIteam
    
    ''' Write Materials To Array '''

    Dim material As MaterialAsset
    count1 = 0
    
    For Each material In oDoc.Assets.Application.AssetLibraries.Item("Inventor Material Library").MaterialAssets
        materialArray(count1) = material.DisplayName
        count1 = count1 + 1
    Next
    
    ''' Add materialArray To ComboBox's '''

    ComboBoxMaterialPart.List = materialArray
    ComboBox6.List = materialArray
    ComboBox7.List = materialArray
    ComboBox8.List = materialArray
    ComboBox9.List = materialArray

    ResizePages

    Dim oOcc2 As ComponentOccurrence
    Dim Sub_oOcc As ComponentOccurrence
    Dim Temp As String
    
    ''' Check Door and Aft Are Visible '''

    Dim Check1, Check2 as Boolean

    For Each oOcc2 In oDoc.ComponentDefinition.Occurrences
        For Each Sub_oOcc In oOcc2.Definition.Occurrences
            
            If Left(Sub_oOcc.Name, 1) = "6" Then
                CheckBox1.Value = Sub_oOcc.Visible

                If Check2 = True Then
                    Exit For
                    Exit For
                Else
                    Check1 = True
                End If

            ElseIf Left(Sub_oOcc.Name, 2) = "41" Then
                CheckBox2.Value = Sub_oOcc.Visible

                If Check1 = True Then
                    Exit For
                    Exit For
                Else
                    Check2 = True
                End If

            End If
            
        Next
    Next

    ''' Set ComboBox List Array '''

    Dim listArray(4) As String

    listArray(0) = "FARSI"
    listArray(1) = "NONE"
    listArray(2) = "PVC"
    listArray(3) = "SHIAR"

    ComboBoxD1.List = listArray
    ComboBoxD2.List = listArray
    ComboBoxWH1.List = listArray
    ComboBoxWH2.List = listArray
    
End Sub

''' CommandButton & btn Events '''

Private Sub CommandButton1_Click()

    SetFormola

End Sub

Private Sub CommandButton2_Click()

    SetFormola
    Unload Me

End Sub

Private Sub CommandButton3_Click()

    Dim text1, text2 As String
    text1 = txtboxPartDFormula.Text
    text2 = txtboxPartWHFormula.Text
    
    txtboxPartDFormula.Text = text2
    txtboxPartWHFormula.Text = text1
    
    text1 = txtboxPartDValue.Text
    text2 = txtboxPartWHValue.Text
    
    txtboxPartDValue.Text = text2
    txtboxPartWHValue.Text = text1
    
End Sub

Private Sub CommandButton7_Click()

    SetFormolaAssembly
    
End Sub

Private Sub CommandButton8_Click()

    SetFormolaAssembly
    Unload Me
    
End Sub

Private Sub CommandButton10_Click()
    
    Dim oOcc, part As ComponentOccurrence
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        If oOcc.Name = ComboBoxPart.Text Then
            
            Dim oAppearance As Asset
            Set oAppearance = oOcc.Definition.Document.ActiveAppearance
            
            Dim oValue As AssetValue
        
            For Each oValue In oAppearance
                If oValue.ValueType = AssetValueTypeEnum.kAssetValueTextureType Then
                
                    Dim oTextureAssetValue As TextureAssetValue
                    Set oTextureAssetValue = oValue
                    Dim oTexture As AssetTexture
                    Set oTexture = oTextureAssetValue.Value
                    
                    If oTexture.Item("unifiedbitmap_Bitmap").Value <> "" Then
                
                        If oTexture.Item("texture_WAngle").Value = 0 Then
                            oTexture.Item("texture_WAngle").Value = 90
                        ElseIf oTexture.Item("texture_WAngle").Value = 90 Then
                            oTexture.Item("texture_WAngle").Value = 0
                        End If
                        
                        Exit For
                    End If
                End If
            Next

            Dim oParameters As Parameters
            Set oParameters = oOcc.Definition.Parameters
            
            '''''''''''' Get the parameter named "D".
            Dim oDParam As Parameter
            Set oDParam = oParameters.Item("D")
            oDParam.Name = "WH2"
            
            ''''''''''''' Get the parameter named "WH".
            Dim oWHParam As Parameter
            Set oWHParam = oParameters.Item("WH")
            oWHParam.Name = "D"
            
            oDParam.Name = "WH"
            
            ThisApplication.ActiveDocument.Update
    
            Exit For
        End If
    Next
    
End Sub

Private Sub btnPartSelect_Click()

    selectIteam

End Sub

Private Sub btnAssemblySelect_Click()

    selectIteam

End Sub

Private Sub btnSetProperty_Click()

    Dim oOcc As ComponentOccurrence
    Dim iPropertySubject As PropertySets
    Dim SubOccurrence As ComponentOccurrence

    If TextBox15.Text <> "" Then
        For Each oOcc In oDoc.SelectSet

            Set iPropertySubject = oOcc.Definition.Document.PropertySets
            iPropertySubject.Item(1).Item(2).Expression = TextBox15.Text

            If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
                
                For Each Sub_oOcc In oOcc.Definition.Occurrences
                    Set iPropertySubject = Sub_oOcc.Definition.Document.PropertySets
                    iPropertySubject.Item(1).Item(2).Expression = TextBox15.Text
                Next

            End If

        Next
    End If
    
End Sub

Private Sub btnFixPartNumber_Click()

    Dim iProperty As PropertySets
    Dim oOcc As ComponentOccurrence
    Dim Sub_oOcc As ComponentOccurrence

    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then

            Set iProperty = oOcc.Definition.Document.PropertySets
            iProperty.Item(3).Item(2).Expression = "=<Subject><Manager>"

            For Each Sub_oOcc In oOcc.Definition.Occurrences
                Set iProperty = Sub_oOcc.Definition.Document.PropertySets
                iProperty.Item(3).Item(2).Expression = "=<Subject><Manager>.<title>"
            Next

        ElseIf oOcc.DefinitionDocumentType = kPartDocumentObject Then

            Set iProperty = oOcc.Definition.Document.PropertySets
            iProperty.Item(3).Item(2).Expression = "=<Subject><Manager>.<title>"

        End If

    Next

End Sub

Private Sub btnRotateAllDoor_Click()

    RotateAllDoor

End Sub

''' ComboBox Events '''

Private Sub ComboBox5_Change()

    If ComboBox5.Text = "" Then
        ComboBox8.Text = ""
        ComboBox8.Enabled = False
        ComboBox8.BackStyle = fmBackStyleTransparent
    Else
        Dim oOcc As ComponentOccurrence
        Dim SubOccurrence As ComponentOccurrence
        
        For Each oOcc In oDoc.ComponentDefinition.Occurrences
            If oOcc.Name = ComboBoxAssembly.Text Then
                For Each SubOccurrence In oOcc.Definition.Occurrences
                    If SubOccurrence.Name = ComboBox5.Text Then
                        ComboBox8.Enabled = True
                        ComboBox8.BackStyle = fmBackStyleOpaque
                        ComboBox8.Text = SubOccurrence.Definition.Document.ActiveMaterial.DisplayName
                        Exit For
                    End If
                Next
            End If
        Next
    End If
    
End Sub

Private Sub ComboBox6_Change()

    OptionButton7.Value = 1

End Sub

Private Sub ComboBox7_Change()

    OptionButton8.Value = 1

End Sub

Private Sub ComboBox8_Change()

    OptionButton9.Value = 1

End Sub

Private Sub ComboBox9_Change()

    OptionButton12.Value = 1

End Sub

Private Sub ComboBoxAssemblyParameters_Change()

    If ComboBoxAssemblyParameters.Text = "" Then
        TextBox11.Enabled = False
        TextBox11.BackColor = &H80000004
    Else
        
        Dim unitName, parameterName As String
        unitName = ComboBoxAssembly.Text
        parameterName = ComboBoxAssemblyParameters.Text
        
        TextBox11.Enabled = True
        TextBox11.BackColor = &H80000005
        
        Dim oOcc As ComponentOccurrence

        For Each oOcc In oDoc.ComponentDefinition.Occurrences
            If unitName = oOcc.Name Then
                Exit For
            End If
        Next
        
        Dim userParameters As userParameters
        Set userParameters = oOcc.Definition.Parameters.userParameters
        
        TextBox11.Text = userParameters.Item(parameterName).Expression
        
    End If

End Sub

Private Sub ComboBoxPart_Change()

    If ComboBoxPart.Text = "" Then
        txtboxPartDFormula.Text = ""
        txtboxPartWHFormula.Text = ""
        
        txtboxPartDValue.Text = ""
        txtboxPartWHValue.Text = ""
        
        OptionButton1.Value = False
        OptionButton2.Value = False
        
        Frame13.Enabled = False
        FrameParametersPart.Enabled = False
        FrameMaterialPart.Enabled = False
        txtboxPartDFormula.BackColor = &H80000004
        txtboxPartWHFormula.BackColor = &H80000004
        
        txtboxPartDValue.BackColor = &H80000004
        txtboxPartWHValue.BackColor = &H80000004
        
        Label2.ForeColor = &H80000006
        Label3.ForeColor = &H80000006
        
        ComboBoxMaterialPart.Text = ""
        ComboBoxMaterialPart.BackStyle = fmBackStyleTransparent

        ComboBoxD1.Text = ""
        ComboBoxD2.Text = ""
        ComboBoxWH1.Text = ""
        ComboBoxWH2.Text = ""

        ComboBoxD1.BackStyle = fmBackStyleTransparent
        ComboBoxD2.BackStyle = fmBackStyleTransparent
        ComboBoxWH1.BackStyle = fmBackStyleTransparent
        ComboBoxWH2.BackStyle = fmBackStyleTransparent
        
        oDoc.SelectSet.Clear
    Else
    
        Dim partName, assemblyname As String
        partName = ComboBoxPart.Value

        assemblyname = oDoc.DisplayName
        assemblyname = Replace(assemblyname, ".iam", "")

        ''' Select Part '''

        For Each oOcc In oDoc.ComponentDefinition.Occurrences
            If oOcc.DefinitionDocumentType = kPartDocumentObject Then
            
                If partName = oOcc.Name Then
                    oDoc.SelectSet.Clear
                    oDoc.SelectSet.Select (oOcc)
                    Exit For
                End If
                
            End If
        Next

        ''' Dim Variables '''
        
        Dim componentOcc As ComponentOccurrences
        Set componentOcc = oDoc.ComponentDefinition.Occurrences
        
        Dim userParams As userParameters
        Set userParams = oDoc.ComponentDefinition.Parameters.userParameters
        
        Dim param As Parameter
        Dim existD, existWH As Boolean
        existD = False
        existWH = False
        
        
    
        Dim MLA As Integer 'MLA = MinusLocAssembly
        MLA = InStr(1, assemblyname, "-") - 1
        
        If MLA > 0 Then
            assemblyname = Left(assemblyname, MLA)
        End If
        
        partName = Replace(partName, assemblyname + "-", "")
        
        Dim DPL As Integer 'DPL = DoublePointLoc
        DPL = InStr(1, partName, ":") - 1
        Dim ML As Integer 'ML = MinusLoc
        ML = InStr(1, partName, "-") - 1
        
        partName = Left(partName, DPL)
        If ML > 0 Then
            partName = Left(partName, ML)
        End If
        
        For Each param In userParams
            If param.Name = "d_" + partName Then
                txtboxPartDFormula.Text = param.Expression
                existD = True
                Exit For
            End If
        Next
        
        For Each param In userParams
            If param.Name = "wh_" + partName Then
                txtboxPartWHFormula.Text = param.Expression
                existWH = True
                Exit For
            End If
        Next
        
        If existD = False Then
            txtboxPartDFormula.Text = ""
        End If
        
        If existWH = False Then
            txtboxPartWHFormula.Text = ""
        End If
        
        Dim partPram As ModelParameter
        
        For Each partPram In oOcc.Definition.Parameters.ModelParameters
        
            If partPram.Name = "D" Then
                txtboxPartDValue.Text = partPram.Expression
            ElseIf partPram.Name = "WH" Then
                txtboxPartWHValue.Text = partPram.Expression
            End If
            
        Next
        
        Dim partPramUser As UserParameter
        
        For Each partPramUser In oOcc.Definition.Parameters.userParameters
        
            If partPramUser.Name = "L1" Then
                ComboBoxWH1.Text = partPramUser.Value
            ElseIf partPramUser.Name = "L2" Then
                ComboBoxWH2.Text = partPramUser.Value
            ElseIf partPramUser.Name = "W1" Then
                ComboBoxD1.Text = partPramUser.Value
            ElseIf partPramUser.Name = "W2" Then
                ComboBoxD2.Text = partPramUser.Value
            End If
            
        Next

        ComboBoxD1.BackStyle = fmBackStyleOpaque
        ComboBoxD2.BackStyle = fmBackStyleOpaque
        ComboBoxWH1.BackStyle = fmBackStyleOpaque
        ComboBoxWH2.BackStyle = fmBackStyleOpaque
        
        Frame13.Enabled = True
        FrameParametersPart.Enabled = True
        FrameMaterialPart.Enabled = True
        txtboxPartDFormula.BackColor = &H80000005
        txtboxPartWHFormula.BackColor = &H80000005
        'txtboxPartDValue.BackColor = &H80000005
        'txtboxPartWHValue.BackColor = &H80000005
        
        Label2.ForeColor = &H80000012
        Label3.ForeColor = &H80000012
        
        ComboBoxMaterialPart.Text = oOcc.Definition.Document.ActiveMaterial.DisplayName
        ComboBoxMaterialPart.BackStyle = fmBackStyleOpaque
        
        txtboxCostMaterialPart.Text = oOcc.Definition.Document.ActiveMaterial.Item(4).Value
        
    End If

End Sub

Private Sub ComboBoxAssembly_Change()

    If ComboBoxAssembly.Text = "" Then
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        
        OptionButton3.Value = False
        OptionButton4.Value = False
        OptionButton5.Value = False
    
        Frame4.Enabled = False
        FrameProperties.Enabled = False
        FrameMaterialAssembly.Enabled = False
        
        TextBox3.BackColor = &H80000004
        TextBox4.BackColor = &H80000004
        TextBox5.BackColor = &H80000004
        TextBox11.BackColor = &H80000004
        TextBox12.BackColor = &H80000004
        TextBox13.BackColor = &H80000004
        TextBox14.BackColor = &H80000004
        
        lbWidthAssembly = ""
        lbDepthAssembly = ""
        lbHeightAssembly = ""
        Label26.Caption = ""
        
        Label7.ForeColor = &H80000006
        Label8.ForeColor = &H80000006
        Label9.ForeColor = &H80000006
        Label20.ForeColor = &H80000006
        Label10.ForeColor = &H80000006
        Label12.ForeColor = &H80000006
        Label21.ForeColor = &H80000006
        Label22.ForeColor = &H80000006
        Label23.ForeColor = &H80000006
        Label24.ForeColor = &H80000006
        Label25.ForeColor = &H80000006
        Label26.ForeColor = &H80000006
        Label32.ForeColor = &H80000006
        
        ComboBox5.Text = ""
        ComboBox5.BackStyle = fmBackStyleTransparent
        ComboBox6.Text = ""
        ComboBox6.BackStyle = fmBackStyleTransparent
        ComboBox7.Text = ""
        ComboBox7.BackStyle = fmBackStyleTransparent
        ComboBox8.Text = ""
        ComboBox8.BackStyle = fmBackStyleTransparent
        ComboBox9.Text = ""
        ComboBox9.BackStyle = fmBackStyleTransparent
        ComboBoxAssemblyParameters.Text = ""
        ComboBoxAssemblyParameters.BackStyle = fmBackStyleTransparent

    Else
    
        ''' Get Unit Name '''
        Dim unitName As String
        unitName = ComboBoxAssembly.Text
        
        ''' Select Unit '''
        For Each oOcc In oDoc.ComponentDefinition.Occurrences
            If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
            
                If unitName = oOcc.Name Then
                    oDoc.SelectSet.Clear
                    oDoc.SelectSet.Select (oOcc)
                    ExistPart = True
                    Exit For
                End If
                
            End If
        Next

        Dim counter As Integer
        Dim partListAssembly(25) As String

        For Each part In oOcc.Definition.Occurrences
            partListAssembly(counter) = part.Name
            counter = counter + 1
        Next

        ComboBox5.List = partListAssembly

        ''' Set Parameter For UserParameters '''

        Dim shortUnitName As String
        shortUnitName = ComboBoxAssembly.Text
        shortUnitName = Left(shortUnitName, InStr(1, shortUnitName, ":") - 1)
        shortUnitName = Left(shortUnitName, InStr(1, shortUnitName, "-") - 1)
        
        For Each param In oDoc.ComponentDefinition.Parameters.userParameters
        
            If param.Name = "width" + "_" + shortUnitName Then
                TextBox3.Text = param.Expression
                existWidth = True
                
            ElseIf param.Name = "depth" + "_" + shortUnitName Then
                TextBox4.Text = param.Expression
                existDepth = True
                
            ElseIf param.Name = "height" + "_" + shortUnitName Then
                TextBox5.Text = param.Expression
                existHeight = True
            End If
            
        Next
        
        If existWidth = False Then
            TextBox3.Text = ""
        End If
        
        If existDepth = False Then
            TextBox4.Text = ""
        End If
        
        If existHeight = False Then
            TextBox5.Text = ""
        End If
        
        '>>>'>>>' More Parametrs Frame '<<<'<<<'
        
            ''' Get Width & Depth & Height Value '''
            Set userParams = oOcc.Definition.Parameters.userParameters
            For Each param In userParams
            
                If param.Name = "width" Then
                    lbWidthAssembly.Caption = param.Expression
                    
                ElseIf param.Name = "depth" Then
                    lbDepthAssembly.Caption = param.Expression
                    
                ElseIf param.Name = "height" Then
                    lbHeightAssembly.Caption = param.Expression
                End If
                
            Next
            
            Dim count As Integer
            count = 0
            Dim parametersOcc As Parameter
            
            For Each parametersOcc In oOcc.Definition.Parameters.userParameters
                unitParametersArray(count) = parametersOcc.Name
                count = count + 1
            Next
            
            ComboBoxAssemblyParameters_Change
        
        '>>>'>>>' Material Frame '<<<'<<<'
        
            ''' Door Material '''
            
            Dim DoorOccurrence As ComponentOccurrence
            Dim ExistDoor As Boolean
            Dim Temp As String
    
            For Each DoorOccurrence In oOcc.Definition.Occurrences
                Temp = DoorOccurrence.Name
                Temp = Replace(Temp, shortUnitName + "-", "")
                
                If Left(DoorOccurrence.Name, 4) = "Door" Then
                    ExistDoor = True
                    Exit For
                ElseIf Left(Temp, 1) = "6" Then
                    ExistDoor = True
                    Exit For
                End If
            Next
            
            If ExistDoor = True Then
                ComboBox6.Text = DoorOccurrence.Definition.Document.ActiveMaterial.DisplayName
            End If

            ''' Aft Material '''
            
            Dim AftOccurrence As ComponentOccurrence
            Dim ExistAft As Boolean
    
            For Each AftOccurrence In oOcc.Definition.Occurrences

                Temp = AftOccurrence.Name
                Temp = Replace(Temp, shortUnitName + "-", "")

                If Left(Temp, 2) = "41" Then
                    ExistAft = True
                    Exit For
                End If

            Next
            
            If ExistAft = True Then
                ComboBox9.Text = AftOccurrence.Definition.Document.ActiveMaterial.DisplayName
            End If
            
            ''' Body Material '''
            
            ' If ExistDoor = True OR ExistAft = True Then
            '     If oOcc.Definition.Occurrences.Item(1).Name <> ( DoorOccurrence.Name AND AftOccurrence.Name ) Then
            '         ComboBox7.Text = oOcc.Definition.Occurrences.Item(1).Definition.Document.ActiveMaterial.DisplayName
            '     Else
            '         ComboBox7.Text = oOcc.Definition.Occurrences.Item(2).Definition.Document.ActiveMaterial.DisplayName
            '     End If
            ' Else
            '     ComboBox7.Text = oOcc.Definition.Occurrences.Item(1).Definition.Document.ActiveMaterial.DisplayName
            ' End If
            
        '>>>'>>>' Propersite '<<<'<<<'
        
        Dim iProperty As PropertySets
        Set iProperty = oOcc.Definition.Document.PropertySets
        
        Label26.Caption = iProperty.Item(3).Item(2).Value
        TextBox12.Text = iProperty.Item(1).Item(2).Value
        TextBox13.Text = iProperty.Item(2).Item(2).Value
        TextBox14.Text = iProperty.Item(3).Item(2).Expression
        
        '''''''''' UI Changes ''''''''''''
        
        Frame4.Enabled = True
        FrameProperties.Enabled = True
        FrameMaterialAssembly.Enabled = True
        
        TextBox3.BackColor = &H80000005
        TextBox4.BackColor = &H80000005
        TextBox5.BackColor = &H80000005
        TextBox11.BackColor = &H80000005
        TextBox12.BackColor = &H80000005
        TextBox13.BackColor = &H80000005
        TextBox14.BackColor = &H80000005
        
        Label7.ForeColor = &H80000012
        Label8.ForeColor = &H80000012
        Label9.ForeColor = &H80000012
        Label10.ForeColor = &H80000012
        Label12.ForeColor = &H80000012
        Label20.ForeColor = &H80000012
        Label21.ForeColor = &H80000012
        Label22.ForeColor = &H80000012
        Label23.ForeColor = &H80000012
        Label24.ForeColor = &H80000012
        Label25.ForeColor = &H80000012
        Label26.ForeColor = &H80000012
        Label32.ForeColor = &H80000012
        
        ComboBox5.BackStyle = fmBackStyleOpaque
        ComboBox6.BackStyle = fmBackStyleOpaque
        ComboBox7.BackStyle = fmBackStyleOpaque
        ComboBox9.BackStyle = fmBackStyleOpaque
        ComboBoxAssemblyParameters.BackStyle = fmBackStyleOpaque
        ComboBoxAssemblyParameters.List = unitParametersArray
        
    End If

    ComboBox5.Text = ""

End Sub

''' Founctions '''

Sub SetFormola()
    
    'Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    
    Dim partName, assemblyname As String
    partName = ComboBoxPart.Value
    assemblyname = oDoc.DisplayName
    
    Dim ExistPart As Boolean
    ExistPart = False
    
    Dim i As Integer
    i = 1
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        If oOcc.DefinitionDocumentType = kPartDocumentObject Then
        
            If partName = oOcc.Name Then
                ExistPart = True
                Exit For
            End If
            i = i + 1
            
        End If
    Next
    
    Dim D_Pram As String
    D_Pram = txtboxPartDFormula.Text
    
    Dim WH_Pram As String
    WH_Pram = txtboxPartWHFormula.Text
    
    Dim userParams As userParameters
    Set userParams = oDoc.ComponentDefinition.Parameters.userParameters
    
    Dim param As Parameter
    Dim existD, existWH As Boolean
    existD = False
    existWH = False
    
    Dim setD As Boolean
    Dim setWH As Boolean
    
    assemblyname = Replace(assemblyname, ".iam", "")
    
    Dim MLA As Integer 'MLA = MinusLocAssembly
    MLA = InStr(1, assemblyname, "-") - 1
    If MLA > 0 Then
        assemblyname = Left(assemblyname, MLA)
    End If
    partName = Replace(partName, assemblyname + "-", "")
    
    Dim DPL As Integer 'DPL = DoublePointLoc
    DPL = InStr(1, partName, ":") - 1
    Dim ML As Integer 'ML = MinusLoc
    ML = InStr(1, partName, "-") - 1
    
    partName = Left(partName, DPL)
    If ML > 0 Then
        partName = Left(partName, ML)
    End If
    
    For Each param In userParams
        If param.Name = "d_" + partName Then
            existD = True
            Exit For
        End If
    Next
    
    For Each param In userParams
        If param.Name = "wh_" + partName Then
            existWH = True
            Exit For
        End If
    Next
    
    If D_Pram <> "" Then
        If existD = False Then
            Set param = userParams.AddByExpression("d_" + partName, D_Pram, kCentimeterLengthUnits)
        Else
            userParams.Item("d_" + partName).Expression = D_Pram
        End If
        
        setD = True
    End If
    
    If WH_Pram <> "" Then
        If existWH = False Then
            Set param = userParams.AddByExpression("wh_" + partName, WH_Pram, kCentimeterLengthUnits)
        Else
            userParams.Item("wh_" + partName).Expression = WH_Pram
        End If
        
        setWH = True
    End If
    
    Dim oParameter As Parameters
    setProperty = False
    
    Dim partnametemp As String
    partnametemp = assemblyname + "-" + partName
    
    Dim material As MaterialAsset
    
    Dim materialName As String
    materialName = ComboBoxMaterialPart.Text
    
    Dim findMaterial As Boolean
    
    For Each material In oDoc.Assets.Application.AssetLibraries.Item("Inventor Material Library").MaterialAssets
        
        If materialName = material.DisplayName Then
            findMaterial = True
            Exit For
        End If
        
    Next
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences

        If oOcc.Name = ComboBoxPart.Text Then

            Set oParameter = oOcc.Definition.Parameters
            If setD = True Then
                oParameter.Item("D").Expression = userParams("d_" + partName).Value * 10
            End If
            If setWH = True Then
                oParameter.Item("WH").Expression = userParams("wh_" + partName).Value * 10
            End If
            
            oOcc.Definition.Document.ActiveMaterial = material
            
            Exit For
        End If
    Next

    Dim partPramUser As UserParameter
        
    For Each partPramUser In oOcc.Definition.Parameters.userParameters
    
        If partPramUser.Name = "L1" Then
            partPramUser.Value = ComboBoxWH1.Text
        ElseIf partPramUser.Name = "L2" Then
            partPramUser.Value = ComboBoxWH2.Text
        ElseIf partPramUser.Name = "W1" Then
            partPramUser.Value = ComboBoxD1.Text
        ElseIf partPramUser.Name = "W2" Then
            partPramUser.Value = ComboBoxD2.Text
        End If
        
    Next
    
    If txtboxCostMaterialPart.Text <> "" Then
        oOcc.Definition.Document.ActiveMaterial.Item(4).Value = CInt(txtboxCostMaterialPart.Text)
    End If

    ''' Update Assembly
    
    ThisApplication.ActiveDocument.Update
    
End Sub

Sub SetFormolaAssembly()
    
    'Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Dim Sub_oOcc As ComponentOccurrence
    
    Dim unitName As String
    unitName = ComboBoxAssembly.Value
    
    Dim existUnit As Boolean
    existUnit = False
    
    Dim i As Integer
    i = 1
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
        
            If unitName = oOcc.Name Then
                existUnit = True
                Exit For
            End If
            i = i + 1
            
        End If
    Next
    
    Dim Width_Pram As String
    Width_Pram = TextBox3.Text
    
    Dim Depth_Pram As String
    Depth_Pram = TextBox4.Text
    
    Dim Height_Pram As String
    Height_Pram = TextBox5.Text
    
    Dim userParams As userParameters
    Set userParams = oDoc.ComponentDefinition.Parameters.userParameters
    
    Dim param As Parameter
    Dim tempStr As String
    Dim existWidth, existDepth, existHeight As Boolean
    Dim setWidth, setDepth, setHeight As Boolean
    
    unitName = Left(unitName, InStr(1, unitName, ":") - 1)
    unitName = Replace(unitName, "-", "_")
    
    For Each param In userParams
        
        If param.Name = "width" + "_" + unitName Then
            existWidth = True
                
        ElseIf param.Name = "depth" + "_" + unitName Then
            existDepth = True
                
        ElseIf param.Name = "height" + "_" + unitName Then
            existHeight = True
            
        End If
            
    Next
    
    
    If Width_Pram <> "" Then
        If existWidth = False Then
            Set param = userParams.AddByExpression("width" + "_" + unitName, Width_Pram, kCentimeterLengthUnits)
        Else
            userParams.Item("width" + "_" + unitName).Expression = Width_Pram
        End If
        
        setWidth = True
    End If
    
    If Depth_Pram <> "" Then
        If existDepth = False Then
            Set param = userParams.AddByExpression("depth" + "_" + unitName, Depth_Pram, kCentimeterLengthUnits)
        Else
            userParams.Item("depth" + "_" + unitName).Expression = Depth_Pram
        End If
        
        setDepth = True
    End If
    
    If Height_Pram <> "" Then
        If existHeight = False Then
            Set param = userParams.AddByExpression("height" + "_" + unitName, Height_Pram, kCentimeterLengthUnits)
        Else
            userParams.Item("height" + "_" + unitName).Expression = Height_Pram
        End If
        
        setHeight = True
    End If
    
    Dim oParameter As Parameters
    setProperty = False
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        If oOcc.Name = ComboBoxAssembly.Text Then
        
            Set oParameter = oOcc.Definition.Parameters
            
            If setWidth = True Then
                oParameter.Item("width").Expression = userParams("width" + "_" + unitName).Value
                lbWidthAssembly.Caption = oParameter.Item("width").Expression
            End If
            
            If setDepth = True Then
                oParameter.Item("depth").Expression = userParams("depth" + "_" + unitName).Value
                lbDepthAssembly.Caption = oParameter.Item("depth").Expression
            End If
            
            If setHeight = True Then
                oParameter.Item("height").Expression = userParams("height" + "_" + unitName).Value
                lbHeightAssembly.Caption = oParameter.Item("height").Expression
            End If
            
            Exit For
        End If
        
    Next
    
    
    ''' Set Value For User Parameters '''
    
    If OptionButton6.Value = True Then
        Dim paramName As String
        paramName = ComboBoxAssemblyParameters.Text
        oParameter.Item(paramName).Expression = TextBox11.Text
    End If
    
    ''' Change Door Material

    unitName = ComboBoxAssembly.Value
    unitName = Left(unitName, InStr(1, unitName, ":") - 1)
    
    If OptionButton7.Value = True Then
        
        Dim DoorOccurrence As ComponentOccurrence
        Dim materials As AssetsEnumerator
        Dim materialDoorName As String
        Dim Temp As String
        
        materialDoorName = ComboBox6.Text
        Set materials = oDoc.Assets.Application.AssetLibraries.Item("Inventor Material Library").MaterialAssets

        For Each DoorOccurrence In oOcc.Definition.Occurrences

            Temp = DoorOccurrence.Name
            Temp = Replace(Temp, unitName + "-", "")
            
            If Left(DoorOccurrence.Name, 4) = "Door" Then
                DoorOccurrence.Definition.Document.ActiveMaterial = materials.Item(materialDoorName)
            ElseIf Left(Temp, 1) = "6" Then
                DoorOccurrence.Definition.Document.ActiveMaterial = materials.Item(materialDoorName)
            End If

        Next
        
    End If

    ''' Change Body Material
    
    If OptionButton8.Value = True Then
        
        Dim BodyOccurrence As ComponentOccurrence
        Dim materialsB As AssetsEnumerator
        Dim materialBodyName As String
        Dim TempB As String
        
        materialDoorName = ComboBox7.Text
        Set materialsB = oDoc.Assets.Application.AssetLibraries.Item("Inventor Material Library").MaterialAssets

        For Each DoorOccurrence In oOcc.Definition.Occurrences

            TempB = DoorOccurrence.Name
            TempB = Replace(TempB, unitName + "-", "")
            
            If Left(DoorOccurrence.Name, 4) <> "Door" And Left(TempB, 1) <> "6" Then

                DoorOccurrence.Definition.Document.ActiveMaterial = materialsB.Item(materialDoorName)

            End If

        Next
        
    End If

    ''' Change Aft Material
    
    If OptionButton12.Value = True Then
        
        Dim DoorOccurrence As ComponentOccurrence
        Dim materials As AssetsEnumerator
        Dim materialDoorName As String
        Dim Temp As String
        
        materialDoorName = ComboBox6.Text
        Set materials = oDoc.Assets.Application.AssetLibraries.Item("Inventor Material Library").MaterialAssets

        For Each DoorOccurrence In oOcc.Definition.Occurrences

            Temp = DoorOccurrence.Name
            Temp = Replace(Temp, unitName + "-", "")
            
            If Left(DoorOccurrence.Name, 4) = "Door" Then
                DoorOccurrence.Definition.Document.ActiveMaterial = materials.Item(materialDoorName)
            ElseIf Left(Temp, 1) = "6" Then
                DoorOccurrence.Definition.Document.ActiveMaterial = materials.Item(materialDoorName)
            End If

        Next
        
    End If
    
    ''' Change Part Material '''
    
    If OptionButton9.Value = True Then
        
        Dim PartOccurrence As ComponentOccurrence
        Dim materialPart As MaterialAsset
        Dim materialPartName As String
        Dim ExistPart As Boolean
        materialPartName = ComboBox8.Text
        
        For Each PartOccurrence In oOcc.Definition.Occurrences
            If PartOccurrence.Name = ComboBox5.Text Then
                ExistPart = True
                Exit For
            End If
        Next
        
        If ExistPart = True Then
            For Each materialPart In oDoc.Assets.Application.AssetLibraries.Item("Inventor Material Library").MaterialAssets
                If materialPartName = materialPart.DisplayName Then
                    PartOccurrence.Definition.Document.ActiveMaterial = materialPart
                    Exit For
                End If
            Next
        End If
        
    End If
    
    ''' Set Subject Unit '''
    
    If OptionButton10.Value = True Then
        
        Dim iPropertySubject As PropertySets
        Set iPropertySubject = oOcc.Definition.Document.PropertySets
        iPropertySubject.Item(1).Item(2).Expression = TextBox12.Text
        
        For Each Sub_oOcc In oOcc.Definition.Occurrences
            Set iPropertySubject = Sub_oOcc.Definition.Document.PropertySets
            iPropertySubject.Item(1).Item(2).Expression = TextBox12.Text
        Next
        
        Set iPropertySubject = oOcc.Definition.Document.PropertySets
        Label26.Caption = iPropertySubject.Item(3).Item(2).Value
        
    End If
    
    ''' Set Maneage Unit '''
    
    If OptionButton11.Value = True Then
        
        Dim iPropertyManeage As PropertySets
        Set iPropertyManeage = oOcc.Definition.Document.PropertySets
        iPropertyManeage.Item(2).Item(2).Expression = TextBox13.Text
        
        For Each Sub_oOcc In oOcc.Definition.Occurrences
            Set iPropertyManeage = Sub_oOcc.Definition.Document.PropertySets
            iPropertyManeage.Item(2).Item(2).Expression = TextBox13.Text
        Next
        
        Set iPropertyManeage = oOcc.Definition.Document.PropertySets
        Label26.Caption = iPropertyManeage.Item(3).Item(2).Value
        'TextBox7.Text = iPropertyManeage.Item(3).Item(2).Expression + " ( " + iPropertyManeage.Item(3).Item(2).Value + " )"
        
    End If
    
      
    ThisApplication.ActiveDocument.Update
    
End Sub

Sub selectIteam()
    
    If oDoc.SelectSet.count = 1 Then
    
        If oDoc.SelectSet.Item(1).DefinitionDocumentType = kPartDocumentObject Then
            MultiPage1.Value = 0
            ComboBoxPart.Text = oDoc.SelectSet.Item(1).Name
            ComboBoxPart_Change
        ElseIf oDoc.SelectSet.Item(1).DefinitionDocumentType = kAssemblyDocumentObject Then
            MultiPage1.Value = 1
            ComboBoxAssembly.Text = oDoc.SelectSet.Item(1).Name
            ComboBoxAssembly_Change
        
        End If
    End If
    
End Sub

Sub ResizePages()

    If MultiPage1.Value = 0 Then        'Part Page

        MultiPage1.Width = 300
        MultiPage1.Height = 395
        lisPram.Left = 312
        lisPram.Height = 380
        lbLisPram.Left = 330
        Set_Formula_Form1.Width = 460
        Set_Formula_Form1.Height = 435

    ElseIf MultiPage1.Value = 1 Then    'Assembly Page

        MultiPage1.Height = 444
        MultiPage1.Width = 342
        lisPram.Left = 354
        lisPram.Height = 422
        lbLisPram.Left = 372
        Set_Formula_Form1.Width = 502
        Set_Formula_Form1.Height = 482

    ElseIf MultiPage1.Value = 2 Then    'Kichen Page

        MultiPage1.Height = 192
        MultiPage1.Width = 240
        Set_Formula_Form1.Width = 262
        Set_Formula_Form1.Height = 233

    End If

End Sub

Sub RotateAllDoor()

    Dim oDoc As AssemblyDocument
    Dim oOcc, part As ComponentOccurrence
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oAppearance As Asset
    
    Dim Unit_Name, Part_Name As String
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        Unit_Name = oOcc.Name
        For Each part In oOcc.Definition.Occurrences
            
            Part_Name = part.Name
            
            If Left(part.Name, 4) = "Door" Or Left(part.Name, 1) = "6" Then
            
                Dim oValue As AssetValue
                Set oAppearance = part.Definition.Document.ActiveAppearance
                            
                For Each oValue In oAppearance
                    If oValue.ValueType = AssetValueTypeEnum.kAssetValueTextureType Then
                    
                        Dim oTexture As AssetTexture
                        Dim oTextureAssetValue As TextureAssetValue
                        Set oTextureAssetValue = oValue
                        Set oTexture = oTextureAssetValue.Value
                        
                        If oTexture.Item("unifiedbitmap_Bitmap").Value <> "" Then
                            
                            If oTexture.Item("texture_WAngle").Value = 0 Then
                                oTexture.Item("texture_WAngle").Value = 90
                            ElseIf oTexture.Item("texture_WAngle").Value = 90 Then
                                oTexture.Item("texture_WAngle").Value = 0
                            End If
                            Exit For

                        End If
                        
                    End If
                Next

            End If
        Next
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub

''' CheckBox Events '''

Private Sub CheckBox1_Change()

    Dim oOcc As ComponentOccurrence
    Dim Sub_oOcc As ComponentOccurrence
    
    Dim partName, Temp As String
    partName = oDoc.DisplayName
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        For Each Sub_oOcc In oOcc.Definition.Occurrences
            Temp = Sub_oOcc.Name
            Temp = Replace(Temp, "PartName" + "-", "")
            
            If Left(Sub_oOcc.Name, 4) = "Door" Then
                Sub_oOcc.Visible = CheckBox1.Value
            ElseIf Left(Temp, 1) = "6" Then
                Sub_oOcc.Visible = CheckBox1.Value
            End If
        Next
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub

Private Sub CheckBox2_Change()

    Dim oOcc As ComponentOccurrence
    Dim Sub_oOcc As ComponentOccurrence
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        For Each Sub_oOcc In oOcc.Definition.Occurrences
            If Left(Sub_oOcc.Name, 3) = "Aft" Then
                Sub_oOcc.Visible = CheckBox2.Value
            End If
        Next
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub

''' TextBox & txtbox Events '''

Private Sub TextBox11_Change()

    OptionButton6.Value = 1

End Sub

Private Sub TextBox12_Change()

    OptionButton10.Value = 1

End Sub

Private Sub TextBox13_Change()

    OptionButton11.Value = 1

End Sub

Private Sub TextBox3_Enter()

    OptionButton3.Value = True
    OptionButton4.Value = False
    OptionButton5.Value = False

End Sub

Private Sub TextBox4_Enter()

    OptionButton3.Value = False
    OptionButton4.Value = True
    OptionButton5.Value = False

End Sub

Private Sub TextBox5_Enter()

    OptionButton3.Value = False
    OptionButton4.Value = False
    OptionButton5.Value = True

End Sub

Private Sub txtboxPartDFormula_Enter()

    OptionButton1.Value = True
    OptionButton2.Value = False

End Sub

Private Sub txtboxPartWHFormula_Enter()

    OptionButton1.Value = False
    OptionButton2.Value = True

End Sub

Private Sub Label24_Click()

End Sub

Private Sub lisPram_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If MultiPage1.Value = 0 Then
    
        If OptionButton1.Value = True Then
            txtboxPartDFormula.Text = txtboxPartDFormula.Text + lisPram.Text
        ElseIf OptionButton2.Value = True Then
            txtboxPartWHFormula.Text = txtboxPartWHFormula.Text + lisPram.Text
        End If
        
    ElseIf MultiPage1.Value = 1 Then
    
        If OptionButton3.Value = True Then
            TextBox3.Text = TextBox3.Text + lisPram.Text
        ElseIf OptionButton4.Value = True Then
            TextBox4.Text = TextBox4.Text + lisPram.Text
        ElseIf OptionButton5.Value = True Then
            TextBox5.Text = TextBox5.Text + lisPram.Text
        End If
        
    End If
    
End Sub

Private Sub MultiPage1_Change()

    ResizePages

End Sub
