VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Set_Formula_Form2 
   Caption         =   "Set Formula"
   ClientHeight    =   9432.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12216
   OleObjectBlob   =   "Set_Formula_Form2_v1.2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Set_Formula_Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''' Global Parameters '''

Dim oDoc As AssemblyDocument

Dim isUnit As Boolean
Dim started As Boolean
Dim partnameArray(100) As String
Dim materialArray(150) As String
Dim parametersArray(150) As String
Dim parametersValueArray(150) As String
Dim subOccurrenceUnit(50) As String
Dim assemblyNameArray(100) As String
Dim unitParametersNames(100) As String
Dim unitParametersValues(100) As String
Dim keyParametersName(20) As String
Dim keyParametersValue(20) As String

Private Sub CommandButton11_Click()
    Management_Styles.Show
End Sub

''' Load Form '''

Private Sub UserForm_Activate()
    
    'Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    Dim oOcc As ComponentOccurrence

    ''' Get Assembly and Part Count and Write To Array '''

    Dim Count1, Count2 As Integer
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        If oOcc.DefinitionDocumentType = kPartDocumentObject Then
            partnameArray(Count1) = oOcc.Name
            Count1 = Count1 + 1
        ElseIf oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
            assemblyNameArray(Count2) = oOcc.Name
            Count2 = Count2 + 1
        End If
        
    Next

    ''' Get Parameters Count and Write To Array '''
    
    Dim param As Parameter
    Dim parametersLenght As Integer
    parametersLenght = 0
    
    Dim Style As Integer
    Dim styleCount As Integer
    Dim keyParametersCount As Integer
    Style = 1
    styleCount = 1
    keyParametersCount = 0
    
    For Each param In oDoc.ComponentDefinition.Parameters.UserParameters

        If param.IsKey = True Then
            parametersArray(parametersLenght) = param.Name
            parametersValueArray(parametersLenght) = param.Expression
            parametersLenght = parametersLenght + 1

            keyParametersName(keyParametersCount) = param.Name
            keyParametersValue(keyParametersCount) = param.Expression
            keyParametersCount = keyParametersCount + 1
        End If
        
        If param.Name = "StyleCount" Then
            styleCount = param.Value
        ElseIf param.Name = "Style" Then
            Style = param.Value
        End If

    Next
    
    
    
    MoreParameters (keyParametersCount)
    
    ''' Add Table Parameters Count and Write To Array '''

    Dim Tprams As ParameterTable
    Dim Tpram As TableParameter
    
    For Each Tprams In oDoc.ComponentDefinition.Parameters.ParameterTables
    
        For Each Tpram In Tprams.TableParameters
            parametersArray(parametersLenght) = Tpram.Name
            parametersValueArray(parametersLenght) = Tpram.Value
            parametersLenght = parametersLenght + 1
        Next
        
    Next
    
    ''' Add Arrays To ComboBox's '''

    ComboBoxPart.List = partnameArray
    ComboBoxAssembly.List = assemblyNameArray
    ComboBoxAssemblyParameters.List = parametersArray

    lisPram.ColumnWidths = "60;40"
    lisPram.Clear
    
    For i = 0 To parametersLenght
        lisPram.AddItem parametersArray(i)
        lisPram.List(i, 1) = parametersValueArray(i)
    Next

    selectIteam
    
    ''' Write Materials To Array '''

    Dim material As MaterialAsset
    Count1 = 0
    
    For Each material In oDoc.Assets.Application.AssetLibraries.Item("Inventor Material Library").MaterialAssets
        materialArray(Count1) = material.DisplayName
        Count1 = Count1 + 1
    Next
    
    ''' Add materialArray To ComboBox's '''

    ComboBoxMaterialPart.List = materialArray
    ComboBox6.List = materialArray
    ComboBox7.List = materialArray
    ComboBox8.List = materialArray
    ComboBox9.List = materialArray
    
    ''' Check Door and Aft Are Visible '''

    Dim Check1, Check2 As Boolean
    Dim oOcc2 As ComponentOccurrence
    Dim Sub_oOcc As ComponentOccurrence

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

    listArray(0) = "NONE"
    listArray(1) = "PVC"
    listArray(2) = "FARSI"
    listArray(3) = "SHIAR"

    ComboBoxD1.List = listArray
    ComboBoxD2.List = listArray
    ComboBoxWH1.List = listArray
    ComboBoxWH2.List = listArray

    If ComboBoxAssembly.Text <> "" Then
        For Each occ In oDoc.ComponentDefinition.Occurrences
            If ComboBoxAssembly.Text = occ.Name Then

                oDoc.SelectSet.Clear
                oDoc.SelectSet.Select (occ)
                Exit For

            End If
        Next
    End If

    If ComboBoxPart.Text <> "" Then
        For Each occ In oDoc.ComponentDefinition.Occurrences
            If ComboBoxPart.Text = occ.Name Then

                oDoc.SelectSet.Clear
                oDoc.SelectSet.Select (occ)
                Exit For

            End If
        Next
    End If

    If oDoc.SelectSet.Count = 0 Then
        CheckIsUnit
    End If

    ResizePages
    
    started = True
    
    Dim oProperty As PropertySets
    Set oProperty = oDoc.ComponentDefinition.Document.PropertySets
    
    txtComment.Text = oProperty.Item(1).Item(5).Expression
    
    ComboBoxStyle.Clear
    For i = 1 To styleCount
        ComboBoxStyle.AddItem ("Style" & i)
    Next
    ComboBoxStyle.Text = "Style" & Style
    
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
    
    If isUnit = True Then
        SetFormolaAssemblyMaster
    Else
        SetFormolaAssembly
    End If
    
End Sub

Private Sub CommandButton8_Click()

    If isUnit = True Then
        SetFormolaAssemblyMaster
    Else
        SetFormolaAssembly
    End If
    
    Unload Me
    
End Sub

Private Sub CommandButton10_Click()
    
    Dim isRotated As Boolean
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
                        
                        isRotated = True
                        
                        Exit For
                    End If
                End If
            Next
            
            If isRotated = True Then

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
                
                CommandButton3_Click
                SetFormola
                
            End If

            Exit For
        End If
    Next

    ThisApplication.ActiveDocument.Update

End Sub

Private Sub btnAddStyle_Click()
        
    If MsgBox("Are you Sure To Add Style ?", vbYesNo + vbQuestion, "Add Style") = vbYes Then

        Dim styleCount As Integer
        Dim userParam As UserParameter
        Dim userParams As UserParameters
        Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters
        
        For Each userParam In userParams
            If userParam.Name = "StyleCount" Then
                styleCount = userParam.Value + 1
                userParam.Expression = userParam.Value + 1
            End If
        Next
        
        For Each userParam In userParams
    
            If Left(userParam.Name, 2) = "d_" Then
                Set param = userParams.AddByExpression("s" & styleCount & "_" & userParam.Name, userParam.Expression, kCentimeterLengthUnits)
            ElseIf Left(userParam.Name, 3) = "wh_" Then
                Set param = userParams.AddByExpression("s" & styleCount & "_" & userParam.Name, userParam.Expression, kCentimeterLengthUnits)
            End If
            
        Next
        
        ComboBoxStyle.AddItem "Style" & styleCount
        
    End If
    
End Sub

Private Sub btnDelUnitStyle_Click()
        
    SetAllPartNormal
    Dim delUnits As String
    Dim existDelParam As Boolean
    Dim oOcc As ComponentOccurrence
    
    For Each oOcc In oDoc.SelectSet
        delUnits = delUnits & "-" & Left(oOcc.Name, 2)
    Next
    
    delUnits = Mid(delUnits, 2)
    
    Dim userParam As UserParameter
    Dim userParams As UserParameters
    Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters
    
    For Each userParam In userParams

        If userParam.Name = "Style" & Right(ComboBoxStyle.Text, 1) & "_Del" Then
            existDelParam = True
            userParam.Value = delUnits
        End If
        
    Next
    
    If existDelParam = False Then
        Set userParam = userParams.AddByValue("Style" & Right(ComboBoxStyle.Text, 1) & "_Del", delUnits, kTextUnits)
    End If
    
    ComboBoxStyle_Change
    
End Sub

Private Sub btnPartSelect_Click()

    selectIteam

End Sub

Private Sub btnAssemblySelect_Click()

    selectIteam

End Sub

Private Sub btnSaveComment_Click()

    Dim oProperty As PropertySets
    Set oProperty = oDoc.ComponentDefinition.Document.PropertySets
    
    oProperty.Item(1).Item(5).Expression = txtComment.Text
    
End Sub

Private Sub btnSetProperty_Click()

    ''' Kichen Page '''

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

    ''' Kichen Page '''

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

    ''' Part Material in Assembly Page '''

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
    
    If started = True Then
        CheckBox7.Value = 1
    End If

End Sub

Private Sub ComboBox7_Change()
    
    If started = True Then
        CheckBox8.Value = 1
    End If

End Sub

Private Sub ComboBox8_Change()

    If started = True Then
        CheckBox9.Value = 1
    End If

End Sub

Private Sub ComboBox9_Change()

    If started = True Then
        CheckBox12.Value = 1
    End If

End Sub

Sub SetAllPartNormal()

    Dim oOcc As ComponentOccurrence
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        If oOcc.BOMStructure = kReferenceBOMStructure Then
            oOcc.Visible = True
            oOcc.BOMStructure = kDefaultBOMStructure
        End If
        
    Next
    
End Sub

Sub SetPartReference(ByVal partName As String)
    
    Dim oOcc As ComponentOccurrence
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        If Left(oOcc.Name, 2) = partName Then
            If oOcc.BOMStructure = kNormalBOMStructure Then
                oOcc.Visible = False
                oOcc.BOMStructure = kReferenceBOMStructure
            End If
        End If
        
    Next
    
End Sub

Private Sub ComboBoxStyle_Change()

    If ComboBoxStyle.Text <> "" Then
        
        Dim numStyle As Integer
        numStyle = CInt(Right(ComboBoxStyle.Text, 1))
        
        SetAllPartNormal
    
        Dim userParam As UserParameter
        For Each userParam In oDoc.ComponentDefinition.Parameters.UserParameters
            If userParam.Name = "Style" Then
            
                userParam.Expression = numStyle
                ComboBoxPart_Change
                
            ElseIf userParam.Name = "Style" & numStyle & "_Del" Then
            
                Dim WrdArray() As String
                Dim partNum
                WrdArray() = Split(userParam.Value, "-")
                
                For Each partNum In WrdArray
                    SetPartReference (partNum)
                Next
                
            End If
        Next
        
        
        
    End If
    
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
        
        Dim userParams As UserParameters
        
        If isUnit = True Then
            Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters
        Else
            Set userParams = oOcc.Definition.Parameters.UserParameters
        End If
        
        TextBox11.Text = userParams.Item(parameterName).Expression
        
    End If

End Sub

Private Sub ComboBoxPart_Change()

    If ComboBoxPart.Text = "" Then
    
        txtboxPartDFormula.Text = ""
        txtboxPartWHFormula.Text = ""
        
        txtboxPartDValue.Text = ""
        txtboxPartWHValue.Text = ""
        
        CheckBox13.Value = False
        CheckBox14.Value = False
        
        Frame13.Enabled = False
        FrameParametersPart.Enabled = False
        FrameMaterialPart.Enabled = False
        txtboxPartDFormula.BackColor = &H80000004
        txtboxPartWHFormula.BackColor = &H80000004
        txtboxCostMaterialPart.BackColor = &H80000004
        
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

        CommandButton1.Enabled = False
        CommandButton2.Enabled = False
        
        oDoc.SelectSet.Clear
        
    Else
    
        Dim partName, shortPartName, assemblyname As String

        partName = ComboBoxPart.Value
        shortPartName = Left(partName, 2)

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
        
        Dim userParams As UserParameters
        Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters
        
        Dim param As Parameter
        Dim existD, existWH As Boolean
        existD = False
        existWH = False

        ''' Find "d_" and "wh_" Parameters and get Expression '''
        
        Dim numStyle As String
        numStyle = CInt(Right(ComboBoxStyle.Text, 1))
        
        If numStyle = 1 Then
            For Each param In userParams
            
                If param.Name = "d_" + shortPartName Then
                    txtboxPartDFormula.Text = param.Expression
                    existD = True
                ElseIf param.Name = "wh_" + shortPartName Then
                    txtboxPartWHFormula.Text = param.Expression
                    existWH = True
                End If
            
            Next
        ElseIf numStyle > 1 Then
            For Each param In userParams
            
                If param.Name = "s" & numStyle & "_d_" & shortPartName Then
                    txtboxPartDFormula.Text = param.Expression
                    existD = True
                ElseIf param.Name = "s" & numStyle & "_wh_" & shortPartName Then
                    txtboxPartWHFormula.Text = param.Expression
                    existWH = True
                End If
                
            Next
        End If
        
        If existD = False Then
            txtboxPartDFormula.Text = ""
        End If
        
        If existWH = False Then
            txtboxPartWHFormula.Text = ""
        End If
        
        ''' Get Expression Of "D" and "WH" Parameters '''

        Dim partModelPrams As ModelParameters
        Set partModelPrams = oOcc.Definition.Parameters.ModelParameters
        
        If existD = True Then
            txtboxPartDValue.Text = partModelPrams.Item("D").Expression
        End If
        If existWH = True Then
            txtboxPartWHValue.Text = partModelPrams.Item("WH").Expression
        End If

        ''' Get Value Of Farsi Family Parameters '''
        
        Set userParams = oOcc.Definition.Parameters.UserParameters
        
        For Each userParam In userParams

            If userParam.Name = "W1" Then
                ComboBoxD1.Text = userParam.Value
            ElseIf userParam.Name = "W2" Then
                ComboBoxD2.Text = userParam.Value
            ElseIf userParam.Name = "L1" Then
                ComboBoxWH1.Text = userParam.Value
            ElseIf userParam.Name = "L2" Then
                ComboBoxWH2.Text = userParam.Value
            End If

        Next

        ''' Set UI '''

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
        txtboxCostMaterialPart.BackColor = &H80000005
        
        Label2.ForeColor = &H80000012
        Label3.ForeColor = &H80000012

        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
        
        ComboBoxMaterialPart.Text = oOcc.Definition.Document.ActiveMaterial.DisplayName
        ComboBoxMaterialPart.BackStyle = fmBackStyleOpaque
        
        txtboxCostMaterialPart.Text = oOcc.Definition.Document.ActiveMaterial.Item(4).Value
        
    End If

End Sub

Private Sub ComboBoxAssembly_Change()

    If ComboBoxAssembly.Text = "" And isUnit = False Then
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        
        CheckBox3.Value = False
        CheckBox4.Value = False
        CheckBox5.Value = False
    
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

        CommandButton7.Enabled = False
        CommandButton8.Enabled = False

    ElseIf isUnit = True And oDoc.SelectSet.Count = 0 Then

        ''' Get Unit Name '''
        Dim unitName As String
        unitName = oDoc.DisplayName
        unitName = Replace(unitName, ".iam", "")

        ''' Get Part List in Assembly '''

        Dim counter As Integer
        Dim partListAssembly(50) As String

        For Each part In oDoc.ComponentDefinition.Occurrences
            partListAssembly(counter) = part.Name
            counter = counter + 1
        Next

        ComboBox5.List = partListAssembly

        ''' Set Parameter For UserParameters '''
        
        Dim userParams As UserParameters
        Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters

        TextBox3.Text = userParams.Item("width").Value
        TextBox4.Text = userParams.Item("depth").Value
        TextBox5.Text = userParams.Item("height").Value

        lbWidthAssembly.Caption = userParams.Item("width").Expression
        lbDepthAssembly.Caption = userParams.Item("depth").Expression
        lbHeightAssembly.Caption = userParams.Item("height").Expression
        
        '>>>'>>>' More Parametrs Frame '<<<'<<<'
        
        counter = 0
        For Each param In userParams
        
            If Left(param.Name, 2) <> "d_" And Left(param.Name, 3) <> "wh_" Then
                unitParametersNames(counter) = param.Name
                unitParametersValues(counter) = param.Value
                counter = counter + 1
            End If

        Next
        
        ComboBoxAssemblyParameters_Change
        
        '>>>'>>>' Material Frame '<<<'<<<'

        Dim ExistDoor As Boolean
        Dim ExistAft As Boolean
        Dim ExistBody As Boolean

        For Each occurrence In oDoc.ComponentDefinition.Occurrences

            ''' Door Material '''
            If ExistDoor = False And Left(occurrence.Name, 1) = "6" Then

                ExistDoor = True
                ComboBox6.Text = occurrence.Definition.Document.ActiveMaterial.DisplayName

            ''' Aft Material '''
            ElseIf ExistAft = False And Left(occurrence.Name, 2) = "41" Then

                ExistAft = True
                ComboBox9.Text = occurrence.Definition.Document.ActiveMaterial.DisplayName

            ''' Body Material '''
            ElseIf ExistBody = False And Left(occurrence.Name, 2) <> "41" And Left(occurrence.Name, 1) <> "6" Then

                ExistBody = True
                ComboBox7.Text = occurrence.Definition.Document.ActiveMaterial.DisplayName

            End If

        Next

        '>>>'>>>' Propersite '<<<'<<<'
        
        Dim iProperty As PropertySets
        Set iProperty = oDoc.ComponentDefinition.Document.PropertySets
        
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

        lisPram.ColumnWidths = "60;40"
        lisPram.Clear
        For i = 0 To counter
            lisPram.AddItem unitParametersNames(i)
            lisPram.List(i, 1) = unitParametersValues(i)
        Next

        CommandButton7.Enabled = True
        CommandButton8.Enabled = True

    Else
    
        ''' Get Unit Name '''

        unitName = ComboBoxAssembly.Text
        
        ''' Select Unit '''

        For Each oOcc In oDoc.ComponentDefinition.Occurrences

            If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
                If unitName = oOcc.Name Then

                    oDoc.SelectSet.Clear
                    oDoc.SelectSet.Select (oOcc)
                    Exit For

                End If
            End If

        Next

        ''' Get Part List in Assembly '''

        'ReDim partListAssembly(50) As String

        counter = 0
        For Each part In oOcc.Definition.Occurrences
            partListAssembly(counter) = part.Name
            counter = counter + 1
        Next

        ComboBox5.List = partListAssembly

        ''' Set Parameter For UserParameters '''

        shortUnitName = ComboBoxAssembly.Text
        shortUnitName = Left(shortUnitName, InStr(1, shortUnitName, ":") - 1)
        shortUnitName = Replace(shortUnitName, "-", "_")
        
        For Each param In oDoc.ComponentDefinition.Parameters.UserParameters
        
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
        
        Set userParams = oOcc.Definition.Parameters.UserParameters
        counter = 0

        For Each param In userParams
        
            If param.Name = "width" Then
                lbWidthAssembly.Caption = param.Expression
                
            ElseIf param.Name = "depth" Then
                lbDepthAssembly.Caption = param.Expression
                
            ElseIf param.Name = "height" Then
                lbHeightAssembly.Caption = param.Expression
            End If
            
            unitParametersNames(counter) = param.Name
            unitParametersValues(counter) = param.Value
            counter = counter + 1

        Next

        ComboBoxAssemblyParameters.List = unitParametersNames
        
        ComboBoxAssemblyParameters_Change
        
        '>>>'>>>' Material Frame '<<<'<<<'

        'Dim ExistDoor As Boolean
        'Dim ExistAft As Boolean
        'Dim ExistBody As Boolean

        For Each occurrence In oOcc.Definition.Occurrences

            ''' Door Material '''
            If ExistDoor = False And Left(occurrence.Name, 1) = "6" Then

                ExistDoor = True
                ComboBox6.Text = occurrence.Definition.Document.ActiveMaterial.DisplayName

            ''' Aft Material '''
            ElseIf ExistAft = False And Left(occurrence.Name, 2) = "41" Then

                ExistAft = True
                ComboBox9.Text = occurrence.Definition.Document.ActiveMaterial.DisplayName

            ''' Body Material '''
            ElseIf ExistBody = False And Left(occurrence.Name, 2) <> "41" And Left(occurrence.Name, 1) <> "6" Then

                ExistBody = True
                ComboBox7.Text = occurrence.Definition.Document.ActiveMaterial.DisplayName

            End If

        Next

        '>>>'>>>' Propersite '<<<'<<<'
        
        'Dim iProperty As PropertySets
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

        lisPram.ColumnWidths = "60;40"
        lisPram.Clear
        For i = 0 To counter
            lisPram.AddItem unitParametersNames(i)
            lisPram.List(i, 1) = unitParametersValues(i)
        Next

        CommandButton7.Enabled = True
        CommandButton8.Enabled = True
        
    End If

    ComboBox5.Text = ""

End Sub

''' Founctions '''

Sub SetFormola()
    
    Dim oOcc As ComponentOccurrence
    
    Dim partName, shortPartName, assemblyname As String

    partName = ComboBoxPart.Value
    shortPartName = Left(partName, 2)

    assemblyname = oDoc.DisplayName
    assemblyname = Replace(assemblyname, ".iam", "")

    For Each oOcc In oDoc.ComponentDefinition.Occurrences

        If partName = oOcc.Name Then
            Exit For
        End If

    Next
    
    Dim D_Pram As String
    D_Pram = txtboxPartDFormula.Text
    
    Dim WH_Pram As String
    WH_Pram = txtboxPartWHFormula.Text
    
    Dim oParameters As Parameters
    Dim userParams As UserParameters
    Set oParameters = oDoc.ComponentDefinition.Parameters
    Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters
    
    Dim param As Parameter
    Dim existD, existWH As Boolean
    existD = False
    existWH = False
    
    Dim setD As Boolean
    Dim setWH As Boolean
    
    Dim numStyle As String
    numStyle = CInt(Right(ComboBoxStyle.Text, 1))
    
    If numStyle = 1 Then
        For Each param In userParams
        
            If param.Name = "d_" + shortPartName Then
                existD = True
            ElseIf param.Name = "wh_" + shortPartName Then
                existWH = True
            End If
        
        Next
    ElseIf numStyle > 1 Then
        For Each param In userParams
        
            If param.Name = "s" & numStyle & "_d_" & shortPartName Then
                existD = True
            ElseIf param.Name = "s" & numStyle & "_wh_" & shortPartName Then
                existWH = True
            End If
            
        Next
    End If
    
    If WH_Pram <> "" Then
    
        If existWH = False Then
        
            If numStyle = 1 Then
                If oParameters.IsExpressionValid(WH_Pram, kCentimeterLengthUnits) = True Then
                    setWH = True
                    Set param = userParams.AddByExpression("wh_" + shortPartName, WH_Pram, kCentimeterLengthUnits)
                Else
                    Var = MsgBox("WH parameter is not valid", vbCritical, "Error")
                End If
            ElseIf numStyle > 1 Then
                If oParameters.IsExpressionValid(WH_Pram, kCentimeterLengthUnits) = True Then
                    setWH = True
                    Set param = userParams.AddByExpression("s" & numStyle & "_wh_" + shortPartName, WH_Pram, kCentimeterLengthUnits)
                Else
                    Var = MsgBox("WH parameter is not valid", vbCritical, "Error")
                End If
            End If
            
        Else
        
            If numStyle = 1 Then
                If oParameters.IsExpressionValid(WH_Pram, kCentimeterLengthUnits) = True Then
                    setWH = True
                    userParams.Item("wh_" + shortPartName).Expression = WH_Pram
                Else
                    Var = MsgBox("WH parameter is not valid", vbCritical, "Error")
                End If
            ElseIf numStyle > 1 Then
                If oParameters.IsExpressionValid(WH_Pram, kCentimeterLengthUnits) = True Then
                    setWH = True
                    userParams.Item("s" & numStyle & "_wh_" + shortPartName).Expression = WH_Pram
                Else
                    Var = MsgBox("WH parameter is not valid", vbCritical, "Error")
                End If
            End If
            
        End If
        
    End If
    
    If D_Pram <> "" Then

        If existD = False Then
        
            If numStyle = 1 Then
                If oParameters.IsExpressionValid(D_Pram, kCentimeterLengthUnits) = True Then
                    setD = True
                    Set param = userParams.AddByExpression("d_" + shortPartName, D_Pram, kCentimeterLengthUnits)
                Else
                    Var = MsgBox("D parameter is not valid", vbCritical, "Error")
                End If
            ElseIf numStyle > 1 Then
                If oParameters.IsExpressionValid(D_Pram, kCentimeterLengthUnits) = True Then
                    setD = True
                    Set param = userParams.AddByExpression("s" & numStyle & "_d_" + shortPartName, D_Pram, kCentimeterLengthUnits)
                Else
                    Var = MsgBox("WH parameter is not valid", vbCritical, "Error")
                End If
            End If
                
        Else
        
            If numStyle = 1 Then
                If oParameters.IsExpressionValid(D_Pram, kCentimeterLengthUnits) = True Then
                    setD = True
                    userParams.Item("d_" + shortPartName).Expression = D_Pram
                Else
                    Var = MsgBox("D parameter is not valid", vbCritical, "Error")
                End If
            ElseIf numStyle > 1 Then
                If oParameters.IsExpressionValid(D_Pram, kCentimeterLengthUnits) = True Then
                    setD = True
                    userParams.Item("s" & numStyle & "_d_" + shortPartName).Expression = D_Pram
                Else
                    Var = MsgBox("D parameter is not valid", vbCritical, "Error")
                End If
            End If
            
        End If

    End If
    
    Dim oParameter As Parameters
    setProperty = False
    
    Dim materialExist As Boolean
    Dim material As MaterialAsset
    
    For Each material In oDoc.Assets.Application.AssetLibraries.Item("Inventor Material Library").MaterialAssets
        
        If ComboBoxMaterialPart.Text = material.DisplayName Then
            materialExist = True
            Exit For
        End If
        
    Next
    
    If materialExist = False Then
        Var = MsgBox("Selected Material Don't Exist !", vbInformation, "Error")
    End If
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences

        If oOcc.Name = ComboBoxPart.Text Then

            Set oParameter = oOcc.Definition.Parameters
            If setD = True Then
            
                If numStyle = 1 Then
                    oParameter.Item("D").Expression = userParams("d_" + shortPartName).Value
                    txtboxPartDValue.Text = oParameter.Item("D").Expression
                ElseIf numStyle > 1 Then
                    oParameter.Item("D").Expression = userParams("s" & numStyle & "_d_" + shortPartName).Value
                    txtboxPartDValue.Text = oParameter.Item("D").Expression
                End If
                
            End If
            
            If setWH = True Then
            
                If numStyle = 1 Then
                    oParameter.Item("WH").Expression = userParams("wh_" + shortPartName).Value
                    txtboxPartWHValue.Text = oParameter.Item("WH").Expression
                ElseIf numStyle > 1 Then
                    oParameter.Item("WH").Expression = userParams("s" & numStyle & "_wh_" + shortPartName).Value
                    txtboxPartWHValue.Text = oParameter.Item("WH").Expression
                End If
                
            End If
            
            If materialExist = True Then
                oOcc.Definition.Document.ActiveMaterial = material
            End If
            
            Exit For
        End If
    Next

    Dim partPramUsers As UserParameters
    Set partPramUsers = oOcc.Definition.Parameters.UserParameters
    
    If ComboBoxD1.Text <> "" Then
        partPramUsers.Item("W1").Value = ComboBoxD1.Text
    End If
    If ComboBoxD2.Text <> "" Then
        partPramUsers.Item("W2").Value = ComboBoxD2.Text
    End If
    If ComboBoxWH1.Text <> "" Then
        partPramUsers.Item("L1").Value = ComboBoxWH1.Text
    End If
    If ComboBoxWH2.Text <> "" Then
        partPramUsers.Item("L2").Value = ComboBoxWH2.Text
    End If
    
    If txtboxCostMaterialPart.Text <> "" Then
        oOcc.Definition.Document.ActiveMaterial.Item(4).Value = CInt(txtboxCostMaterialPart.Text)
    End If

    ''' Update Assembly
    
    ThisApplication.ActiveDocument.Update
    
End Sub

Sub SetFormolaAssembly()
    
    Dim oOcc As ComponentOccurrence
    Dim Sub_oOcc As ComponentOccurrence
    
    Dim unitName, shortUnitName As String
    unitName = ComboBoxAssembly.Value
    shortUnitName = Left(unitName, InStr(1, unitName, ":") - 1)
    shortUnitName = Replace(shortUnitName, "-", "_")
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences

        If unitName = oOcc.Name Then
            Exit For
        End If

    Next
    
    Dim Width_Pram As String
    Width_Pram = TextBox3.Text
    
    Dim Depth_Pram As String
    Depth_Pram = TextBox4.Text
    
    Dim Height_Pram As String
    Height_Pram = TextBox5.Text

    Dim existWidth As Boolean
    Dim existDepth As Boolean
    Dim existHeight As Boolean

    Dim setWidth As Boolean
    Dim setDepth As Boolean
    Dim setHeight As Boolean

    Dim param As Parameter
    
    Dim userParams As UserParameters
    Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters
    
    For Each param In userParams
        
        If param.Name = "width" + "_" + shortUnitName Then
            existWidth = True
                
        ElseIf param.Name = "depth" + "_" + shortUnitName Then
            existDepth = True
                
        ElseIf param.Name = "height" + "_" + shortUnitName Then
            existHeight = True
            
        End If
            
    Next
    
    Dim oParameters As Parameters
    Set oParameters = oDoc.ComponentDefinition.Parameters
    
    If Width_Pram <> "" Then
        If existWidth = False Then
        
            If oParameters.IsExpressionValid(Width_Pram, kCentimeterLengthUnits) = True Then
                setWidth = True
                Set param = userParams.AddByExpression("width" + "_" + shortUnitName, Width_Pram, kCentimeterLengthUnits)
            Else
                Var = MsgBox("Width parameter is not valid", vbCritical, "Error")
            End If
            
        Else
        
            If oParameters.IsExpressionValid(Width_Pram, kCentimeterLengthUnits) = True Then
                setWidth = True
                userParams.Item("width" + "_" + shortUnitName).Expression = Width_Pram
            Else
                Var = MsgBox("Width parameter is not valid", vbCritical, "Error")
            End If
            
        End If
        
    End If
    
    If Depth_Pram <> "" Then
        If existDepth = False Then
            If oParameters.IsExpressionValid(Depth_Pram, kCentimeterLengthUnits) = True Then
                setDepth = True
                Set param = userParams.AddByExpression("depth" + "_" + shortUnitName, Depth_Pram, kCentimeterLengthUnits)
            Else
                Var = MsgBox("Depth parameter is not valid", vbCritical, "Error")
            End If
        Else
            If oParameters.IsExpressionValid(Depth_Pram, kCentimeterLengthUnits) = True Then
                setDepth = True
                userParams.Item("depth" + "_" + shortUnitName).Expression = Depth_Pram
            Else
                Var = MsgBox("Depth parameter is not valid", vbCritical, "Error")
            End If
        End If
        
        
    End If
    
    If Height_Pram <> "" Then
        If existHeight = False Then
            If oParameters.IsExpressionValid(Height_Pram, kCentimeterLengthUnits) = True Then
                setHeight = True
                Set param = userParams.AddByExpression("height" + "_" + shortUnitName, Height_Pram, kCentimeterLengthUnits)
            Else
                Var = MsgBox("Height parameter is not valid", vbCritical, "Error")
            End If
        Else
            If oParameters.IsExpressionValid(Height_Pram, kCentimeterLengthUnits) = True Then
                setHeight = True
                userParams.Item("height" + "_" + shortUnitName).Expression = Height_Pram
            Else
                Var = MsgBox("Height parameter is not valid", vbCritical, "Error")
            End If
        End If
        
    End If
    
    Dim oParameter As Parameters
    setProperty = False
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        If oOcc.Name = ComboBoxAssembly.Text Then
        
            Set oParameter = oOcc.Definition.Parameters
            
            If setWidth = True Then
                oParameter.Item("width").Expression = userParams.Item("width" + "_" + shortUnitName).Value
                lbWidthAssembly.Caption = oParameter.Item("width").Expression
            End If
            
            If setDepth = True Then
                oParameter.Item("depth").Expression = userParams.Item("depth" + "_" + shortUnitName).Value
                lbDepthAssembly.Caption = oParameter.Item("depth").Expression
            End If
            
            If setHeight = True Then
                oParameter.Item("height").Expression = userParams.Item("height" + "_" + shortUnitName).Value
                lbHeightAssembly.Caption = oParameter.Item("height").Expression
            End If
            
            Exit For
        End If
        
    Next
    
    ''' Set Value For User Parameters '''
    
    If CheckBox6.Value = True And TextBox11.Text <> "" Then

        Dim paramName As String
        paramName = ComboBoxAssemblyParameters.Text
        
        If oParameter.IsExpressionValid(TextBox11.Text, oParameter.Item(paramName).Units) = True Then
            oParameter.Item(paramName).Expression = TextBox11.Text
        Else
            Var = MsgBox(paramName & " parameter is not valid", vbCritical, "Error")
        End If

    End If

    ''' Change Materials '''

    Dim occurrence As ComponentOccurrence

    Dim materials As AssetsEnumerator
    Set materials = oDoc.Assets.Application.AssetLibraries.Item("Inventor Material Library").MaterialAssets
    
    Dim AftMaterialExist As Boolean
    Dim DoorMaterialExist As Boolean
    Dim BodyMaterialExist As Boolean
    Dim SelectedMaterialExist As Boolean

    For Each occurrence In oOcc.Definition.Occurrences

        If Left(occurrence.Name, 1) = "6" Then                  ''' Door Material

            If CheckBox7.Value = True And ComboBox6.Text <> "" Then
                For Each materialDoor In materials
                    if
                Next
                occurrence.Definition.Document.ActiveMaterial = materials.Item(ComboBox6.Text)
            End If

        ElseIf Left(occurrence.Name, 2) = "41" Then             ''' Aft Material

            If CheckBox12.Value = True And ComboBox9.Text <> "" Then
                CheckBox12.Value = False
                occurrence.Definition.Document.ActiveMaterial = materials.Item(ComboBox9.Text)
            End If

        ElseIf occurrence.Name = ComboBox5.Text Then            ''' Selected Material

            If CheckBox9.Value = True And ComboBox8.Text <> "" Then
                CheckBox9.Value = False
                occurrence.Definition.Document.ActiveMaterial = materials.Item(ComboBox8.Text)
            End If

        Else                                                    ''' Body Material

            If CheckBox8.Value = True And ComboBox7.Text <> "" Then
                occurrence.Definition.Document.ActiveMaterial = materials.Item(ComboBox7.Text)
            End If

        End If

    Next
    
    ''' Set Subject Unit '''
    
    If CheckBox10.Value = True And TextBox12.Text <> "" Then
        
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
    
    If CheckBox11.Value = True And TextBox13.Text <> "" Then
        
        Dim iPropertyManeage As PropertySets
        Set iPropertyManeage = oOcc.Definition.Document.PropertySets
        iPropertyManeage.Item(2).Item(2).Expression = TextBox13.Text
        
        For Each Sub_oOcc In oOcc.Definition.Occurrences
            Set iPropertyManeage = Sub_oOcc.Definition.Document.PropertySets
            iPropertyManeage.Item(2).Item(2).Expression = TextBox13.Text
        Next
        
        Set iPropertyManeage = oOcc.Definition.Document.PropertySets
        Label26.Caption = iPropertyManeage.Item(3).Item(2).Value
        
    End If
      
    ThisApplication.ActiveDocument.Update
    
End Sub

Sub SetFormolaAssemblyMaster()

    Dim oOcc As ComponentOccurrence
    Dim Sub_oOcc As ComponentOccurrence
    
    Dim Width_Pram As String
    Width_Pram = TextBox3.Text
    
    Dim Depth_Pram As String
    Depth_Pram = TextBox4.Text
    
    Dim Height_Pram As String
    Height_Pram = TextBox5.Text

    Dim setWidth As Boolean
    Dim setDepth As Boolean
    Dim setHeight As Boolean

    Dim param As Parameter
    
    Dim userParams As UserParameters
    Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters
    
    Dim oParameters As Parameters
    Set oParameters = oDoc.ComponentDefinition.Parameters
    
    If Width_Pram <> "" Then
        
        If oParameters.IsExpressionValid(Width_Pram, kCentimeterLengthUnits) = True Then
            setWidth = True
            userParams.Item("width").Expression = Width_Pram
        Else
            Var = MsgBox("Width parameter is not valid", vbCritical, "Error")
        End If

    End If
    
    If Depth_Pram <> "" Then
        
        If oParameters.IsExpressionValid(Depth_Pram, kCentimeterLengthUnits) = True Then
            setDepth = True
            userParams.Item("depth").Expression = Depth_Pram
        Else
            Var = MsgBox("Depth parameter is not valid", vbCritical, "Error")
        End If
        
    End If
    
    If Height_Pram <> "" Then
        
        If oParameters.IsExpressionValid(Height_Pram, kCentimeterLengthUnits) = True Then
            setHeight = True
            userParams.Item("height").Expression = Height_Pram
        Else
            Var = MsgBox("Height parameter is not valid", vbCritical, "Error")
        End If

    End If

    lbWidthAssembly.Caption = userParams.Item("width").Expression
    lbDepthAssembly.Caption = userParams.Item("depth").Expression
    lbHeightAssembly.Caption = userParams.Item("height").Expression
    
    ''' Set Value For User Parameters '''
    
    If CheckBox6.Value = True Then

        Dim paramName As String
        paramName = ComboBoxAssemblyParameters.Text
        userParams.Item(paramName).Expression = TextBox11.Text

    End If

    ''' Change Materials '''

    Dim occurrence As ComponentOccurrence

    Dim materials As AssetsEnumerator
    Set materials = oDoc.Assets.Application.AssetLibraries.Item("Inventor Material Library").MaterialAssets

    For Each occurrence In oDoc.ComponentDefinition.Occurrences

        If Left(occurrence.Name, 1) = "6" Then                  ''' Door Material

            If CheckBox7.Value = True And ComboBox6.Text <> "" Then
                occurrence.Definition.Document.ActiveMaterial = materials.Item(ComboBox6.Text)
            End If

        ElseIf Left(occurrence.Name, 2) = "41" Then             ''' Aft Material

            If CheckBox12.Value = True And ComboBox9.Text <> "" Then
                CheckBox12.Value = False
                occurrence.Definition.Document.ActiveMaterial = materials.Item(ComboBox9.Text)
            End If

        ElseIf occurrence.Name = ComboBox5.Text Then            ''' Selected Material

            If CheckBox9.Value = True And ComboBox8.Text <> "" Then
                CheckBox9.Value = False
                occurrence.Definition.Document.ActiveMaterial = materials.Item(ComboBox8.Text)
            End If

        Else                                                    ''' Body Material

            If CheckBox8.Value = True And ComboBox7.Text <> "" Then
                occurrence.Definition.Document.ActiveMaterial = materials.Item(ComboBox7.Text)

            End If

        End If

    Next
    
    ''' Set Subject Unit '''
    
    If CheckBox10.Value = True And TextBox12.Text <> "" Then
        
        Dim iPropertySubject As PropertySets
        Set iPropertySubject = oDoc.ComponentDefinition.Document.PropertySets
        iPropertySubject.Item(1).Item(2).Expression = TextBox12.Text
        Label26.Caption = iPropertySubject.Item(3).Item(2).Value
        
        For Each Sub_oOcc In oDoc.ComponentDefinition.Occurrences
            Set iPropertySubject = Sub_oOcc.Definition.Document.PropertySets
            iPropertySubject.Item(1).Item(2).Expression = TextBox12.Text
        Next
        
    End If
    
    ''' Set Maneage Unit '''
    
    If CheckBox11.Value = True And TextBox13.Text <> "" Then
        
        Dim iPropertyManeage As PropertySets
        Set iPropertyManeage = oDoc.ComponentDefinition.Document.PropertySets
        iPropertyManeage.Item(2).Item(2).Expression = TextBox13.Text
        
        For Each Sub_oOcc In oDoc.ComponentDefinition.Occurrences
            Set iPropertyManeage = Sub_oOcc.Definition.Document.PropertySets
            iPropertyManeage.Item(2).Item(2).Expression = TextBox13.Text
        Next
        
        Set iPropertyManeage = oDoc.ComponentDefinition.Document.PropertySets
        Label26.Caption = iPropertyManeage.Item(3).Item(2).Value
        
    End If
    
      
    ThisApplication.ActiveDocument.Update
    
End Sub

Sub selectIteam()
    
    If oDoc.SelectSet.Count = 1 Then
        If oDoc.SelectSet.Item(1).Type <> kAssemblyComponentDefinitionObject Then
        
            If oDoc.SelectSet.Item(1).Type = kRectangularOccurrencePatternObject Then
                
                Dim oOcc As ComponentOccurrence
                Set oOcc = oDoc.SelectSet.Item(1).OccurrencePatternElements.Item(1).Components.Item(1)
                
                If oOcc.DefinitionDocumentType = kPartDocumentObject Then
                    isUnit = False
                    MultiPage1.Value = 0
                    ComboBoxPart.Text = oOcc.Name
                ElseIf oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
                    isUnit = False
                    MultiPage1.Value = 1
                    ComboBoxAssembly.Text = oOcc.Name
                End If
                
            Else
            
                If oDoc.SelectSet.Item(1).DefinitionDocumentType = kPartDocumentObject Then
                    isUnit = False
                    MultiPage1.Value = 0
                    ComboBoxPart.Text = oDoc.SelectSet.Item(1).Name
                ElseIf oDoc.SelectSet.Item(1).DefinitionDocumentType = kAssemblyDocumentObject Then
                    isUnit = False
                    MultiPage1.Value = 1
                    ComboBoxAssembly.Text = oDoc.SelectSet.Item(1).Name
                End If
                    
            End If
            
        End If
    End If
    
    
    
End Sub

Sub ResizePages()

    If MultiPage1.Value = 0 Then        '''Part Page

        MultiPage1.Width = 270
        MultiPage1.Height = 452
        lisPram.Left = 280
        lisPram.Height = 435
        lbLisPram.Left = 300
        Set_Formula_Form2.Width = 428
        Set_Formula_Form2.Height = 490
        ToggleMore.Visible = False

    ElseIf MultiPage1.Value = 1 Then    '''Assembly Page

        MultiPage1.Width = 306
        MultiPage1.Height = 462
        lisPram.Left = 318
        lisPram.Height = 410
        lbLisPram.Left = 336
        Set_Formula_Form2.Width = 465
        Set_Formula_Form2.Height = 500
        ToggleMore.Visible = True

    ElseIf MultiPage1.Value = 2 Then    '''Kichen Page

        MultiPage1.Width = 234
        MultiPage1.Height = 342
        Set_Formula_Form2.Width = 256
        Set_Formula_Form2.Height = 380

    End If

End Sub

Sub RotateAllDoor()

    Dim oAppearance As Asset
    Dim oOcc, part As ComponentOccurrence
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        For Each part In oOcc.Definition.Occurrences
            
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

                Dim oParameters As Parameters
                Set oParameters = oOcc.Definition.Parameters
                
                ''' Get the parameter named "D".
                Dim oDParam As Parameter
                Set oDParam = oParameters.Item("D")
                oDParam.Name = "WH2"
                
                ''' Get the parameter named "WH".
                Dim oWHParam As Parameter
                Set oWHParam = oParameters.Item("WH")
                oWHParam.Name = "D"
                
                oDParam.Name = "WH"

            End If
        Next
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub

Sub MoreParameters(ByVal keyParametersCount As Integer)
    
    If keyParametersCount > 0 Then
        lbMore1.Visible = True
        lbMore1.Caption = keyParametersName(0)
        txtMore1.Visible = True
        txtMore1.Text = keyParametersValue(0)
    End If

    If keyParametersCount > 1 Then
        lbMore2.Visible = True
        lbMore2.Caption = keyParametersName(1)
        txtMore2.Visible = True
        txtMore2.Text = keyParametersValue(1)
    End If

    If keyParametersCount > 2 Then
        lbMore3.Visible = True
        lbMore3.Caption = keyParametersName(2)
        txtMore3.Visible = True
        txtMore3.Text = keyParametersValue(2)
    End If

    If keyParametersCount > 3 Then
        lbMore4.Visible = True
        lbMore4.Caption = keyParametersName(3)
        txtMore4.Visible = True
        txtMore4.Text = keyParametersValue(3)
    End If

    If keyParametersCount > 4 Then
        lbMore5.Visible = True
        lbMore5.Caption = keyParametersName(4)
        txtMore5.Visible = True
        txtMore5.Text = keyParametersValue(4)
    End If

    If keyParametersCount > 5 Then
        lbMore6.Visible = True
        lbMore6.Caption = keyParametersName(5)
        txtMore6.Visible = True
        txtMore6.Text = keyParametersValue(5)
    End If

    If keyParametersCount > 6 Then
        lbMore7.Visible = True
        lbMore7.Caption = keyParametersName(6)
        txtMore7.Visible = True
        txtMore7.Text = keyParametersValue(6)
    End If

    If keyParametersCount > 7 Then
        lbMore8.Visible = True
        lbMore8.Caption = keyParametersName(7)
        txtMore8.Visible = True
        txtMore8.Text = keyParametersValue(7)
    End If

    If keyParametersCount > 8 Then
        lbMore9.Visible = True
        lbMore9.Caption = keyParametersName(8)
        txtMore9.Visible = True
        txtMore9.Text = keyParametersValue(8)
    End If

End Sub

Sub CheckIsUnit()

    For Each param In oDoc.ComponentDefinition.Parameters.UserParameters
        If param.Name = "Unit" And param.Value = True Then
            isUnit = True
            MultiPage1.Value = 1
            Exit For
        End If
    Next

    If isUnit = True Then
        ComboBoxAssembly_Change
    End If

End Sub

''' CheckBox Events '''

Private Sub CheckBox1_Change()

    Dim oOcc As ComponentOccurrence
    Dim Sub_oOcc As ComponentOccurrence
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        If Left(oOcc.Name, 1) = "6" Then
            oOcc.Visible = CheckBox1.Value
        End If
        
        For Each Sub_oOcc In oOcc.Definition.Occurrences
            
            If Left(Sub_oOcc.Name, 1) = "6" Then
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
    
        If Left(oOcc.Name, 2) = "41" Then
            oOcc.Visible = CheckBox2.Value
        End If
        
        For Each Sub_oOcc In oOcc.Definition.Occurrences
            If Left(Sub_oOcc.Name, 2) = "41" Then
                Sub_oOcc.Visible = CheckBox2.Value
            End If
        Next
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub

''' TextBox & txtbox Events '''

Private Sub TextBox11_Change()
    
    If started = True Then
        CheckBox6.Value = 1
    End If

End Sub

Private Sub TextBox12_Enter()

    CheckBox10.Value = True

End Sub

Private Sub TextBox13_Change()

    If started = True Then
        CheckBox11.Value = 1
    End If

End Sub

Private Sub TextBox3_Enter()

    CheckBox3.Value = True
    CheckBox4.Value = False
    CheckBox5.Value = False

End Sub

Private Sub TextBox4_Enter()

    CheckBox3.Value = False
    CheckBox4.Value = True
    CheckBox5.Value = False

End Sub

Private Sub TextBox5_Enter()

    CheckBox3.Value = False
    CheckBox4.Value = False
    CheckBox5.Value = True

End Sub

Private Sub txtboxPartDFormula_Enter()

    CheckBox13.Value = True
    CheckBox14.Value = False

End Sub

Private Sub txtboxPartWHFormula_Enter()

    CheckBox13.Value = False
    CheckBox14.Value = True

End Sub

Private Sub lisPram_Click()
    
    If MultiPage1.Value = 1 And (isUnit = True Or ComboBoxAssembly.Text <> "") Then
        ComboBoxAssemblyParameters.Text = lisPram.Text
    End If

End Sub

Private Sub lisPram_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If MultiPage1.Value = 0 Then
    
        If CheckBox13.Value = True Then
            txtboxPartDFormula.Text = txtboxPartDFormula.Text + lisPram.Text
        ElseIf CheckBox14.Value = True Then
            txtboxPartWHFormula.Text = txtboxPartWHFormula.Text + lisPram.Text
        End If
        
    ElseIf MultiPage1.Value = 1 Then
    
        If CheckBox3.Value = True Then
            TextBox3.Text = TextBox3.Text + lisPram.Text
        ElseIf CheckBox4.Value = True Then
            TextBox4.Text = TextBox4.Text + lisPram.Text
        ElseIf CheckBox5.Value = True Then
            TextBox5.Text = TextBox5.Text + lisPram.Text
        End If
        
    End If
    
End Sub

Private Sub MultiPage1_Change()

    ResizePages

End Sub

Private Sub ToggleMore_Click()

    If ToggleMore.Value = True Then

        Set_Formula_Form2.Width = 621
        ToggleMore.Caption = "More <<"

    Else

        Set_Formula_Form2.Width = 465
        ToggleMore.Caption = "More >>"

    End If
    
End Sub
