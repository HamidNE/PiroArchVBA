VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Set_Formula_Form 
   Caption         =   "Set Formula"
   ClientHeight    =   5640
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7188
   OleObjectBlob   =   "Set_Formula_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Set_Formula_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim oDoc As AssemblyDocument

Dim partnameArray(100) As String
Dim assemblyNameArray(30) As String
Dim parametersArray(100) As String

Private Sub ComboBox1_Change()

    If ComboBox1.Text = "" Then
        TextBox1.Text = ""
        TextBox2.Text = ""
        
        OptionButton1.Value = False
        OptionButton2.Value = False
    
        Frame2.Enabled = False
        TextBox1.BackColor = &H80000004
        TextBox2.BackColor = &H80000004
        
        Label2.ForeColor = &H80000006
        Label3.ForeColor = &H80000006
    Else
    
        Dim partName, assemblyname As String
        partName = ComboBox1.Value
        assemblyname = oDoc.DisplayName
        
        Dim i As Integer
        i = 1
        For Each oOcc In oDoc.ComponentDefinition.Occurrences
            If oOcc.DefinitionDocumentType = kPartDocumentObject Then
            
                If partName = oOcc.Name Then
                    oDoc.SelectSet.Clear
                    oDoc.SelectSet.Select (oOcc)
                    ExistPart = True
                    Exit For
                End If
                i = i + 1
                
            End If
        Next
        
        Dim componentOcc As ComponentOccurrences
        Set componentOcc = oDoc.ComponentDefinition.Occurrences
        
        Dim userParams As userParameters
        Set userParams = oDoc.ComponentDefinition.Parameters.userParameters
        
        Dim param As Parameter
        Dim existD, existWH As Boolean
        existD = False
        existWH = False
        
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
                TextBox1.Text = param.Expression
                existD = True
                Exit For
            End If
        Next
        
        For Each param In userParams
            If param.Name = "wh_" + partName Then
                TextBox2.Text = param.Expression
                existWH = True
                Exit For
            End If
        Next
        
        If existD = False Then
            TextBox1.Text = ""
        End If
        
        If existWH = False Then
            TextBox2.Text = ""
        End If
        
        Frame2.Enabled = True
        TextBox1.BackColor = &H80000005
        TextBox2.BackColor = &H80000005
        
        Label2.ForeColor = &H80000012
        Label3.ForeColor = &H80000012
        
    End If
End Sub

Private Sub ComboBox1_Enter()

End Sub

Private Sub ComboBox2_Change()

    If ComboBox2.Text = "" Then
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        
        OptionButton3.Value = False
        OptionButton4.Value = False
        OptionButton5.Value = False
    
        Frame4.Enabled = False
        TextBox3.BackColor = &H80000004
        TextBox4.BackColor = &H80000004
        TextBox5.BackColor = &H80000004
        
        Label7.ForeColor = &H80000006
        Label8.ForeColor = &H80000006
        Label9.ForeColor = &H80000006
    Else
        
        Dim unitName As String
        unitName = ComboBox2.Text
        
        Dim i As Integer
        i = 1
        For Each oOcc In oDoc.ComponentDefinition.Occurrences
            If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
            
                If unitName = oOcc.Name Then
                    oDoc.SelectSet.Clear
                    oDoc.SelectSet.Select (oOcc)
                    ExistPart = True
                    Exit For
                End If
                i = i + 1
                
            End If
        Next
        
        Dim componentOcc As ComponentOccurrences
        Set componentOcc = oDoc.ComponentDefinition.Occurrences
        
        Dim userParams As userParameters
        Set userParams = oDoc.ComponentDefinition.Parameters.userParameters
        
        Dim param As Parameter
        Dim existWidth, existDepth, existHeight As Boolean
        existWidth = False
        existDepth = False
        existHeight = False
        
        For Each param In userParams
        
            If param.Name = "width" + "_" + unitName Then
                TextBox3.Text = param.Expression
                existWidth = True
                
            ElseIf param.Name = "depth" + "_" + unitName Then
                TextBox4.Text = param.Expression
                existDepth = True
                
            ElseIf param.Name = "height" + "_" + unitName Then
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
        
        Frame4.Enabled = True
        TextBox3.BackColor = &H80000005
        TextBox4.BackColor = &H80000005
        TextBox5.BackColor = &H80000005
        
        Label7.ForeColor = &H80000012
        Label8.ForeColor = &H80000012
        Label9.ForeColor = &H80000012
        
    End If
End Sub

Sub SetFormola()
    
    'Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    
    Dim partName, assemblyname As String
    partName = ComboBox1.Value
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
    D_Pram = TextBox1.Text
    
    Dim WH_Pram As String
    WH_Pram = TextBox2.Text
    
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
            Set param = userParams.AddByExpression("d_" + partName, D_Pram, kMillimeterLengthUnits)
        Else
            userParams.Item("d_" + partName).Expression = D_Pram
        End If
        
        setD = True
    End If
    
    If WH_Pram <> "" Then
        If existWH = False Then
            Set param = userParams.AddByExpression("wh_" + partName, WH_Pram, kMillimeterLengthUnits)
        Else
            userParams.Item("wh_" + partName).Expression = WH_Pram
        End If
        
        setWH = True
    End If
    
    Dim oParameter As Parameters
    setProperty = False
    
    Dim partnametemp As String
    partnametemp = assemblyname + "-" + partName
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        If oOcc.Name = partnametemp Then

            Set oParameter = oOcc.Definition.Parameters
            If setD = True Then
                oParameter.Item("D").Expression = userParams("d_" + partName).Value * 10
            End If
            If setWH = True Then
                oParameter.Item("WH").Expression = userParams("wh_" + partName).Value * 10
            End If

            Exit For
        End If
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub

Sub SetFormolaAssembly()
    
    'Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    
    Dim unitName As String
    unitName = ComboBox2.Value
    
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
    Dim existWidth, existDepth, existHeight As Boolean
    Dim setWidth, setDepth, setHeight As Boolean
    
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
            Set param = userParams.AddByExpression("width" + "_" + unitName, Width_Pram, kMillimeterLengthUnits)
        Else
            userParams.Item("width" + "_" + unitName).Expression = Width_Pram
        End If
        
        setWidth = True
    End If
    
    If Depth_Pram <> "" Then
        If existDepth = False Then
            Set param = userParams.AddByExpression("depth" + "_" + unitName, Depth_Pram, kMillimeterLengthUnits)
        Else
            userParams.Item("depth" + "_" + unitName).Expression = Depth_Pram
        End If
        
        setDepth = True
    End If
    
    If Height_Pram <> "" Then
        If existHeight = False Then
            Set param = userParams.AddByExpression("height" + "_" + unitName, Height_Pram, kMillimeterLengthUnits)
        Else
            userParams.Item("height" + "_" + unitName).Expression = Height_Pram
        End If
        
        setHeight = True
    End If
    
    Dim oParameter As Parameters
    setProperty = False
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        If oOcc.Name = unitName Then
        
            Set oParameter = oOcc.Definition.Parameters
            
            If setWidth = True Then
                oParameter.Item("width").Expression = userParams("width" + "_" + unitName).Value
            End If
            
            If setDepth = True Then
                oParameter.Item("depth").Expression = userParams("depth" + "_" + unitName).Value
            End If
            
            If setHeight = True Then
                oParameter.Item("height").Expression = userParams("height" + "_" + unitName).Value
            End If
            
            Exit For
        End If
        
    Next
      
    ThisApplication.ActiveDocument.Update
    
End Sub

Private Sub CommandButton1_Click()

    SetFormola
    
End Sub

Private Sub CommandButton2_Click()

    SetFormola
    Unload Me
    
End Sub

Sub selectIteam()
    
    If oDoc.SelectSet.count = 1 Then
    
        If oDoc.SelectSet.Item(1).DefinitionDocumentType = kPartDocumentObject Then
            MultiPage1.Value = 0
            ComboBox1.Text = oDoc.SelectSet.Item(1).Name
            ComboBox1_Change
        ElseIf oDoc.SelectSet.Item(1).DefinitionDocumentType = kAssemblyDocumentObject Then
            MultiPage1.Value = 1
            ComboBox2.Text = oDoc.SelectSet.Item(1).Name
            ComboBox2_Change
        
        End If
    End If
    
End Sub

Private Sub CommandButton5_Click()
    selectIteam
End Sub

Private Sub CommandButton6_Click()
    selectIteam
End Sub

Private Sub CommandButton7_Click()

    SetFormolaAssembly
    
End Sub

Private Sub CommandButton8_Click()

    SetFormolaAssembly
    Unload Me
    
End Sub

Private Sub CommandButton3_Click()
    Dim text1, text2 As String
    text1 = TextBox1.Text
    text2 = TextBox2.Text
    
    TextBox1.Text = text2
    TextBox2.Text = text1
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If OptionButton1.Value = True Then
        TextBox1.Text = TextBox1.Text + ListBox1.Text
    ElseIf OptionButton2.Value = True Then
        TextBox2.Text = TextBox2.Text + ListBox1.Text
    End If
    
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If OptionButton3.Value = True Then
        TextBox3.Text = TextBox3.Text + ListBox2.Text
    ElseIf OptionButton4.Value = True Then
        TextBox4.Text = TextBox4.Text + ListBox2.Text
    ElseIf OptionButton5.Value = True Then
        TextBox5.Text = TextBox5.Text + ListBox2.Text
    End If
    
End Sub

Private Sub TextBox1_Enter()
    OptionButton1.Value = True
    OptionButton2.Value = False
End Sub

Private Sub TextBox2_Enter()
    OptionButton1.Value = False
    OptionButton2.Value = True
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

Private Sub UserForm_Activate()
    
    'Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    Dim oOcc As ComponentOccurrence
    
    Dim count As Integer
    count = 0
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        If oOcc.DefinitionDocumentType = kPartDocumentObject Then
            partnameArray(count) = oOcc.Name
            count = count + 1
        End If
        
    Next
    
    count = 0
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
            assemblyNameArray(count) = oOcc.Name
            count = count + 1
        End If
        
    Next
    
    Dim param As Parameter
    Dim parametersLenght As Integer
    parametersLenght = 0
    
    For Each param In oDoc.ComponentDefinition.Parameters.userParameters
        If Left(param.Name, 2) <> "d_" And Left(param.Name, 3) <> "wh_" Then
            parametersArray(parametersLenght) = param.Name
            parametersLenght = parametersLenght + 1
        End If
        
    Next
    
    Dim Tprams As ParameterTable
    Dim Tpram As TableParameter
    
    For Each Tprams In oDoc.ComponentDefinition.Parameters.ParameterTables
    
        For Each Tpram In Tprams.TableParameters
            parametersArray(parametersLenght) = Tpram.Name
            parametersLenght = parametersLenght + 1
        Next
        
    Next
   
    ComboBox1.List = partnameArray
    ComboBox2.List = assemblyNameArray
    ListBox1.List = parametersArray
    ListBox2.List = parametersArray
    
    If oDoc.SelectSet.count = 1 Then
    
        If oDoc.SelectSet.Item(1).DefinitionDocumentType = kPartDocumentObject Then
            MultiPage1.Value = 0
            ComboBox1.Text = oDoc.SelectSet.Item(1).Name
            ComboBox1_Change
        ElseIf oDoc.SelectSet.Item(1).DefinitionDocumentType = kAssemblyDocumentObject Then
            MultiPage1.Value = 1
            ComboBox2.Text = oDoc.SelectSet.Item(1).Name
            ComboBox2_Change
        
        End If
    End If
End Sub

