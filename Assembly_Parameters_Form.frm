VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Assembly_Parameters_Form 
   Caption         =   "Assembly Parameters"
   ClientHeight    =   9345.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5496
   OleObjectBlob   =   "Assembly_Parameters_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Assembly_Parameters_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim oDoc As AssemblyDocument

Dim unitNameArray(50) As String
Dim unitParametersArray(100) As String
Dim location As Integer

Sub setParameters()
    
    If location > 0 Then
        Dim oOcc As ComponentOccurrence
        Set oOcc = oDoc.ComponentDefinition.Occurrences.Item(location)
        Dim iProperty As PropertySets
        Set iProperty = oOcc.Definition.Document.PropertySets
    
        If TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox3.Text <> "" Then
            
            Dim userParameters As userParameters
            Set userParameters = oOcc.Definition.Parameters.userParameters
            
            userParameters.Item("width").Expression = TextBox1.Text
            userParameters.Item("depth").Expression = TextBox2.Text
            userParameters.Item("height").Expression = TextBox3.Text
            
            If ComboBox2.Text <> "" And TextBox4.Text <> "" Then
                userParameters.Item(ComboBox2.Text).Expression = TextBox4.Text
            End If
    
        End If
        
        If TextBox5.Text <> "" Then
            iProperty.Item(1).Item(2).Expression = TextBox5.Text
        End If
        
        If TextBox6.Text <> "" Then
            iProperty.Item(2).Item(2).Expression = TextBox6.Text
        End If
        
        TextBox7.Text = iProperty.Item(3).Item(2).Expression + " ( " + iProperty.Item(3).Item(2).Value + " )"
        
        Dim Sub_oOcc As ComponentOccurrence
                
        For Each Sub_oOcc In oOcc.Definition.Occurrences
            Set iProperty = Sub_oOcc.Definition.Document.PropertySets
            
            If TextBox5.Text <> "" Then
                iProperty.Item(1).Item(2).Expression = TextBox5.Text
            End If
        
            If TextBox6.Text <> "" Then
                iProperty.Item(2).Item(2).Expression = TextBox6.Text
            End If
            
            iProperty.Item(3).Item(2).Expression = "=<Subject><Manager>.<title>"
        Next
        
        ThisApplication.ActiveDocument.Update
    End If
    
End Sub

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

Private Sub ComboBox1_Change()
    'Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    ComboBox2.Text = ""
    TextBox4.Text = ""
    
    If ComboBox1.Text = "" Then
    
        Frame2.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
        
        TextBox1.BackColor = &H80000004
        TextBox2.BackColor = &H80000004
        TextBox3.BackColor = &H80000004
        TextBox4.BackColor = &H80000004
        TextBox5.BackColor = &H80000004
        TextBox6.BackColor = &H80000004
        TextBox7.BackColor = &H80000004
        
        ComboBox2.Enabled = False
        ComboBox2.BackColor = &H80000004
        
        Label2.ForeColor = &H80000006
        Label3.ForeColor = &H80000006
        Label4.ForeColor = &H80000006
        Label5.ForeColor = &H80000006
        Label6.ForeColor = &H80000006
        Label7.ForeColor = &H80000006
        Label8.ForeColor = &H80000006
        Label9.ForeColor = &H80000006
        
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        
    Else
        
        Dim unitName As String
        unitName = ComboBox1.Text
        
        Dim Lenght As Integer
        Lenght = oDoc.ComponentDefinition.Occurrences.count
        
        Dim ExistPart As Boolean
        Dim i As Integer
        i = 0
        
        For Each oOcc In oDoc.ComponentDefinition.Occurrences
            If unitName = oOcc.Name Then
                ExistPart = True
                location = i + 1
                Exit For
            End If
            i = i + 1
        Next
        
        oDoc.SelectSet.Clear
        oDoc.SelectSet.Select (oOcc)
        
        Dim iProperty As PropertySets
        Set iProperty = oOcc.Definition.Document.PropertySets
        
        Dim userParameters As userParameters
        Set userParameters = oOcc.Definition.Parameters.userParameters
        
        TextBox1.Text = userParameters.Item("width").Expression
        TextBox2.Text = userParameters.Item("depth").Expression
        TextBox3.Text = userParameters.Item("height").Expression
        
        TextBox5.Text = iProperty.Item(1).Item(2).Expression
        TextBox6.Text = iProperty.Item(2).Item(2).Expression
        TextBox7.Text = iProperty.Item(3).Item(2).Expression + " ( " + iProperty.Item(3).Item(2).Value + " )"
        
        count = 0
        Dim parametersOcc As Parameter
        
        For Each parametersOcc In oOcc.Definition.Parameters.userParameters
            unitParametersArray(count) = parametersOcc.Name
            count = count + 1
        Next
        
        ComboBox2.List = unitParametersArray
    
        Frame2.Enabled = True
        Frame3.Enabled = True
        Frame4.Enabled = True
        TextBox1.BackColor = &H80000005
        TextBox2.BackColor = &H80000005
        TextBox3.BackColor = &H80000005
        TextBox5.BackColor = &H80000005
        TextBox6.BackColor = &H80000005
        TextBox7.BackColor = &H80000005
        ComboBox2.Enabled = True
        ComboBox2.BackColor = &H80000005
        Label2.ForeColor = &H80000012
        Label3.ForeColor = &H80000012
        Label4.ForeColor = &H80000012
        Label5.ForeColor = &H80000012
        Label6.ForeColor = &H80000012
        Label7.ForeColor = &H80000012
        Label8.ForeColor = &H80000012
        Label9.ForeColor = &H80000012
    End If
    
End Sub

Private Sub ComboBox2_Change()
    If ComboBox2.Text = "" Then
        TextBox4.Enabled = False
        TextBox4.BackColor = &H80000004
    Else
        
        Dim unitName, parameterName As String
        unitName = ComboBox1.Text
        parameterName = ComboBox2.Text
        
        TextBox4.Enabled = True
        TextBox4.BackColor = &H80000005
        
        Dim oOcc As ComponentOccurrence
        Set oOcc = oDoc.ComponentDefinition.Occurrences.Item(location)
        
        Dim userParameters As userParameters
        Set userParameters = oOcc.Definition.Parameters.userParameters
        
        TextBox4.Text = userParameters.Item(parameterName).Expression
        
        
    End If
End Sub

Private Sub CommandButton1_Click()
    setParameters
End Sub

Private Sub CommandButton2_Click()
    setParameters
    Unload Me
End Sub

Private Sub CommandButton3_Click()

        Dim oOcc As ComponentOccurrence
        Dim unitName As String
        unitName = ComboBox1.Text
        
        Dim ExistPart As Boolean
        Dim i As Integer
        i = 0
        
        For Each oOcc In oDoc.ComponentDefinition.Occurrences
            If unitName = oOcc.Name Then
                ExistPart = True
                location = i + 1
                Exit For
            End If
            i = i + 1
        Next
        
        If ExistPart = True Then
        
            Set oOcc = oDoc.ComponentDefinition.Occurrences.Item(location)
            Dim iProperty As PropertySets
            Dim Sub_oOcc As ComponentOccurrence
            Set iProperty = oOcc.Definition.Document.PropertySets
            iProperty.Item(3).Item(2).Expression = "=<Subject><Manager>"
            
            For Each Sub_oOcc In oOcc.Definition.Occurrences
                Set iProperty = Sub_oOcc.Definition.Document.PropertySets
                iProperty.Item(3).Item(2).Expression = "=<Subject><Manager>.<title>"
            Next
            
            Set iProperty = oOcc.Definition.Document.PropertySets
            TextBox5.Text = iProperty.Item(1).Item(2).Expression
            TextBox6.Text = iProperty.Item(2).Item(2).Expression
            TextBox7.Text = iProperty.Item(3).Item(2).Expression + " ( " + iProperty.Item(3).Item(2).Value + " )"
            
        End If
                
End Sub

Private Sub CommandButton4_Click()
    
    Dim assemblyname As String
    assemblyname = "E60-27-02"
    
    assemblyname = Left(assemblyname, InStr(InStr(1, assemblyname, "-") + 1, assemblyname, "-") - 1)
    
    
    MsgBox (assemblyname)
    
    partName = Replace(partName, assemblyname + "-", "", 2)
    
End Sub

Private Sub UserForm_Activate()
    
    'Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    Dim oOcc As ComponentOccurrence
    
    Dim count As Integer
    count = 0
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
            unitNameArray(count) = oOcc.Name
            count = count + 1
        End If
    Next
    
    Dim param As Parameter
   
    ComboBox1.List = unitNameArray
    
    If oDoc.SelectSet.count = 1 Then
        ComboBox1.Text = oDoc.SelectSet.Item(1).Name
        ComboBox1_Change
    End If
    
    Dim oOcc2 As ComponentOccurrence
    Dim Sub_oOcc As ComponentOccurrence
    Dim Temp As String
    
    For Each oOcc2 In oDoc.ComponentDefinition.Occurrences
        
        For Each Sub_oOcc In oOcc2.Definition.Occurrences
            Temp = Sub_oOcc.Name
            Temp = Replace(Temp, "PartName" + "-", "")
            
            If Left(Sub_oOcc.Name, 4) = "Door" Then
                CheckBox1.Value = Sub_oOcc.Visible
                Exit For
                Exit For
            ElseIf Left(Temp, 1) = "6" Then
                CheckBox1.Value = Sub_oOcc.Visible
                Exit For
                Exit For
            End If
        Next
    Next
    
    For Each oOcc2 In oDoc.ComponentDefinition.Occurrences
        
        For Each Sub_oOcc In oOcc2.Definition.Occurrences
            Temp = Sub_oOcc.Name
            Temp = Replace(Temp, "PartName" + "-", "")
            
            If Left(Sub_oOcc.Name, 3) = "Aft" Then
                CheckBox2.Value = Sub_oOcc.Visible
                Exit For
                Exit For
            End If
        Next
    Next
    
End Sub
    
