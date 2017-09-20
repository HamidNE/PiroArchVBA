VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Equipment_Form2 
   Caption         =   "Information"
   ClientHeight    =   7800
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9204.001
   OleObjectBlob   =   "Equipment_Form2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Equipment_Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oDoc As AssemblyDocument
Dim TotalCost1, TotalCost2 As Double

Sub Units()
    Dim oOcc As ComponentOccurrence
    Dim oProperty As Property
    
    Dim BaseCost, WallCost, TallCost As Double
    Dim unitName As String
    BaseCost = 0
    WallCost = 0
    TallCost = 0
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        Set oProperty = oOcc.Definition.Document.PropertySets.Item(3).Item(4)
        unitName = Left(oOcc.Name, 1)
        
        If unitName = "B" Then
            If oProperty.Expression <> "" Then
                BaseCost = BaseCost + CDbl(oProperty.Expression)
            End If
        ElseIf unitName = "D" Then
            If oProperty.Expression <> "" Then
                BaseCost = BaseCost + CDbl(oProperty.Expression)
            End If
        ElseIf unitName = "W" Then
            If oProperty.Expression <> "" Then
                WallCost = WallCost + CDbl(oProperty.Expression)
            End If
        ElseIf unitName = "F" Then
            If oProperty.Expression <> "" Then
                WallCost = WallCost + CDbl(oProperty.Expression)
            End If
        ElseIf unitName = "S" Then
            If oProperty.Expression <> "" Then
                WallCost = WallCost + CDbl(oProperty.Expression)
            End If
        ElseIf unitName = "T" Then
            If oProperty.Expression <> "" Then
                TallCost = TallCost + CDbl(oProperty.Expression)
            End If
        End If
    Next
    
    Label20.Caption = Round(BaseCost, 3)
    Label21.Caption = Round(WallCost, 3)
    Label23.Caption = Round(TallCost, 3)
    
    Label28.Caption = Round((BaseCost + WallCost + TallCost), 3)
    Label30.Caption = Round(((BaseCost + WallCost + TallCost) + Label29), 3)
    
End Sub

Sub Equiments()
    
    Dim oOcc As ComponentOccurrence
    Dim iProperty As PropertySets
    Dim oParams As Parameters
    Dim oParam As Parameter
    
    Dim s_hing, c_hing, Handle, stand, Abchek, P_JAck, H_Jack, Pakhore, Rail As Integer
    
    Dim Cost_Center As Double
    Cost_Center = 0
    
    Dim tempStr As String
    
    s_hing = 0
    c_hing = 0
    Handle = 0
    stand = 0
    Abchek = 0
    AbchekCount = 0
    P_JAck = 0
    H_Jack = 0
    Pakhore = 0
    Rail = 0
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
        Set oParams = oOcc.Definition.Parameters
        Set iProperty = oOcc.Definition.Document.PropertySets
        
        êtempStr = iProperty.Item("Design Tracking Properties").Item("Cost Center").Value
        
        If êtempStr = "" Then
            êtempStr = "0"
        End If
        
        'Cost_Center = Cost_Center + CDbl(iProperty.Item("Design Tracking Properties").Item("Cost Center").Value)
        
        For Each oParam In oParams
        
            If oParam.Name = "s_hing" Then
                s_hing = s_hing + oParam.Value
            ElseIf oParam.Name = "c_hing" Then
                c_hing = c_hing + oParam.Value
            ElseIf oParam.Name = "handle" Then
                Handle = Handle + oParam.Value
            ElseIf oParam.Name = "stand" Then
                stand = stand + oParam.Value
            ElseIf oParam.Name = "Abchek" Then
                Abchek = Abchek + oParam.Value
                AbchekCount = AbchekCount + 1
            ElseIf oParam.Name = "p_jack" Then
                P_JAck = P_JAck + oParam.Value
            ElseIf oParam.Name = "h_jack" Then
                H_Jack = H_Jack + oParam.Value
            ElseIf oParam.Name = "width" Then
                Pakhore = Pakhore + oParam.Value
            ElseIf oParam.Name = "t_rail" Then
                Rail = Rail + oParam.Value
            End If
            
        Next
    Next
    
    
    ' Write on Lables
    
    Count1.Caption = s_hing
    Count2.Caption = c_hing
    Count3.Caption = Handle
    Count4.Caption = stand
    Count5.Caption = Rail
    Size6.Caption = Pakhore / 100
    Count7.Caption = P_JAck
    Count8.Caption = H_Jack
    Count9.Caption = AbchekCount
    Size9.Caption = Abchek
    
    Dim temp1, temp2 As Double
    
    temp1 = CInt(avouch1.Text)
    temp2 = CInt(Count1.Caption)
    Cost1.Caption = Format((temp1 * temp2 / 1000), "0.000")
    
    temp1 = CInt(avouch2.Text)
    temp2 = CInt(Count2.Caption)
    Cost2.Caption = Format((temp1 * temp2 / 1000), "0.000")
    
    temp1 = CInt(avouch3.Text)
    temp2 = CInt(Count3.Caption)
    Cost3.Caption = Format((temp1 * temp2 / 1000), "0.000")
    
    temp1 = CInt(avouch4.Text)
    temp2 = CInt(Count4.Caption)
    Cost4.Caption = Format((temp1 * temp2 / 1000), "0.000")
    
    temp1 = CInt(avouch5.Text)
    temp2 = CInt(Count5.Caption)
    Cost5.Caption = Format((temp1 * temp2 / 1000), "0.000")
    
    temp1 = CDbl(avouch6.Text)
    temp2 = CDbl(Size6.Caption)
    Cost6.Caption = Format((temp1 * temp2 / 1000), "0.000")
    
    temp1 = CInt(avouch7.Text)
    temp2 = CInt(Count7.Caption)
    Cost7.Caption = Format((temp1 * temp2 / 1000), "0.000")
    
    temp1 = CInt(avouch8.Text)
    temp2 = CInt(Count8.Caption)
    Cost8.Caption = Format((temp1 * temp2 / 1000), "0.000")
    
    temp1 = CInt(avouch9.Text)
    temp2 = CInt(Count9.Caption)
    Cost9.Caption = Format((temp1 * temp2 / 1000), "0.000")
    
    TotalCost1 = 0
    
    TotalCost1 = TotalCost1 + CDbl(Cost1.Caption * 1000)
    TotalCost1 = TotalCost1 + CDbl(Cost2.Caption * 1000)
    TotalCost1 = TotalCost1 + CDbl(Cost3.Caption * 1000)
    TotalCost1 = TotalCost1 + CDbl(Cost4.Caption * 1000)
    TotalCost1 = TotalCost1 + CDbl(Cost5.Caption * 1000)
    TotalCost1 = TotalCost1 + CDbl(Cost6.Caption * 1000)
    TotalCost1 = TotalCost1 + CDbl(Cost7.Caption * 1000)
    TotalCost1 = TotalCost1 + CDbl(Cost8.Caption * 1000)
    TotalCost1 = TotalCost1 + CDbl(Cost9.Caption * 1000)
    
    
    TotalCost_Label.Caption = Format((TotalCost1 / 1000), "0.000")
    Label29.Caption = Format((TotalCost1 / 1000), "0.000")
    Residue_Label.Caption = 0
    
End Sub

Private Sub CommandButton1_Click()

    If Purchase1.Text = "" Then
        temp1 = CInt(avouch1.Text)
        temp2 = CInt(Count1.Caption)
        Cost1.Caption = Format((temp1 * temp2 / 1000), "0.000")
    Else
        temp1 = CInt(Purchase1.Text)
        temp2 = CInt(Count1.Caption)
        Cost1.Caption = Format((temp1 * temp2 / 1000), "0.000")
    End If
    
    If Purchase2.Text = "" Then
        temp1 = CInt(avouch2.Text)
        temp2 = CInt(Count2.Caption)
        Cost2.Caption = Format((temp1 * temp2 / 1000), "0.000")
    Else
        temp1 = CInt(Purchase2.Text)
        temp2 = CInt(Count2.Caption)
        Cost2.Caption = Format((temp1 * temp2 / 1000), "0.000")
    End If
    
    If Purchase3.Text = "" Then
        temp1 = CInt(avouch3.Text)
        temp2 = CInt(Count3.Caption)
        Cost3.Caption = Format((temp1 * temp2 / 1000), "0.000")
    Else
        temp1 = CInt(Purchase3.Text)
        temp2 = CInt(Count3.Caption)
        Cost3.Caption = Format((temp1 * temp2 / 1000), "0.000")
    End If
    
    If Purchase4.Text = "" Then
        temp1 = CInt(avouch4.Text)
        temp2 = CInt(Count4.Caption)
        Cost4.Caption = Format((temp1 * temp2 / 1000), "0.000")
    Else
        temp1 = CInt(Purchase4.Text)
        temp2 = CInt(Count4.Caption)
        Cost4.Caption = Format((temp1 * temp2 / 1000), "0.000")
    End If
    
    If Purchase5.Text = "" Then
        temp1 = CInt(avouch5.Text)
        temp2 = CInt(Count5.Caption)
        Cost5.Caption = Format((temp1 * temp2 / 1000), "0.000")
    Else
        temp1 = CInt(Purchase5.Text)
        temp2 = CInt(Count5.Caption)
        Cost5.Caption = Format((temp1 * temp2 / 1000), "0.000")
    End If
    
    If Purchase6.Text = "" Then
        temp1 = CDbl(avouch6.Text)
        temp2 = CDbl(Size6.Caption)
        Cost6.Caption = Format((temp1 * temp2 / 1000), "0.000")
    Else
        temp1 = CInt(Purchase6.Text)
        temp2 = CInt(Size6.Caption)
        Cost6.Caption = Format((temp1 * temp2 / 1000), "0.000")
    End If
    
    If Purchase7.Text = "" Then
        temp1 = CInt(avouch7.Text)
        temp2 = CInt(Count7.Caption)
        Cost7.Caption = Format((temp1 * temp2 / 1000), "0.000")
    Else
        temp1 = CInt(Purchase7.Text)
        temp2 = CInt(Count7.Caption)
        Cost7.Caption = Format((temp1 * temp2 / 1000), "0.000")
    End If
    
    If Purchase8.Text = "" Then
        temp1 = CInt(avouch8.Text)
        temp2 = CInt(Count8.Caption)
        Cost8.Caption = Format((temp1 * temp2 / 1000), "0.000")
    Else
        temp1 = CInt(Purchase8.Text)
        temp2 = CInt(Count8.Caption)
        Cost8.Caption = Format((temp1 * temp2 / 1000), "0.000")
    End If
    
    If Purchase9.Text = "" Then
        temp1 = CInt(avouch9.Text)
        temp2 = CInt(Count9.Caption)
        Cost9.Caption = Format((temp1 * temp2 / 1000), "0.000")
    Else
        temp1 = CInt(Purchase9.Text)
        temp2 = CInt(Count9.Caption)
        Cost9.Caption = Format((temp1 * temp2 / 1000), "0.000")
    End If
    
    TotalCost1 = 0
    
    TotalCost1 = TotalCost1 + CDbl(Count1.Caption * avouch1.Value)
    TotalCost1 = TotalCost1 + CDbl(Count2.Caption * avouch2.Value)
    TotalCost1 = TotalCost1 + CDbl(Count3.Caption * avouch3.Value)
    TotalCost1 = TotalCost1 + CDbl(Count4.Caption * avouch4.Value)
    TotalCost1 = TotalCost1 + CDbl(Count5.Caption * avouch5.Value)
    TotalCost1 = TotalCost1 + CDbl(Size6.Caption * avouch6.Value)
    TotalCost1 = TotalCost1 + CDbl(Count7.Caption * avouch7.Value)
    TotalCost1 = TotalCost1 + CDbl(Count8.Caption * avouch8.Value)
    TotalCost1 = TotalCost1 + CDbl(Count9.Caption * avouch9.Value)
    
    TotalCost2 = 0
    
    TotalCost2 = TotalCost2 + CDbl(Cost1.Caption * 1000)
    TotalCost2 = TotalCost2 + CDbl(Cost2.Caption * 1000)
    TotalCost2 = TotalCost2 + CDbl(Cost3.Caption * 1000)
    TotalCost2 = TotalCost2 + CDbl(Cost4.Caption * 1000)
    TotalCost2 = TotalCost2 + CDbl(Cost5.Caption * 1000)
    TotalCost2 = TotalCost2 + CDbl(Cost6.Caption * 1000)
    TotalCost2 = TotalCost2 + CDbl(Cost7.Caption * 1000)
    TotalCost2 = TotalCost2 + CDbl(Cost8.Caption * 1000)
    TotalCost2 = TotalCost2 + CDbl(Cost9.Caption * 1000)
    
    TotalCost_Label.Caption = Format(TotalCost2 / 1000, "0.000")
    
    If TotalCost2 - TotalCost1 = 0 Then
        Residue_Label.Caption = 0
    Else
        Residue_Label.Caption = Format((TotalCost2 - TotalCost1) / 1000, "0.000")
    End If
    
    Label29.Caption = Format((TotalCost2 / 1000), "0.000")
    
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub CommandButton4_Click()
    Unload Me
End Sub

Private Sub avouch1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub avouch2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub avouch3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub avouch4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub avouch5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub avouch6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub avouch7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub avouch8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub avouch9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub Purchase1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub
Private Sub Purchase2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub
Private Sub Purchase3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub
Private Sub Purchase4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub
Private Sub Purchase5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub
Private Sub Purchase6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub
Private Sub Purchase7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub
Private Sub Purchase8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub
Private Sub Purchase9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub UserForm_Activate()
    
    Set oDoc = ThisApplication.ActiveDocument
    Equiments
    Units
    
End Sub


