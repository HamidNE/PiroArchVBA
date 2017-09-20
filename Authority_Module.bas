Attribute VB_Name = "Authority_Module"
Sub Authority()

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oOcc As ComponentOccurrence
    Dim oUserPram As userParameters
    Dim oPram As UserParameter
    Dim iProperty As PropertySets
    
    Dim Authority As String
    Authority = InputBox("Authority")
    
    Do While Authority = ""
        MsgBox ("Plase Enter Valid Number")
        Authority = InputBox("Authority")
    Loop
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        Dim Width, Height, Depth As Double
        Dim existParameters(2) As Boolean
        Set oUserPram = oOcc.Definition.Parameters.userParameters
        
        For Each oPram In oUserPram
        
            If oPram.Name = "width" Then
                Width = oPram.Value
                existParameters(0) = True
            ElseIf oPram.Name = "height" Then
                Height = oPram.Value
                existParameters(1) = True
            ElseIf oPram.Name = "depth" Then
                Depth = oPram.Value
                existParameters(2) = True
            End If
            
        Next
        
        Set iProperty = oOcc.Definition.Document.PropertySets
        iProperty.Item("Design Tracking Properties").Item("Authority").Value = Authority
        
        If existParameters(0) = True And existParameters(0) = True And existParameters(0) = True Then
        
            Dim v, a, b, c As Double
        
            v = Width / 100 * Height / 100 * Depth / 100
            a = 1.2172 * (Depth / 100) ^ (-0.36)
            b = CDbl(iProperty.Item("Design Tracking Properties").Item("Authority").Value)
            c = iProperty.Item("Summary Information").Item("Revision Number").Value
            iProperty.Item("Design Tracking Properties").Item("Engineer").Value = a * b * Height / 100 * Depth / 1000
            iProperty.Item("Design Tracking Properties").Item("Cost center").Value = a * b * v * c

        End If
        
    Next
    
End Sub
