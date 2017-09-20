Attribute VB_Name = "Door_Angle"
Sub Door_Angle()

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oOcc As ComponentOccurrence
    
    ' Get the active document.
    Dim invDoc As Document
    Set invDoc = ThisApplication.ActiveDocument
    
    Dim Angle As String
    Angle = InputBox("Angle")
    
    Do While Angle = ""
        Angle = InputBox("Angle")
    Loop
    
    Dim oParam As Parameter
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        For Each oParam In oOcc.Definition.Parameters
            If oParam.Name = "do" Then
                oParam.Expression = Angle + " deg"
            End If
        Next
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub
