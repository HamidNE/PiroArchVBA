Attribute VB_Name = "Fix_Moudle"
Sub Fixparameters()
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim UPrams As userParameters
    Dim UPram As UserParameter
    Set UPrams = oDoc.ComponentDefinition.Parameters.userParameters
    
    Dim DPL As Integer
    
    For Each UPram In UPrams
        If Left(UPram.Name, 2) = "d_" Or Left(UPram.Name, 3) = "wh_" Then
            DPL = InStr(1, UPram.Name, ":")
            If DPL > 0 Then
                UPram.Name = Left(UPram.Name, DPL - 1)
            End If
        End If
    Next
    
End Sub
