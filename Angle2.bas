Attribute VB_Name = "Angle2"
Sub ba()
    Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Set oDoc = ThisApplication.ActiveDocument

    ' For Each oOcc In oDoc.ComponentDefinition.Occurrences
    '     Dim surface As SurfaceBody
    '     Set surface = oOcc.Definition.SurfaceBodies.Item(1)
    '     Dim ofaces As faces
    '     Set ofaces = surface.faces
        
    '     If False Then
    '     End If
    ' Next
    
    Dim count As Integer
    count = oDoc.ComponentDefinition.Occurrences.count
    
    Dim faceA As Face
    Dim faceB As Face
    Dim oOccA As ComponentOccurrence
    Dim oOccB As ComponentOccurrence
    
    
    For i = 2 To count - 1

        Set oOccA = oDoc.ComponentDefinition.Occurrences.Item(i)
        Set faceA = oOccA.Definition.SurfaceBodies.Item(1).Faces.Item(1)

        For Each Face In oOccA.Definition.SurfaceBodies.Item(1).Faces
            If Face.Evaluator.Area > faceA.Evaluator.Area Then
                Set faceA = Face
            End If
        Next

        For j = i + 1 To count

            Set oOccB = oDoc.ComponentDefinition.Occurrences.Item(j)
            Set faceB = oOccB.Definition.SurfaceBodies.Item(1).Faces.Item(1)

            For Each Face In oOccB.Definition.SurfaceBodies.Item(1).Faces
                If Face.Evaluator.Area > faceB.Evaluator.Area Then
                    Set faceB = Face
                End If
            Next

            'Dist = ThisApplication.MeasureTools.GetAngle(faceA, faceB)
            MsgBox (ThisApplication.MeasureTools.GetAngle(faceA, faceB))
            
        Next j
    Next i
    
End Sub
