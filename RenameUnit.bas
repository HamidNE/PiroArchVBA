Attribute VB_Name = "RenameUnit"
Sub RenameUnitOld()
    
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oOcc As ComponentOccurrence
    Dim LastName As String
    
    Dim path As String
    Dim pathDir As String
    Dim pathFileL As String
    Dim pathFileN As String
    
    Dim unitName As String
    Dim partName As String
    Dim subjectName As String
    
    Dim i As Integer
    i = 1
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        path = oOcc.Definition.Document.file.FullFileName
        pathDir = Left(path, InStrRev(path, "\"))
        pathFileL = Mid(path, InStrRev(path, "\") + 1)
        
        partName = oOcc.Name
        unitName = Left(partName, InStr(partName, "-") - 1)
        subjectName = Mid(partName, InStr(partName, "-") + 1, 2)
        
        
        partName = subjectName & "-" & unitName & Mid(partName, InStr(partName, "-") + 3)
        pathFileN = subjectName & "-" & unitName & Mid(pathFileL, InStr(pathFileL, "-") + 3)
        'MsgBox (partName)
        
        oOcc.Name = partName
        
        If Dir(pathDir & pathFileL) <> "" Then
            Name pathDir & pathFileL As pathDir & pathFileN
        End If
        
        For Each file In oDoc.file.ReferencedFileDescriptors
            If pathFileL = Mid(file.RelativeFileName, InStrRev(file.RelativeFileName, "\") + 1) Then
                file.ReplaceReference (pathDir & pathFileN)
            End If
        Next
        
        i = i + 1
        
    Next
    
End Sub
