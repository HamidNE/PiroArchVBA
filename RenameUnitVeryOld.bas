Attribute VB_Name = "RenameUnitVeryOld"
Sub RenameUnitVeryOld()
    
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oOcc As ComponentOccurrence
    Dim LastName As String
    
    Dim path As String
    Dim pathDir As String
    Dim pathFileL As String
    Dim pathFileN As String
    
    Dim partName As String
    Dim subjectName As String
    
    Dim i As Integer
    i = 1
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        path = oOcc.Definition.Document.file.FullFileName
        pathDir = Left(path, InStrRev(path, "\"))
        pathFileL = Mid(path, InStrRev(path, "\") + 1)
        
        partName = oOcc.Name
        subjectName = Left(partName, InStr(partName, "-") - 1)
        
        If subjectName = "Bott" Then
            subjectName = "11"
        ElseIf subjectName = "Bott1" Then
            subjectName = "11"
        ElseIf subjectName = "Aft" Then
            subjectName = "41"
        ElseIf subjectName = "Aft1" Then
            subjectName = "41"
        ElseIf subjectName = "Aft2" Then
            subjectName = "42"
        ElseIf subjectName = "Aft3" Then
            subjectName = "43"
        ElseIf subjectName = "Side1" Then
            subjectName = "21"
        ElseIf subjectName = "Side2" Then
            subjectName = "22"
        ElseIf subjectName = "Side3" Then
            subjectName = "23"
        ElseIf subjectName = "Shelf" Then
            subjectName = "51"
        ElseIf subjectName = "Shelf1" Then
            subjectName = "51"
        ElseIf subjectName = "Shelf2" Then
            subjectName = "52"
        ElseIf subjectName = "Door" Then
            subjectName = "61"
        ElseIf subjectName = "Door1" Then
            subjectName = "61"
        ElseIf subjectName = "Door2" Then
            subjectName = "62"
        ElseIf subjectName = "Door3" Then
            subjectName = "63"
        ElseIf subjectName = "Top" Then
            subjectName = "31"
        ElseIf subjectName = "Top1" Then
            subjectName = "31"
        ElseIf subjectName = "Top2" Then
            subjectName = "32"
        ElseIf subjectName = "Top3" Then
            subjectName = "33"
        End If
        
        partName = subjectName & Mid(partName, InStr(partName, "-"))
        pathFileN = subjectName & Mid(pathFileL, InStr(pathFileL, "-"))
        
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
