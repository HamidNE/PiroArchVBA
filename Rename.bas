Attribute VB_Name = "Rename"
Sub Rename()
    
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    
    If oDoc.SelectSet.Count = 1 Then
        
        Dim path As String
        Dim pathDir As String
        Dim unitName As String
        Dim pathFileL As String
        Dim pathFileN As String
        Dim oOcc As ComponentOccurrence
    
        NewAssemblyName = InputBox(NewAssemblyName, "Rename Item")
        
        Set oOcc = oDoc.SelectSet.Item(1)
        oOcc.Definition.Document.Save
        
        AssemblyName = oOcc.Name
        AssemblyName = Left(AssemblyName, InStr(1, AssemblyName, ":") - 1)
        
        path = oOcc.Definition.Document.File.FullFileName
        pathDir = Left(path, InStrRev(path, "\"))
        pathFileL = Mid(path, InStrRev(path, "\") + 1)
        
        oOcc.Name = Replace(oOcc.Name, AssemblyName, NewAssemblyName)
        pathFileN = Replace(pathFileL, AssemblyName, NewAssemblyName)
        
        If Dir(pathDir & pathFileL) <> "" Then
            Name pathDir & pathFileL As pathDir & pathFileN
        End If
        
        For Each oFile In oDoc.File.ReferencedFileDescriptors
            If pathFileL = Mid(oFile.ResolvedFullFileName, InStrRev(oFile.ResolvedFullFileName, "\") + 1) Then
                oFile.ReplaceReference (pathDir & pathFileN)
            End If
        Next
    
    Else
        Var = MsgBox("Plase select one item.", vbInformation, "Warrning")
    End If
    
    ThisApplication.ActiveDocument.Update

End Sub
