Attribute VB_Name = "Rename_UnitInUnit"
Sub Rename_UnitInUnit()
    
    Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Dim childOcc As ComponentOccurrence
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim AssemblyName As String
    Dim unitName As String
    Dim PartName As String
    
    Dim path As String
    Dim pathDir As String
    Dim pathFileL As String
    Dim pathFileN As String
    
    AssemblyName = oDoc.DisplayName
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        path = oOcc.Definition.Document.File.FullFileName
        pathDir = Left(path, InStrRev(path, "\"))
        pathFileL = Mid(path, InStrRev(path, "\") + 1)
        unitName = Left(oOcc.Name, InStr(oOcc.Name, ":") - 1)
        
        If InStrRev(unitName, AssemblyName) = 0 Then
        
            pathFileN = Replace(unitName, "-", "-" & AssemblyName & "-", 1, 1)
            
            If Len(pathFileN) = InStrRev(pathFileN, "-01") + 2 Then
                pathFileN = Left(pathFileN, InStrRev(pathFileN, "-01") - 1)
            End If
            
            oOcc.Name = pathFileN & Mid(oOcc.Name, InStrRev(oOcc.Name, ":"))
            pathFileN = pathFileN & Mid(pathFileL, InStrRev(pathFileL, "."))
            
            If Dir(pathDir & pathFileL) <> "" Then
                Name pathDir & pathFileL As pathDir & pathFileN
            End If
            
            For Each oFile In oDoc.File.ReferencedFileDescriptors
                If pathFileL = Mid(oFile.ResolvedFullFileName, InStrRev(oFile.ResolvedFullFileName, "\") + 1) Then
                    oFile.ReplaceReference (pathDir & pathFileN)
                End If
            Next
        
        End If
        
        For Each childOcc In oOcc.Definition.Occurrences
            
            path = childOcc.Definition.Document.File.FullFileName
            pathDir = Left(path, InStrRev(path, "\"))
            pathFileL = Mid(path, InStrRev(path, "\") + 1)
            unitName = Left(childOcc.Name, InStr(4, childOcc.Name, ":") - 1)
            
            If InStrRev(unitName, AssemblyName) = 0 Then
            
                pathFileN = Left(unitName, InStr(1, unitName, "-")) & Replace(unitName, "-", "-" & AssemblyName & "-", InStr(1, unitName, "-") + 1, 1)
                If Len(pathFileN) = InStrRev(pathFileN, "-01") + 2 Then
                    pathFileN = Left(pathFileN, InStrRev(pathFileN, "-01") - 1)
                End If
                
                childOcc.Name = pathFileN & Mid(childOcc.Name, InStrRev(childOcc.Name, ":"))
                pathFileN = pathFileN & Mid(pathFileL, InStrRev(pathFileL, "."))
                    
                If Dir(pathDir & pathFileL) <> "" Then
                    Name pathDir & pathFileL As pathDir & pathFileN
                End If
                
                For Each oFile In oOcc.Definition.Document.File.ReferencedFileDescriptors
                    If pathFileL = Mid(oFile.ResolvedFullFileName, InStrRev(oFile.ResolvedFullFileName, "\") + 1) Then
                        oFile.ReplaceReference (pathDir & pathFileN)
                        Exit For
                    End If
                Next
            
            End If
            
        Next
        
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub

