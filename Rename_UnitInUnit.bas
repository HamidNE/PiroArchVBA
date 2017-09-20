Attribute VB_Name = "Rename_UnitInUnit"
Sub Rename_UnitInUnit()
    
    Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Dim childOcc As ComponentOccurrence
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim AssemblyName As String
    Dim UnitName As String
    Dim PartName As String
    
    Dim path As String
    Dim pathDir As String
    Dim pathFileL As String
    Dim pathFileN As String
    
    AssemblyName = Left(oDoc.DisplayName, 2)
    
    For Each oOcc In oDoc.SelectSet
        
        path = oOcc.Definition.Document.File.FullFileName
        pathDir = Left(path, InStrRev(path, "\"))
        pathFileL = Mid(path, InStrRev(path, "\") + 1)
        UnitName = Left(oOcc.Name, InStr(oOcc.Name, ":") - 1)
        
        If Len(UnitName) - InStrRev(UnitName, AssemblyName) <> Len(AssemblyName) - 1 Then
        
            pathFileN = UnitName & "-" & AssemblyName & Mid(pathFileL, InStrRev(pathFileL, "."))
            oOcc.Name = UnitName & "-" & AssemblyName & Mid(oOcc.Name, InStrRev(oOcc.Name, ":"))
            
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
            UnitName = Left(childOcc.Name, InStr(4, childOcc.Name, ":") - 1)
            
            If Len(UnitName) - InStrRev(UnitName, AssemblyName) <> Len(AssemblyName) - 1 Then
            
                pathFileN = UnitName & "-" & AssemblyName & Mid(pathFileL, InStrRev(pathFileL, "."))
            
                childOcc.Name = UnitName & "-" & AssemblyName & Mid(childOcc.Name, InStrRev(childOcc.Name, ":"))
                    
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

