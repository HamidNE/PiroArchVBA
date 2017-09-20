VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Rename_Unit 
   Caption         =   "UserForm2"
   ClientHeight    =   2448
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4824
   OleObjectBlob   =   "Rename_Unit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Rename_Unit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oDoc As AssemblyDocument

Private Sub UserForm_Activate()
    
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim fileName As String
    fileName = oDoc.File.FullFileName
    fileName = Left(fileName, InStrRev(fileName, "\") - 1)
    fileName = Mid(fileName, InStrRev(fileName, "\") + 1)
    
    txtNewUnitName.Text = fileName
    
    Dim oUserParameters As UserParameters
    Set oUserParameters = oDoc.ComponentDefinition.Parameters.UserParameters
    For Each param In oUserParameters
        If param.Name = "Kitchen" Then
            Var = MsgBox("This app is not running in the kitchen environment.", vbExclamation, "Warrning")
            Unload Me
        End If
    Next
    
End Sub

Private Sub btnRenameSubUnit_Click()
       
    Dim oOcc As ComponentOccurrence
    Dim childOcc As ComponentOccurrence
    
    Dim AssemblyName As String
    Dim unitName As String
    Dim PartName As String
    
    Dim path As String
    Dim pathDir As String
    Dim pathFileL As String
    Dim pathFileN As String
    
    Dim SubUnitExist As Boolean
    Dim oUserParameters As UserParameters
    
    AssemblyName = oDoc.DisplayName
    AssemblyName = Replace(AssemblyName, ".iam", "")
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        path = oOcc.Definition.Document.File.FullFileName
        pathDir = Left(path, InStrRev(path, "\"))
        pathFileL = Mid(path, InStrRev(path, "\") + 1)
        unitName = Left(oOcc.Name, InStr(oOcc.Name, ":") - 1)
        
        SubUnitExist = False
        Set oUserParameters = oOcc.Definition.Parameters.UserParameters
        
        If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
            For Each parm In oUserParameters
                If parm.Name = "SubUnit" Then
                    SubUnitExist = True
                    parm.Value = True
                End If
            Next
        End If
        
        If SubUnitExist = False Then
            Dim oBooleanParam As UserParameter
            Set oBooleanParam = oUserParameters.AddByValue("SubUnit", True, kBooleanUnits)
        End If
        
        If InStrRev(unitName, AssemblyName) = 0 Then
        
            If InStr(1, unitName, "-") = 0 Then
                pathFileN = unitName + "-" + AssemblyName
            ElseIf InStr(1, unitName, "-" & AssemblyName) = 0 Then
                pathFileN = Replace(unitName, "-", "-" & AssemblyName & "-", 1, 1)
            Else
                pathFileN = unitName
            End If
            
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
                    Exit For
                End If
            Next
        
        End If
        
        For Each childOcc In oOcc.Definition.Occurrences
            
            path = childOcc.Definition.Document.File.FullFileName
            pathDir = Left(path, InStrRev(path, "\"))
            pathFileL = Mid(path, InStrRev(path, "\") + 1)
            unitName = Left(childOcc.Name, InStr(4, childOcc.Name, ":") - 1)
            
            If InStrRev(unitName, AssemblyName) = 0 Then
            
                If InStr(1, unitName, "-") = InStrRev(unitName, "-") Then
                    pathFileN = unitName + "-" + AssemblyName
                ElseIf InStr(1, unitName, "-" & AssemblyName) = 0 Then
                    pathFileN = Left(unitName, InStr(1, unitName, "-")) & Replace(unitName, "-", "-" & AssemblyName & "-", InStr(1, unitName, "-") + 1, 1)
                Else
                    pathFileN = unitName
                End If
            
                
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



Private Sub btnRenameUnit_Click()

    If CheckBox1.Value = False Then
        RenameNotSelected
    Else
        RenameSelected
    End If
    
End Sub
    
Sub RenameSelected()

    If oDoc.SelectSet.Count <> 1 Then
        Var = MsgBox("Plase Select One Item", vbExclamation, "Warrning")
        Exit Sub
    End If

    Dim oOcc As ComponentOccurrence
    Dim childOcc As ComponentOccurrence

    Dim AssemblyName As String
    Dim NewAssemblyName As String
    Dim unitName As String
    Dim PartName As String
    
    Dim path As String
    Dim pathDir As String
    Dim pathFileL As String
    Dim pathFileN As String
    
    NewAssemblyName = txtNewUnitName.Text
    
    For Each oOcc In oDoc.SelectSet
    
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
        
        For Each childOcc In oOcc.Definition.Occurrences
            
            path = childOcc.Definition.Document.File.FullFileName
            pathDir = Left(path, InStrRev(path, "\"))
            pathFileL = Mid(path, InStrRev(path, "\") + 1)
            
            childOcc.Name = Replace(childOcc.Name, AssemblyName, NewAssemblyName)
            pathFileN = Replace(pathFileL, AssemblyName, NewAssemblyName)
            
            If Dir(pathDir & pathFileL) <> "" Then
                Name pathDir & pathFileL As pathDir & pathFileN
            End If
            
            For Each oFile In oOcc.Definition.Document.File.ReferencedFileDescriptors
                If pathFileL = Mid(oFile.ResolvedFullFileName, InStrRev(oFile.ResolvedFullFileName, "\") + 1) Then
                    oFile.ReplaceReference (pathDir & pathFileN)
                    Exit For
                End If
            Next
            
        Next
        
    Next
    
    ThisApplication.ActiveDocument.Update

End Sub

Sub RenameNotSelected()


    Dim oOcc As ComponentOccurrence
    Dim childOcc As ComponentOccurrence
    
    Dim AssemblyName As String
    Dim NewAssemblyName As String
    Dim unitName As String
    Dim PartName As String
    
    Dim path As String
    Dim pathDir As String
    Dim pathFileL As String
    Dim pathFileN As String
    
    NewAssemblyName = txtNewUnitName.Text
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
    
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
        
        For Each childOcc In oOcc.Definition.Occurrences
            
            path = childOcc.Definition.Document.File.FullFileName
            pathDir = Left(path, InStrRev(path, "\"))
            pathFileL = Mid(path, InStrRev(path, "\") + 1)
            
            childOcc.Name = Replace(childOcc.Name, AssemblyName, NewAssemblyName)
            pathFileN = Replace(pathFileL, AssemblyName, NewAssemblyName)
            
            If Dir(pathDir & pathFileL) <> "" Then
                Name pathDir & pathFileL As pathDir & pathFileN
            End If
            
            For Each oFile In oOcc.Definition.Document.File.ReferencedFileDescriptors
                If pathFileL = Mid(oFile.ResolvedFullFileName, InStrRev(oFile.ResolvedFullFileName, "\") + 1) Then
                    oFile.ReplaceReference (pathDir & pathFileN)
                    Exit For
                End If
            Next
            
        Next
        
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub
