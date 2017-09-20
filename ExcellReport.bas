Attribute VB_Name = "ExcellReport"
Sub Report()

    Dim material_3
    Dim material_16
    Dim objDict_3 As Object
    Dim objDict_16 As Object
    Dim excellTemplatePath As String
    
    excellTemplatePath = "C:\Users\HamidNE\Documents\Custom Office Templates\andaze-363-().xltx"
    
    Dim path
    path = BrowseForFolder("")
    
    If path <> False Then
    
        path = path & "\"
    
        BOMExportExcel (path)
        Workbooks.Open path & "BOM.xlsx"
        Sort
        
        ''' Get Material with tinknees 3
        Worksheets("Sorted").Select
        Range("B1").AutoFilter Field:=2, Criteria1:="3"
        Range("A1").CurrentRegion.Copy
    
        Worksheets.Add
        ActiveSheet.Name = "3"
    
        Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
        
        Set objDict_3 = CreateObject("Scripting.Dictionary")
        material_3 = Application.Transpose(Range([a1], Cells(Rows.Count, "A").End(xlUp)))
        
        If IsArray(material_3) = True Then
            
            For i = 2 To UBound(material_3, 1)
                objDict_3(material_3(i)) = 1
            Next
                
            Range("DD1:DD" & objDict_3.Count) = Application.Transpose(objDict_3.Keys)
            material_3 = Application.Transpose(Range([dd1], Cells(Rows.Count, "DD").End(xlUp)))
            Range("DD1:DD" & objDict_3.Count).Delete
            
            Columns("A:A").Select
            Selection.EntireColumn.Hidden = True
            Columns("B:B").Select
            Selection.EntireColumn.Hidden = True
            
            Rows("1:1").Select
            Selection.EntireRow.Hidden = True
            
        End If
        
        ''' Get Material with tinknees 16
        Worksheets("Sorted").Select
        Range("B1").AutoFilter Field:=2, Criteria1:="16"
        Range("A1").CurrentRegion.Copy
    
        Worksheets.Add
        ActiveSheet.Name = "16"
        
        Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    
        Set objDict_16 = CreateObject("Scripting.Dictionary")
        material_16 = Application.Transpose(Range([a1], Cells(Rows.Count, "A").End(xlUp)))
        
        If IsArray(material_16) = True Then
        
            For i = 2 To UBound(material_16, 1)
                objDict_16(material_16(i)) = 1
            Next
            
            Range("DD1:DD" & objDict_16.Count) = Application.Transpose(objDict_16.Keys)
            material_16 = Application.Transpose(Range([dd1], Cells(Rows.Count, "DD").End(xlUp)))
            Range("DD1:DD" & objDict_16.Count).Delete
            
            Columns("A:A").Select
            Selection.EntireColumn.Hidden = True
            Columns("B:B").Select
            Selection.EntireColumn.Hidden = True
            
            Rows("1:1").Select
            Selection.EntireRow.Hidden = True
        
        End If
        
        ActiveWorkbook.Close True
        Workbooks.Open path & "BOM.xlsx"
        
        For i = 1 To objDict_3.Count
        
            Workbooks("BOM.xlsx").Activate
            Worksheets("3").Select
            
            If objDict_3.Count = 1 Then
                Range("A1").AutoFilter Field:=1, Criteria1:=material_3
            ElseIf objDict_3.Count > 1 Then
                Range("A1").AutoFilter Field:=1, Criteria1:=material_3(i)
            End If
            
            Range(Cells(2, 3), Cells(Cells(1, 1).End(xlDown).Row, 12)).Copy
            
            Workbooks.Add Template:=excellTemplatePath
            
            Range("C11").PasteSpecial xlPasteValues
            Range("G6").Value = "3"
            Range("G4:M4").Select
    
            If objDict_3.Count = 1 Then
                ActiveCell.Value = material_3
                ActiveWorkbook.SaveAs path & material_3 & "_3.xlsx"
            ElseIf objDict_3.Count > 1 Then
                ActiveCell.Value = material_3(i)
                ActiveWorkbook.SaveAs path & material_3(i) & "_3.xlsx"
            End If
    
            ActiveWorkbook.Close
            
        Next i
        
        For i = 1 To objDict_16.Count
        
            Workbooks("BOM.xlsx").Activate
            Worksheets("16").Select
            
            If objDict_16.Count = 1 Then
                Range("A1").AutoFilter Field:=1, Criteria1:=material_16
            ElseIf objDict_16.Count > 1 Then
                Range("A1").AutoFilter Field:=1, Criteria1:=material_16(i)
            End If
            
            Range(Cells(2, 3), Cells(Cells(1, 1).End(xlDown).Row, 12)).Copy
            
            Workbooks.Add Template:=excellTemplatePath
            
            Range("C11").PasteSpecial xlPasteValues
            Range("G6").Value = "16"
            Range("G4:M4").Select
    
            If objDict_16.Count = 1 Then
                ActiveCell.Value = material_16
                ActiveWorkbook.SaveAs path & material_16 & "_16.xlsx"
            ElseIf objDict_16.Count > 1 Then
                ActiveCell.Value = material_16(i)
                ActiveWorkbook.SaveAs path & material_16(i) & "_16.xlsx"
            End If
            
            ActiveWorkbook.Close
            
        Next i
        
        Workbooks("BOM.xlsx").Activate
        ActiveWorkbook.Close False
        
        Var = MsgBox("Reports of the operation was successful.", vbInformation, "Report")
    
    End If
    
End Sub

Public Sub BOMExportExcel(ByVal path As String)

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim obom As BOM
    Set obom = oDoc.ComponentDefinition.BOM
    
    obom.PartsOnlyViewEnabled = True
    
    Dim oPartsOnlyBOMView As BOMView
    Set oPartsOnlyBOMView = obom.BOMViews.Item("Parts Only")
    
    oPartsOnlyBOMView.Export path & "BOM.xlsx", kMicrosoftExcelFormat
    
End Sub

Public Function GetEndRows(ByVal Title As String)
    
    For i = 1 To 50
    
        Cells(1, i).Select
        
        If ActiveCell.Value = Title Then
            GetEndRows = Cells(1, i).End(xlDown).Row
            Exit For
        End If
        
    Next i
    
End Function

Public Sub SelectTitle(ByVal Title As String)
    
    Dim endRange As Integer
    endRange = GetEndRows("Part Number")
    
    For i = 1 To 50
    
        Cells(1, i).Select
        
        If ActiveCell.Value = Title Then
            Range(Cells(1, i), Cells(endRange, i)).Select
            Exit For
        End If
        
    Next i
    
End Sub

Sub ValidationColumn(ByVal column As String)

    Dim element As Range
    Dim MaxRows As Long
    
    Range(column & "3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:=" mm", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    'Range(column & "3").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.Replace What:=".", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    With Worksheets("Sorted")
        MaxRows = .Cells(.Rows.Count, column).End(xlUp).Row
    End With
    
    For Each element In Worksheets("Sorted").Range(column & "2:" & column & MaxRows)
        If IsNumeric(element.Value) Then
            element.Value = element.Value / 10
        End If
    Next

End Sub

Sub Sort()

    Dim titles
    titles = Array("Material", "t", "WH", "D", "Item QTY", "D-pvc", "WH-pvc", "Part Number", "D1", "D2", "WH1", "WH2")
    
    ActiveSheet.Name = "BOM"
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Sorted"
    
    For i = 1 To 12
        
        Sheets("BOM").Select
        SelectTitle (titles(i - 1))
        Selection.Copy
        
        Sheets("Sorted").Select
        Cells(1, i).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    Next i
    
    ValidationColumn ("C")
    ValidationColumn ("D")
    
    Range("B3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:=".000 mm", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.Replace What:="/", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


End Sub

Function BrowseForFolder(Optional OpenAt As Variant) As Variant

    Dim ShellApp As Object
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Please choose a folder", 0, OpenAt)
    
    On Error Resume Next
    BrowseForFolder = ShellApp.self.path
    On Error GoTo 0
    
    Set ShellApp = Nothing
    
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select

    Exit Function

Invalid:
    BrowseForFolder = False
End Function

