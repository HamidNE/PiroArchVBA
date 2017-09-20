Attribute VB_Name = "Excell_1"
Sub Fileter()

    Workbooks.Open "C:\Users\HamidNE\Desktop\New.xlsm"

    Dim material_3
    Dim material_16
    Dim objDict_3 As Object
    Dim objDict_16 As Object
    
    Dim path
    path = BrowseForFolder("")
    path = path & "\"
    MsgBox (path)
    
    
    ''' Get Material with tinknees 3
    Worksheets("BOM").Select
    Range("S1").AutoFilter Field:=19, Criteria1:="3"
    Range("A1").CurrentRegion.Copy

    Worksheets.Add
    ActiveSheet.Name = "3"

    Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    
    Set objDict_3 = CreateObject("Scripting.Dictionary")
    material_3 = Application.Transpose(Range([b1], Cells(Rows.count, "B").End(xlUp)))
    
    For i = 2 To UBound(material_3, 1)
        objDict_3(material_3(i)) = 1
    Next
    
    Range("DD1:DD" & objDict_3.count) = Application.Transpose(objDict_3.Keys)
    material_3 = Application.Transpose(Range([dd1], Cells(Rows.count, "DD").End(xlUp)))
    Range("DD1:DD" & objDict_3.count).Delete
    
    Rows("1:1").Select
    Selection.EntireRow.Hidden = True
    
    
    
    ''' Get Material with tinknees 16
    Worksheets("BOM").Select
    Range("S1").AutoFilter Field:=19, Criteria1:="16"
    Range("A1").CurrentRegion.Copy

    Worksheets.Add
    ActiveSheet.Name = "16"
    
    Range("A1").PasteSpecial xlPasteValuesAndNumberFormats

    Set objDict_16 = CreateObject("Scripting.Dictionary")
    material_16 = Application.Transpose(Range([b1], Cells(Rows.count, "B").End(xlUp)))
    
    For i = 2 To UBound(material_16, 1)
        objDict_16(material_16(i)) = 1
    Next
    
    Range("DD1:DD" & objDict_16.count) = Application.Transpose(objDict_16.Keys)
    material_16 = Application.Transpose(Range([dd1], Cells(Rows.count, "DD").End(xlUp)))
    Range("DD1:DD" & objDict_16.count).Delete
    
    Rows("1:1").Select
    Selection.EntireRow.Hidden = True
    
    
    
    For i = 1 To objDict_3.count
    
        Workbooks("New.xlsm").Activate
        Worksheets("3").Select
        Range("B1").AutoFilter Field:=2, Criteria1:=material_3(i)
        
        Range("A1").CurrentRegion.Copy
        
        Workbooks.Add Template:= _
        "C:\Users\HamidNE\Documents\Custom Office Templates\Template.xltx"
        
        Range("B11").PasteSpecial xlPasteValues
        
        Range("G2").Value = "3"
        
        Range("G4:M4").Select
        ActiveCell.Value = material_3(i)
        
        ActiveWorkbook.SaveAs path & material_3(i) & "_3.xlsx"
        ActiveWorkbook.Close
        
    Next i
    
    For i = 1 To objDict_16.count
    
        Workbooks("New.xlsm").Activate
        Worksheets("16").Select
        Range("B1").AutoFilter Field:=2, Criteria1:=material_16(i)
        
        Range("A1").CurrentRegion.Copy
        
        Workbooks.Add Template:= _
        "C:\Users\HamidNE\Documents\Custom Office Templates\Template.xltx"
        
        Range("B11").PasteSpecial xlPasteValues
        
        Range("G4:M4").Select
        ActiveCell.Value = material_16(i)
        
        Range("G2").Value = "16"
        
        ActiveWorkbook.SaveAs path & material_16(i) & "_16.xlsx"
        ActiveWorkbook.Close
        
    Next i
    
    Workbooks("New.xlsm").Activate
    ActiveWorkbook.Close False
    
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

