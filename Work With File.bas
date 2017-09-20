Sub ListAllFiles()

    Dim fs As FileSearch, ws As Worksheet, i As Long
    Set fs = Application.FileSearch
    With fs
        .SearchSubFolders = ActiveSheet.Cells(1, 11).Value ' set to true if you want sub-folders included
        .FileType = msoFileTypeAllFiles 'can modify to just Excel files eg with msoFileTypeExcelWorkbooks
        .LookIn = ActiveSheet.Cells(1, 10).Value 'modify this to where you want to serach
        If .Execute > 0 Then
        Set ws = ActiveSheet
        For i = 1 To .FoundFiles.Count
        ws.Cells(i, 1) = .FoundFiles(i)
        Next
        Else
        MsgBox "No files found"
        End If
    End With

End Sub

- Rename Files -

Sub RenameFiles()

    For R = 1 To Range("A1").End(xlDown).Row
        OldFileName = Cells(R, 1).Value
        NewFileName = Cells(R, 2).Value
        On Error Resume Next
        If Not Dir(OldFileName) = "" Then Name OldFileName As NewFileName
        On Error GoTo 0
    Next

End Sub

Sub DoSomething()

    MyPath = "C:\"
    MyFile = "Book2.xls"
    NewName = "Othername.xls"

    If Dir(MyPath & MyFile) <> "" Then
        Name MyPath & MyFile As MyPath & NewName
    Else
        MsgBox "File not found"
    End If

End Sub