Attribute VB_Name = "Rename_Unit_Module"
Sub RenameUnitName()
    On Error GoTo ErrHandler:
        Rename_Unit.Show vbModeless
        Exit Sub
ErrHandler:
        Rename_Unit.Show vbModeless
End Sub


