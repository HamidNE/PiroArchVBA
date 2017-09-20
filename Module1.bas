Attribute VB_Name = "Module1"
Private Sub Browse_Click()

    Dim pathSelected As String         'This needs to be edited to find files, not just folders...alternatively, add .ipj to end of selected path?
    Dim ShellApp As Object
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Choose Folder", 0, OpenAt)
    pathSelected = ShellApp.self.path
    'Me.TextBox1.Text = pathSelected
    Set ShellApp = Nothing

End Sub
