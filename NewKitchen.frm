VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewKitchen 
   Caption         =   "Kitchen"
   ClientHeight    =   6372
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8820.001
   OleObjectBlob   =   "NewKitchen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewKitchen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Path As String
Dim wallList
Dim wallListSize As Integer
Dim comboWallItem As Integer

Sub CreateAssemblyJointWithOffsetSample()

    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.Documents.Add(kAssemblyDocumentObject, "C:\Users\Public\Documents\Autodesk\Inventor 2017\Templates\Kitchen.iam")
    
    oDoc.FullFileName = "D:\Work\bbb.iam"
    oDoc.Save

End Sub

Private Sub CommandButton1_Click()

    Dim oFileDlg As FileDialog
    Call ThisApplication.CreateFileDialog(oFileDlg)

    oFileDlg.Filter = "Assembly Inventor Files (*.iam)|*.iam|All Files (*.*)|*.*"
    oFileDlg.FilterIndex = 1
    oFileDlg.DialogTitle = "Open File Test"
    oFileDlg.CancelError = True
    oFileDlg.OptionsEnabled = True

    On Error Resume Next
    oFileDlg.ShowSave
    
    If oFileDlg.FileName <> "" Then
        txtPath.Text = oFileDlg.FileName
    End If
    
End Sub

Private Sub CommandButton2_Click()

    comboWallItem = (comboWallItem - 1) Mod wallListSize
    ComboBoxWall.Text = wallList(comboWallItem)
    
End Sub

Private Sub CommandButton3_Click()

    comboWallItem = (comboWallItem + 1) Mod wallListSize
    ComboBoxWall.Text = wallList(comboWallItem)
    
    MsgBox (comboWallItem)
    
    Image1.Picture = LoadPicture("C:\Users\HamidNE\Desktop\Wall" & comboWallItem & ".jpg")
    
End Sub

Private Sub CommandButton5_Click()
    Unload Me
End Sub

Private Sub txtPath_Change()

    If txtPath.Text <> "" Then
        Frame1.Enabled = True
    Else
        Frame1.Enabled = False
    End If
    
End Sub

Private Sub UserForm_Activate()

    wallList = Array("L Wall", "U Wall", "G Wall")
    comboWallItem = 0
    wallListSize = 3
    
    ComboBoxWall.List = wallList
    ComboBoxWall.Text = wallList(comboWallItem)
    
End Sub
