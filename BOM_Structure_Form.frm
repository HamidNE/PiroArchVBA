VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BOM_Structure_Form 
   Caption         =   "BOM Structure"
   ClientHeight    =   2172
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   OleObjectBlob   =   "BOM_Structure_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BOM_Structure_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnNormal_Click()

    Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Set oDoc = ThisApplication.ActiveDocument
    
    For Each oOcc In oDoc.SelectSet
    
        If oOcc.BOMStructure = kReferenceBOMStructure Then
            oOcc.Visible = True
            oOcc.BOMStructure = kDefaultBOMStructure
        End If
        
    Next
    
End Sub

Private Sub btnReference_Click()
    
    Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Set oDoc = ThisApplication.ActiveDocument
    
    For Each oOcc In oDoc.SelectSet
    
        If oOcc.BOMStructure = kNormalBOMStructure Then
            oOcc.Visible = False
            oOcc.BOMStructure = kReferenceBOMStructure
        End If
        
    Next
    
End Sub
