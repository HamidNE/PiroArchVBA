VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Creat_Component_Form 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "Creat_Component_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Creat_Component_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function GetLastNamePart(ByVal partID As Integer)
    
    Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim Last As Integer
    Last = 0
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        If Left(oOcc.Name, 1) = partID Then
            
            If Last < CInt(Left(oOcc.Name, 2)) Then
                Last = CInt(Left(oOcc.Name, 2))
            End If
            
        End If
    Next
    
    Dim path As String
    path = Left(oDoc.FullFileName, InStrRev(oDoc.FullFileName, "\"))
    
    If Last = 0 Then
        Last = partID * 10
    End If
    
    GetLastNamePart = path & (Last + 1) & "-" & oDoc.DisplayName
    
    ThisApplication.CommandManager.PostPrivateEvent kFileNameEvent, GetLastNamePart
    
    Dim oDef As ControlDefinition
    Set oDef = ThisApplication.CommandManager.ControlDefinitions.Item("AssemblyCreateComponentCmd")
    
    Unload Me
    oDef.Execute

End Function

Private Sub AftButton_Click()
    GetLastNamePart (4)
End Sub

Private Sub CommandButton1_Click()
    GetLastNamePart (1)
End Sub

Private Sub CommandButton2_Click()
    GetLastNamePart (5)
End Sub

Private Sub DoorButton_Click()
    GetLastNamePart (6)
End Sub

Private Sub SideButton_Click()
    GetLastNamePart (2)
End Sub

Private Sub TopButton_Click()
    GetLastNamePart (3)
End Sub
