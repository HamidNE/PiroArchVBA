VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResizeableForm 
   Caption         =   "UserForm1"
   ClientHeight    =   10608
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9036.001
   OleObjectBlob   =   "ResizeableForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ResizeableForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub UserForm_Activate()
    
    Dim path As String
    path = InputBox(path)
    
    If Dir(path) <> "" Then
        
        Image1.Width = Round(LoadPicture(path).Width / 35.06)
        Image1.Height = Round(LoadPicture(path).Height / 35.06)
        UserForm1.Width = Round(LoadPicture(path).Width / 35.06) + 20
        UserForm1.Height = Round(LoadPicture(path).Height / 35.06) + 40
        Image1.Picture = LoadPicture(path)
    End If
    
End Sub
