VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResizeableForm 
   Caption         =   "UserForm1"
   ClientHeight    =   8700.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6588
   OleObjectBlob   =   "ResizeableForm_v1.0.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ResizeableForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim unitName As String
Dim pathUnits As String
Dim numPicture As Integer

Sub UserForm_Activate()
    
    numPicture = 1
    Dim path As String
    Dim pictureLoaded As Boolean
    
    pathUnits = "D:\Work\Inventor\UNITS\Unit\E\"
    
    If ResizeableForm.Tag <> "" Then
        
        unitName = ResizeableForm.Tag
        path = pathUnits & unitName & "\" & unitName & "-1.jpg"
        If Dir(path) <> "" Then
            Set Image1.Picture = Nothing
            Image1.Picture = LoadPicture(path)
            pictureLoaded = True
        Else
            Unload Me
        End If
        ResizeableForm.Tag = ""
        
    Else
    
        path = InputBox(path)
        If path <> "" And Dir(path) <> "" Then
            Set Image1.Picture = Nothing
            Image1.Picture = LoadPicture(path)
            pictureLoaded = True
        End If
        
    End If
    
    If pictureLoaded = True Then
        
        Image1.Width = Round(LoadPicture(path).Width / 35.06)
        Image1.Height = Round(LoadPicture(path).Height / 35.06)
        ResizeableForm.Width = Round(LoadPicture(path).Width / 35.06) + 20
        ResizeableForm.Height = Round(LoadPicture(path).Height / 35.06) + 40
        
    End If
    
End Sub

Private Sub CommandButton1_Click()

    numPicture = (numPicture + 2) Mod 3
    
    If unitName <> "" Then
        path = pathUnits & unitName & "\" & unitName & "-" & numPicture & ".jpg"
        Set Image1.Picture = Nothing
        Image1.Picture = LoadPicture(path)
        
        Image1.Width = Round(LoadPicture(path).Width / 35.06)
        Image1.Height = Round(LoadPicture(path).Height / 35.06)
        ResizeableForm.Width = Round(LoadPicture(path).Width / 35.06) + 20
        ResizeableForm.Height = Round(LoadPicture(path).Height / 35.06) + 40
        
    End If
    
End Sub

Private Sub CommandButton2_Click()
    
    numPicture = (numPicture Mod 3) + 1
    
    If unitName <> "" Then
        path = pathUnits & unitName & "\" & unitName & "-" & numPicture & ".jpg"
        Set Image1.Picture = Nothing
        Image1.Picture = LoadPicture(path)
        
        Image1.Width = Round(LoadPicture(path).Width / 35.06)
        Image1.Height = Round(LoadPicture(path).Height / 35.06)
        ResizeableForm.Width = Round(LoadPicture(path).Width / 35.06) + 20
        ResizeableForm.Height = Round(LoadPicture(path).Height / 35.06) + 40
        
    End If
    
End Sub
