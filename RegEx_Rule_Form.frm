VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegEx_Rule_Form 
   Caption         =   "UserForm1"
   ClientHeight    =   6420
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9012.001
   OleObjectBlob   =   "RegEx_Rule_Form_v1.3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegEx_Rule_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Function RxReplace( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    ByVal ReplacePattern As String, _
    Optional ByVal IgnoreCase As Boolean = False, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal MatchGlobal As Boolean = True) As String
 
    With New RegExp
        .MultiLine = MultiLine
        .IgnoreCase = IgnoreCase
        .Global = MatchGlobal
        .Pattern = Pattern
        RxReplace = .Replace(SourceString, ReplacePattern)
    End With
 
End Function

Private Sub CommandButton1_Click()
    
    Dim str As String
    Dim Pattern As String
    Dim ReplacePattern As String

    If OptionButton1.Value = True Then

        Pattern = "Parameter\(""(\w+)-""\+s\+"":\d"", ""WH""\)=(.+)\s*Parameter\(""(\w+)-""\+s\+"":\d"", ""D""\)=(.+)"
        ReplacePattern = "d_$1=$2" & Chr(10) & "wh_$3=$4"
        str = RxReplace(TextBox1.Text, Pattern, ReplacePattern)
        
        Pattern = "Parameter\(""(\w+)-""\+s\+"":\d"", ""D""\)=(.+)\s*Parameter\(""(\w+)-""\+s\+"":\d"", ""WH""\)=(.+)"
        ReplacePattern = "d_$1=$2" & Chr(10) & "wh_$3=$4"
        str = RxReplace(str, Pattern, ReplacePattern)

    ElseIf OptionButton2.Value = True Then

        Pattern = "k\d?=""(\d{2})-""\+s\+"":\d""\s*Parameter\(k\d?, ""D""\)=(.+)\s*Parameter\(k\d?, ""WH""\)=(.+)"
        ReplacePattern = "d_$1=$2" & Chr(10) & "wh_$1=$3"
        str = RxReplace(TextBox1.Text, Pattern, ReplacePattern)

        Pattern = "k\d?=""(\d{2})-""\+s\+"":\d""\s*Parameter\(k\d?, ""WH""\)=(.+)\s*Parameter\(k\d?, ""D""\)=(.+)"
        ReplacePattern = "wh_$1=$2" & Chr(10) & "d_$1=$3"
        str = RxReplace(str, Pattern, ReplacePattern)

    End If
    
    Pattern = "(\n\r)+"
    ReplacePattern = "$1"
    str = RxReplace(str, Pattern, ReplacePattern)
    
    Pattern = "^['abvsciPDkIEC\t].*"
    ReplacePattern = ""
    str = RxReplace(str, Pattern, ReplacePattern)
    
    Pattern = "(\n)+"
    ReplacePattern = "$1"
    str = RxReplace(str, Pattern, ReplacePattern)
    
    Dim WrdArray() As String
    WrdArray() = Split(str, Chr(10))
    
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument

    Dim parameterName As String
    Dim parameterValue As String
    
    Dim param As Parameter
    Dim userParams As UserParameters
    Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters
    
    For Each Value In WrdArray

        If Left(Value, 2) = "d_" Then
        
            parameterName = Left(Value, InStr(1, Value, "=") - 1)
            parameterValue = Mid(Value, InStr(1, Value, "=") + 1)

            If OptionButton1.Value = True Then

                subjectName = Mid(parameterName, 3)
            
                If subjectName = "Bott" Then
                    subjectName = "11"
                ElseIf subjectName = "Bott1" Then
                    subjectName = "11"
                ElseIf subjectName = "Aft" Then
                    subjectName = "41"
                ElseIf subjectName = "Aft1" Then
                    subjectName = "41"
                ElseIf subjectName = "Aft2" Then
                    subjectName = "42"
                ElseIf subjectName = "Aft3" Then
                    subjectName = "43"
                ElseIf subjectName = "Side1" Then
                    subjectName = "21"
                ElseIf subjectName = "Side2" Then
                    subjectName = "22"
                ElseIf subjectName = "Side3" Then
                    subjectName = "23"
                ElseIf subjectName = "Shelf" Then
                    subjectName = "51"
                ElseIf subjectName = "Shelf1" Then
                    subjectName = "51"
                ElseIf subjectName = "Shelf2" Then
                    subjectName = "52"
                ElseIf subjectName = "Shelf3" Then
                    subjectName = "53"
                ElseIf subjectName = "Shelf4" Then
                    subjectName = "54"
                ElseIf subjectName = "Shelf5" Then
                    subjectName = "55"
                ElseIf subjectName = "Door" Then
                    subjectName = "61"
                ElseIf subjectName = "Door1" Then
                    subjectName = "61"
                ElseIf subjectName = "Door2" Then
                    subjectName = "62"
                ElseIf subjectName = "Door3" Then
                    subjectName = "63"
                ElseIf subjectName = "Door4" Then
                    subjectName = "64"
                ElseIf subjectName = "Top" Then
                    subjectName = "31"
                ElseIf subjectName = "Top1" Then
                    subjectName = "31"
                ElseIf subjectName = "Top2" Then
                    subjectName = "32"
                ElseIf subjectName = "Top3" Then
                    subjectName = "33"
                ElseIf subjectName = "Top4" Then
                    subjectName = "34"
                End If
                
                parameterName = "d_" & subjectName

            End If
            
            Set param = userParams.AddByExpression(parameterName, parameterValue, kCentimeterLengthUnits)
        
        ElseIf Left(Value, 3) = "wh_" Then
        
            parameterName = Left(Value, InStr(1, Value, "=") - 1)
            parameterValue = Mid(Value, InStr(1, Value, "=") + 1)
            
            If OptionButton1.Value = True Then

                subjectName = Mid(parameterName, 4)
            
                If subjectName = "Bott" Then
                    subjectName = "11"
                ElseIf subjectName = "Bott1" Then
                    subjectName = "11"
                ElseIf subjectName = "Aft" Then
                    subjectName = "41"
                ElseIf subjectName = "Aft1" Then
                    subjectName = "41"
                ElseIf subjectName = "Aft2" Then
                    subjectName = "42"
                ElseIf subjectName = "Aft3" Then
                    subjectName = "43"
                ElseIf subjectName = "Side1" Then
                    subjectName = "21"
                ElseIf subjectName = "Side2" Then
                    subjectName = "22"
                ElseIf subjectName = "Side3" Then
                    subjectName = "23"
                ElseIf subjectName = "Shelf" Then
                    subjectName = "51"
                ElseIf subjectName = "Shelf1" Then
                    subjectName = "51"
                ElseIf subjectName = "Shelf2" Then
                    subjectName = "52"
                ElseIf subjectName = "Shelf3" Then
                    subjectName = "53"
                ElseIf subjectName = "Shelf4" Then
                    subjectName = "54"
                ElseIf subjectName = "Shelf5" Then
                    subjectName = "55"
                ElseIf subjectName = "Door" Then
                    subjectName = "61"
                ElseIf subjectName = "Door1" Then
                    subjectName = "61"
                ElseIf subjectName = "Door2" Then
                    subjectName = "62"
                ElseIf subjectName = "Door3" Then
                    subjectName = "63"
                ElseIf subjectName = "Door4" Then
                    subjectName = "64"
                ElseIf subjectName = "Top" Then
                    subjectName = "31"
                ElseIf subjectName = "Top1" Then
                    subjectName = "31"
                ElseIf subjectName = "Top2" Then
                    subjectName = "32"
                ElseIf subjectName = "Top3" Then
                    subjectName = "33"
                ElseIf subjectName = "Top4" Then
                    subjectName = "34"
                End If
                
                parameterName = "wh_" & subjectName

            End If
            
            Set param = userParams.AddByExpression(parameterName, parameterValue, kCentimeterLengthUnits)
            
        End If

    Next
    
    oDoc.UnitsOfMeasure.LengthUnits = kCentimeterLengthUnits
    
    For Each oPart In oDoc.AllReferencedDocuments
        oPart.UnitsOfMeasure.LengthUnits = kCentimeterLengthUnits
    Next
    
    If OptionButton1.Value = True Then
        RenameUnitVeryOld
    End If
    
    Set param = userParams.AddByExpression("Style", "1", kUnitlessUnits)
    Set param = userParams.AddByExpression("StyleCount", "1", kUnitlessUnits)
    Set param = userParams.AddByValue("Unit", True, kBooleanUnits)
    
    oDoc.UnitsOfMeasure.LengthUnits = kCentimeterLengthUnits
    
    For Each oPart In oDoc.AllReferencedDocuments
        oPart.UnitsOfMeasure.LengthUnits = kCentimeterLengthUnits
    Next
    
    Dim oDoc As AssemblyDocument
    Dim oOcc As ComponentOccurrence
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim ascTemp As Integer
    Dim Params As Parameters
    Set Params = oDoc.ComponentDefinition.Parameters
    
    Dim userParam As UserParameter
    For Each userParam In oDoc.ComponentDefinition.Parameters.UserParameters
        ascTemp = Asc(Left(userParam.Expression, 1))
        If userParam.Units = "mm" Then
            userParam.Units = "cm"
        End If
        If ascTemp < 58 And ascTemp > 47 Then
            userParam.Expression = userParam.Value
        End If
    Next
    
    Dim modelParam As ModelParameter
    For Each modelParam In oDoc.ComponentDefinition.Parameters.ModelParameters
        If modelParam.Units = "mm" Then
            modelParam.Units = "cm"
        End If
        ascTemp = Asc(Left(modelParam.Expression, 1))
        If ascTemp < 58 And ascTemp > 47 Then
            modelParam.Expression = modelParam.Value
        End If
    Next
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        Set Params = oOcc.Definition.Parameters
        
        For Each userParam In Params.UserParameters
            If userParam.Units = "mm" Then
                userParam.Units = "cm"
            End If
            ascTemp = Asc(Left(userParam.Expression, 1))
            If ascTemp < 58 And ascTemp > 47 Then
                userParam.Expression = userParam.Value
            End If
        Next
    
        For Each modelParam In Params.ModelParameters
            If modelParam.Units = "mm" Then
                modelParam.Units = "cm"
            End If
            ascTemp = Asc(Left(modelParam.Expression, 1))
            If ascTemp < 58 And ascTemp > 47 Then
                modelParam.Expression = modelParam.Value
            End If
        Next
        
    Next
    
    ThisApplication.ActiveDocument.Update
    
End Sub

Sub RenameUnitVeryOld()
    
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oOcc As ComponentOccurrence
    Dim LastName As String
    
    Dim path As String
    Dim pathDir As String
    Dim pathFileL As String
    Dim pathFileN As String
    
    Dim PartName As String
    Dim subjectName As String
    
    Dim i As Integer
    i = 1
    
    For Each oOcc In oDoc.ComponentDefinition.Occurrences
        
        path = oOcc.Definition.Document.File.FullFileName
        pathDir = Left(path, InStrRev(path, "\"))
        pathFileL = Mid(path, InStrRev(path, "\") + 1)
        
        PartName = oOcc.Name
        subjectName = Left(PartName, InStr(PartName, "-") - 1)
        
        If subjectName = "Bott" Then
            subjectName = "11"
        ElseIf subjectName = "Bott1" Then
            subjectName = "11"
        ElseIf subjectName = "Aft" Then
            subjectName = "41"
        ElseIf subjectName = "Aft1" Then
            subjectName = "41"
        ElseIf subjectName = "Aft2" Then
            subjectName = "42"
        ElseIf subjectName = "Aft3" Then
            subjectName = "43"
        ElseIf subjectName = "Side1" Then
            subjectName = "21"
        ElseIf subjectName = "Side2" Then
            subjectName = "22"
        ElseIf subjectName = "Side3" Then
            subjectName = "23"
        ElseIf subjectName = "Shelf" Then
            subjectName = "51"
        ElseIf subjectName = "Shelf1" Then
            subjectName = "51"
        ElseIf subjectName = "Shelf2" Then
            subjectName = "52"
        ElseIf subjectName = "Shelf3" Then
            subjectName = "53"
        ElseIf subjectName = "Shelf4" Then
            subjectName = "54"
        ElseIf subjectName = "Shelf5" Then
            subjectName = "55"
        ElseIf subjectName = "Door" Then
            subjectName = "61"
        ElseIf subjectName = "Door1" Then
            subjectName = "61"
        ElseIf subjectName = "Door2" Then
            subjectName = "62"
        ElseIf subjectName = "Door3" Then
            subjectName = "63"
        ElseIf subjectName = "Door4" Then
            subjectName = "64"
        ElseIf subjectName = "Top" Then
            subjectName = "31"
        ElseIf subjectName = "Top1" Then
            subjectName = "31"
        ElseIf subjectName = "Top2" Then
            subjectName = "32"
        ElseIf subjectName = "Top3" Then
            subjectName = "33"
        ElseIf subjectName = "Top4" Then
            subjectName = "34"
        End If
        
        PartName = subjectName & Mid(PartName, InStr(PartName, "-"))
        pathFileN = subjectName & Mid(pathFileL, InStr(pathFileL, "-"))
        
        oOcc.Name = PartName
        
        If Dir(pathDir & pathFileL) <> "" Then
            Name pathDir & pathFileL As pathDir & pathFileN
        End If
        
        For Each File In oDoc.File.ReferencedFileDescriptors
            If pathFileL = Mid(File.RelativeFileName, InStrRev(File.RelativeFileName, "\") + 1) Then
                File.ReplaceReference (pathDir & pathFileN)
            End If
        Next
        
        i = i + 1
        
    Next
    
End Sub



