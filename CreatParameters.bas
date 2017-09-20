Sub CreatParameters()
    
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim parameterStr As String
    parameterStr = "d_21=depth"

    Dim parameterName As String
    Dim parameterValue As String

    parameterName = Left(parameterStr, InStr(1, parameterStr, "=") - 1)
    parameterValue = Right(parameterStr, InStr(1, parameterStr, "="))
    
    Dim userParams As UserParameters
    Set userParams = oDoc.ComponentDefinition.Parameters.UserParameters
    
    Dim param As Parameter
    Set param = userParams.AddByExpression(parameterName, parameterValue, kCentimeterLengthUnits)

End Sub