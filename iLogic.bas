Attribute VB_Name = "iLogic"
Sub Internal_iLogic()

    Dim addIn As ApplicationAddIn
    Dim addIns As ApplicationAddIns
    
    Dim i As Integer
    i = 0
    Set addIns = ThisApplication.ApplicationAddIns
        For Each addIn In addIns
            If InStr(addIn.DisplayName, "iLogic") > 0 Then
                            addIn.Activate
                Dim iLogicAuto As Object
                Set iLogicAuto = addIn.Automation
                Exit For
            End If
            i = i + 1
        Next
    Debug.Print addIn.DisplayName
    
    MsgBox (i)
     
     
    Dim RuleName1 As String
    EXTERNALrule = "Rule - Set Formula2"
    
    Dim RuleName2 As String
    INTERNALrule = "Rule2"
     
      Dim oDoc As Document
     
      Set oDoc = ThisApplication.ActiveDocument
      If oDoc Is Nothing Then
        MsgBox "Missing Inventor Document"
        Exit Sub
      End If
     
    'iLogicAuto.RunRule oDoc, INTERNALrule 'for internal rule
    'iLogicAuto.RunExternalRule oDoc, EXTERNALrule 'for external rule
    
    Dim Var
    Var = iLogicAuto.getRule(oDoc, "Hamid")
    'Var = iLogicAuto.AddRule(oDoc, "Hamid", "Msgbox(""aa"")")

End Sub

