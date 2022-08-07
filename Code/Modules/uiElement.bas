Attribute VB_Name = "uiElement"

Sub loadListFromCollection(element As Object, C_data As Collection, Optional ByVal filter As filterType = none)

    'Range check this is the right type of control
    t_name = TypeName(element)
    If t_name <> "ListBox" And t_name <> "ComboBox" Then Err.Raise 9, , "Expected 'ListBox' or 'ComboBox', got '" & t_name & "' instead"
    
    Application.EnableEvents = False
    
    element.Clear

    For Each c In C_data
        If filter Then element.AddItem c
    Next
    
    Application.EnableEvents = True
    
End Sub

Function filterValueByType(v As String, f As filterType) As Boolean

    Dim reg As New RegExp
    reg.Pattern = "[" & getConfigValue("bad char") & "]"
    hasBadChar = reg.Test(v)
    exceedsMaxLenth = Len(v) > CInt(getConfigValue("Max Len"))

    'return True if filter met
    Select Case f
        Case none
            filterValueByType = True
            Exit Function
            
        Case badchar
            filterValueByType = hasBadChar
            Exit Function
            
        Case badlen
            filterValueByType = exceedsMaxLenth
            Exit Function
            
        Case badcharlen
            filterValueByType = exceedsMaxLenth Or hasBadChar
            Exit Function
            
    End Select
    
    ' error

End Function

