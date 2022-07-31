Attribute VB_Name = "uiElement"

Sub loadListFromCollection(element As Object, C_data As Collection)

    'Range check this is the right type of control
    t_name = TypeName(element)
    If t_name <> "ListBox" And t_name <> "ComboBox" Then Err.Raise 9, , "Expected 'ListBox' or 'ComboBox', got '" & t_name & "' instead"

    For Each c In C_data
        element.AddItem c
    Next
    
End Sub
