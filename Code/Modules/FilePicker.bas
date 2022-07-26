Attribute VB_Name = "FilePicker"
'any const?

Sub testfp()
    Set c = FileDialogDisplayAndGetSelection()
    Debug.Print c.Count = 1
    
    Set c = FileDialogDisplayAndGetSelection(dialogRequest:="Cancel")
    Debug.Print c.Count = 0
    
    Set c = FileDialogDisplayAndGetSelection(selectMany:=True)
    Debug.Print c.Count > 1
    
End Sub

'
' purpose:
' Accepts:
' Returns:
Function FileDialogDisplayAndGetSelection( _
                                            Optional ByVal dialogRequest As String = "", _
                                            Optional ByVal dialogType As MsoFileDialogType = msoFileDialogFilePicker, _
                                            Optional ByVal selectMany As Boolean = False, _
                                            Optional ByVal filterVal As String = "" _
) As Collection


    Set FileDialogDisplayAndGetSelection = New Collection

    With Application.FileDialog(dialogType)
    
        .AllowMultiSelect = selectMany
        .Filters.Clear
        .Title = dialogRequest
        
        If filterVal = "" Then
        
            .Filters.Add "Excel Files", "*.xls?,*.csv"
            
        Else
        
            .Filters.Add "Files", filterVal
            
        End If
        
        If .Show = -1 Then
        
            For Each Item In .SelectedItems
            
                FileDialogDisplayAndGetSelection.Add Item
                
            Next
            
        End If
    
    End With
    
End Function
