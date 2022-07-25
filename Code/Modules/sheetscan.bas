Attribute VB_Name = "sheetscan"

Const defaultStopCondition As String = ""
Const defaultStartRow As Integer = 1
Const defaultStartCol As Integer = 1

'
' Purpose:
' Accepts:
' Returns:

' Common accepts values
'
' ws As Worksheet reference for the sheet to scan
' key As String value to find
' stopCondition As String what to stop the scan and prevent overrun
' startRow As Integer default or user controlled starting row for scan
' startCol As Integer default or user controlled starting column for scan
'


'
' Purpose: column wise scan for a condition
' Accepts:
' Returns:
Public Function scanColumnsForKeyOrConditionFound( _
                                                    ByRef ws As Worksheet, _
                                                    ByVal key As String, _
                                                    Optional ByVal stopCondition As String = defaultStopCondition, _
                                                    Optional ByVal startRow As Integer = defaultStartRow, _
                                                    Optional ByVal startCol As Integer = defaultStartCol _
                                                 ) As Integer
    currentKey = key
    currentRow = startRow
    currentCol = startCol
    currentStopCondition = stopCondition

    ' trap out of bounds conditions
    If IIf(currentRow < 1, True, IIf(currentRow > 1048576, True, False)) Then Err.Raise 9
    If IIf(currentCol < 1, True, IIf(currentCol > 16384, True, False)) Then Err.Raise 9
    If key = "" Then Err.Raise 9
                                           
    With ws
    
        While .Cells(currentRow, currentCol) <> currentKey
        
            If .Cells(currentRow, currentCol) = currentStopCondition Then Exit Function
            currentCol = currentCol + 1
            'trap out of bounds
            If (currentCol > 16384) Then Err.Raise 9
            
        Wend
        
        scanColumnsForKeyOrConditionFound = currentCol
    
    End With

End Function

'
' Purpose: column wise scan for a condition
' Accepts:
' Returns:
Public Function scanRowsForKeyOrConditionFound( _
                                                    ByRef ws As Worksheet, _
                                                    ByVal key As String, _
                                                    Optional ByVal stopCondition As String = defaultStopCondition, _
                                                    Optional ByVal startRow As Integer = defaultStartRow, _
                                                    Optional ByVal startCol As Integer = defaultStartCol _
                                                 ) As Integer
    currentKey = key
    currentRow = startRow
    currentCol = startCol
    currentStopCondition = stopCondition

    ' trap out of bounds conditions
    If IIf(currentRow < 1, True, IIf(currentRow > 1048576, True, False)) Then Err.Raise 9
    If IIf(currentCol < 1, True, IIf(currentCol > 16384, True, False)) Then Err.Raise 9
    If key = "" Then Err.Raise 9
                                           
    With ws
        
        While .Cells(currentRow, currentCol) <> currentKey
                
            If .Cells(currentRow, currentCol) = currentStopCondition Then Exit Function
            
            currentRow = currentRow + 1
            'trap out of bounds
            If (currentRow > 1048576) Then Err.Raise 9
            
        Wend
        
        scanRowsForKeyOrConditionFound = currentRow
    
    End With

End Function

'
' Purpose: column wise scan for a condition
' Accepts:
' Returns: Compacted collection
Public Function scanRowsForKeysUntilConditionFound( _
                                                    ByRef ws As Worksheet, _
                                                    Optional ByVal stopCondition As String = defaultStopCondition, _
                                                    Optional ByVal startRow As Integer = defaultStartRow, _
                                                    Optional ByVal startCol As Integer = defaultStartCol, _
                                                    Optional ByVal controlCol As Integer = defaultStartCol _
                                                  ) As Collection
    Dim c As New Collection
    
    currentRow = startRow
    currentCol = startCol
    currentControlCol = controlCol
    currentStopCondition = stopCondition
    
    ' trap out of bounds conditions
    If IIf(currentRow < 1, True, IIf(currentRow > 1048576, True, False)) Then Err.Raise 9
    If IIf(currentCol < 1, True, IIf(currentCol > 16384, True, False)) Then Err.Raise 9
    
    With ws
        
        While .Cells(currentRow, currentControlCol) <> currentStopCondition
                
            If .Cells(currentRow, currentCol) <> "" Then
                c.Add .Cells(currentRow, currentCol)
            End If
            
            currentRow = currentRow + 1
            'trap out of bounds
            If (currentRow > 1048576) Then Err.Raise 9
            
        Wend
        
        Set scanRowsForKeysUntilConditionFound = c
    
    End With

End Function
