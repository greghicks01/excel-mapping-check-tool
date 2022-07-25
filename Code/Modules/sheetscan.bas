Attribute VB_Name = "sheetscan"

Const defaultStopCondition As String = ""
Const defaultStartRow As Integer = 1
Const defaultStartCol As Integer = 1

'
' Purpose:
' Accepts:
' Returns:

'sheet scans

'
' Purpose: column wise scan for a condition
' Accepts:
' Returns:
Public Function scanColumnsForKeyOrUntilCondition( _
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
        
        scanColumnsForKeyOrUntilCondition = currentCol
    
    End With

End Function

'
' Purpose: column wise scan for a condition
' Accepts:
' Returns:
Public Function scanRowsForKeyOrUntilCondition( _
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
        
            .Cells(currentRow, currentCol).Select
        
            If .Cells(currentRow, currentCol) = currentStopCondition Then Exit Function
            
            currentRow = currentRow + 1
            'trap out of bounds
            If (currentRow > 1048576) Then Err.Raise 9
            
        Wend
        
        scanRowsForKeyOrUntilCondition = currentRow
    
    End With

End Function


