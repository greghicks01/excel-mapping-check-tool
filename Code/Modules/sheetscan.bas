Attribute VB_Name = "sheetscan"

Const defaultStopCondition As String = ""
Const defaultStartRow As Integer = 1
Const defaultStartCol As Integer = 1

' Enum to allow a single range check with optional param
' to pick max range values in row or columns
Enum RowOrColumnWise
    Row
    Column
End Enum

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
' Accepts: common params above
' Returns: column as an integer to show where key was found
Public Function scanColumnsForKeyOrConditionFound( _
                                                    ByRef ws As Worksheet, _
                                                    ByVal key As String, _
                                                    Optional ByVal stopCondition As String = defaultStopCondition, _
                                                    Optional ByVal startRow As Integer = defaultStartRow, _
                                                    Optional ByVal startCol As Integer = defaultStartCol _
                                                 ) As Integer
    'Isolate the incoming values exept the byREF
    currentKey = key
    currentRow = startRow
    currentCol = startCol
    currentStopCondition = stopCondition

    ' trap out of bounds conditions
     If Not sheetRangeCheck(currentRow) Or Not sheetRangeCheck(currentCol, Column) Then Err.Raise 9
    If key = "" Then Err.Raise 9
                                           
    With ws
    
        While .Cells(currentRow, currentCol) <> currentKey
        
            If .Cells(currentRow, currentCol) = currentStopCondition Then Exit Function
            currentCol = currentCol + 1
            'trap out of bounds
            If Not sheetRangeCheck(currentCol, Column) Then Err.Raise 9
            
        Wend
        
        scanColumnsForKeyOrConditionFound = currentCol
    
    End With

End Function

'
' Purpose: column wise collection of headers
' Accepts: common params above
' Returns: column as an integer to show where key was found
Public Function scanColumnsConditionFound( _
                                                    ByRef ws As Worksheet, _
                                                    Optional ByVal stopCondition As String = defaultStopCondition, _
                                                    Optional ByVal startRow As Integer = defaultStartRow, _
                                                    Optional ByVal startCol As Integer = defaultStartCol _
                                                 ) As Collection
    'Isolate the incoming values exept the byREF
    currentRow = startRow
    currentCol = startCol
    currentStopCondition = stopCondition
    
    Dim c As New Collection

    ' trap out of bounds conditions
    If Not sheetRangeCheck(currentRow) Or Not sheetRangeCheck(currentCol, RowOrColumnWise.Column) Then Err.Raise 9
                                           
    With ws
    
        While .Cells(currentRow, currentCol) <> currentStopCondition
            
            c.Add .Cells(currentRow, currentCol)
            currentCol = currentCol + 1
            'trap out of bounds
            If Not sheetRangeCheck(currentCol, Column) Then Err.Raise 9
            
        Wend
    
    End With
    
    Set scanColumnsConditionFound = c

End Function

'
' Purpose: row wise scan for a condition
' Accepts: common params above
' Returns: column as an integer to show where key was found
Public Function scanRowsForKeyOrConditionFound( _
                                                ByRef ws As Worksheet, _
                                                ByVal key As String, _
                                                Optional ByVal stopCondition As String = defaultStopCondition, _
                                                Optional ByVal startRow As Integer = defaultStartRow, _
                                                Optional ByVal startCol As Integer = defaultStartCol _
                                              ) As Integer
    'Isolate the incoming values exept the byREF
    currentKey = key
    currentRow = startRow
    currentCol = startCol
    currentStopCondition = stopCondition

    ' trap out of bounds conditions
    If Not sheetRangeCheck(currentRow) Or Not sheetRangeCheck(currentCol, Column) Then Err.Raise 9
    If key = "" Then Err.Raise 9
                                           
    With ws
        
        While .Cells(currentRow, currentCol) <> currentKey
                
            If .Cells(currentRow, currentCol) = currentStopCondition Then Exit Function
            
            currentRow = currentRow + 1
            'trap out of bounds
            If Not sheetRangeCheck(currentRow) Then Err.Raise 9
            
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

    'Isolate the incoming values exept the byREF
    currentRow = startRow
    currentCol = startCol
    currentControlCol = controlCol
    currentStopCondition = stopCondition

    ' trap out of bounds conditions
    If Not sheetRangeCheck(currentRow) Or Not sheetRangeCheck(currentCol, Column) Then Err.Raise 9
    
    With ws
        
        While .Cells(currentRow, currentControlCol) <> currentStopCondition
                
            If .Cells(currentRow, currentCol) <> "" Then
                c.Add .Cells(currentRow, currentCol)
            End If
            
            currentRow = currentRow + 1
            'trap out of bounds
            If Not sheetRangeCheck(currentRow) Then Err.Raise 9
            
        Wend
        
        Set scanRowsForKeysUntilConditionFound = c
    
    End With

End Function

'
' Purpose: Range checks incoming values, easier to maintain in the longer run
' Accepts: the index to check and the option Row or Column used to pick the maximum value
' Returns: True is in Range, False if out of range
Public Function sheetRangeCheck(ByVal idx As Integer, Optional rowCol As RowOrColumnWise = RowOrColumnWise.Row) As Boolean

    Select Case rowCol
        Case RowOrColumnWise.Row
            maxVal = 1048576
            
        Case RowOrColumnWise.Column
            maxVal = 16384
            
    End Select
    
    sheetRangeCheck = IIf(idx >= 1 And idx <= maxVal, True, False)
    
End Function
