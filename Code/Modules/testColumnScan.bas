Attribute VB_Name = "testColumnScan"
Sub testColumnScan()

    Dim sh As Worksheet

    Set sh = Application.ActiveWorkbook.Sheets("testcolumnscan")
    
    Debug.Print sheetscan.scanColumnsForKeyOrUntilCondition(sh, "test") = 2
    
    Debug.Print sheetscan.scanColumnsForKeyOrUntilCondition(sh, "test", startCol:=4) = 5
    
    Debug.Print sheetscan.scanColumnsForKeyOrUntilCondition(sh, "test", startRow:=3) = 2
    
    Debug.Print sheetscan.scanColumnsForKeyOrUntilCondition(sh, "test", startRow:=5, startCol:=3) = 4

    Debug.Print sheetscan.scanColumnsForKeyOrUntilCondition(sh, "test", startRow:=7) = 0

    Debug.Print sheetscan.scanColumnsForKeyOrUntilCondition(sh, "test", "stop", 9) = 0
    
    On Error Resume Next
    Debug.Assert sheetscan.scanColumnsForKeyOrUntilCondition(sh, "") = 0
    
    Debug.Assert sheetscan.scanColumnsForKeyOrUntilCondition(sh, "test", startRow:=0) = 0
    
    Debug.Assert sheetscan.scanColumnsForKeyOrUntilCondition(sh, "test", startCol:=0) = 0
    
    Debug.Assert sheetscan.scanColumnsForKeyOrUntilCondition(sh, "test", startRow:=1500000) = 0
    
    Debug.Assert sheetscan.scanColumnsForKeyOrUntilCondition(sh, "test", startCol:=20000) = 0
    On Error GoTo 0
    

End Sub

Sub testRowScan()

    Dim sh As Worksheet

    Set sh = Application.ActiveWorkbook.Sheets("testrowscan")
    
    Debug.Print sheetscan.scanRowsForKeyOrUntilCondition(sh, "test") = 2
    
    Debug.Print sheetscan.scanRowsForKeyOrUntilCondition(sh, "test", startCol:=3) = 2
    
    Debug.Print sheetscan.scanRowsForKeyOrUntilCondition(sh, "test", startRow:=4) = 5
    
    Debug.Print sheetscan.scanRowsForKeyOrUntilCondition(sh, "test", startRow:=3, startCol:=5) = 4

    Debug.Print sheetscan.scanRowsForKeyOrUntilCondition(sh, "test", startCol:=7) = 0

    Debug.Print sheetscan.scanRowsForKeyOrUntilCondition(sh, "test", "stop", startCol:=9) = 0
    
    On Error Resume Next
    Debug.Assert sheetscan.scanRowsForKeyOrUntilCondition(sh, "") = 0
    
    Debug.Assert sheetscan.scanRowsForKeyOrUntilCondition(sh, "test", startRow:=0) = 0
    
    Debug.Assert sheetscan.scanRowsForKeyOrUntilCondition(sh, "test", startCol:=0) = 0
    
    Debug.Assert sheetscan.scanRowsForKeyOrUntilCondition(sh, "test", startRow:=1500000) = 0
    
    Debug.Assert sheetscan.scanRowsForKeyOrUntilCondition(sh, "test", startCol:=20000) = 0
    On Error GoTo 0
    

End Sub
