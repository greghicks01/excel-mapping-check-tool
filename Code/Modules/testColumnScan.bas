Attribute VB_Name = "testColumnScan"
Sub testColumnScan()

    Dim sh As Worksheet

    Set sh = Application.ActiveWorkbook.Sheets("testcolumnscan")
    
    Debug.Print sheetscan.scanColumnsForKeyOrConditionFound(sh, "test") = 2
    
    Debug.Print sheetscan.scanColumnsForKeyOrConditionFound(sh, "test", startCol:=4) = 5
    
    Debug.Print sheetscan.scanColumnsForKeyOrConditionFound(sh, "test", startRow:=3) = 2
    
    Debug.Print sheetscan.scanColumnsForKeyOrConditionFound(sh, "test", startRow:=5, startCol:=3) = 4

    Debug.Print sheetscan.scanColumnsForKeyOrConditionFound(sh, "test", startRow:=7) = 0

    Debug.Print sheetscan.scanColumnsForKeyOrConditionFound(sh, "test", "stop", 9) = 0
    
    On Error Resume Next
    Debug.Print sheetscan.scanColumnsForKeyOrConditionFound(sh, "") = 0
    
    Debug.Print sheetscan.scanColumnsForKeyOrConditionFound(sh, "test", startRow:=0) = 0
    
    Debug.Print sheetscan.scanColumnsForKeyOrConditionFound(sh, "test", startCol:=0) = 0
    
    Debug.Print sheetscan.scanColumnsForKeyOrConditionFound(sh, "test", startRow:=1500000) = 0
    
    Debug.Print sheetscan.scanColumnsForKeyOrConditionFound(sh, "test", startCol:=20000) = 0
    On Error GoTo 0
    

End Sub

Sub testRowScan()

    Dim sh As Worksheet

    Set sh = Application.ActiveWorkbook.Sheets("testrowscan")
    
    Debug.Print sheetscan.scanRowsForKeyOrConditionFound(sh, "test") = 2
    
    Debug.Print sheetscan.scanRowsForKeyOrConditionFound(sh, "test", startCol:=3) = 2
    
    Debug.Print sheetscan.scanRowsForKeyOrConditionFound(sh, "test", startRow:=4) = 5
    
    Debug.Print sheetscan.scanRowsForKeyOrConditionFound(sh, "test", startRow:=3, startCol:=5) = 4

    Debug.Print sheetscan.scanRowsForKeyOrConditionFound(sh, "test", startCol:=7) = 0

    Debug.Print sheetscan.scanRowsForKeyOrConditionFound(sh, "test", "stop", startCol:=9) = 0
    
    On Error Resume Next
    Debug.Print sheetscan.scanRowsForKeyOrConditionFound(sh, "") = 0
    
    Debug.Print sheetscan.scanRowsForKeyOrConditionFound(sh, "test", startRow:=0) = 0
    
    Debug.Print sheetscan.scanRowsForKeyOrConditionFound(sh, "test", startCol:=0) = 0
    
    Debug.Print sheetscan.scanRowsForKeyOrConditionFound(sh, "test", startRow:=1500000) = 0
    
    Debug.Print sheetscan.scanRowsForKeyOrConditionFound(sh, "test", startCol:=20000) = 0
    On Error GoTo 0
    
End Sub

Sub testcontrolcol()
    Dim sh As Worksheet
    
    Set sh = Application.ActiveWorkbook.Sheets("testcontrolcolumn")
    
    Debug.Print scanRowsForKeysUntilConditionFound(sh, startRow:=3, startCol:=2).Count = 6
    
    Debug.Print scanRowsForKeysUntilConditionFound(sh, startRow:=16, startCol:=6, controlCol:=5).Count = 5
    
End Sub
