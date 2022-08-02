Attribute VB_Name = "TestUIElement"
Sub testui()

    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Sheets("testmapload")

    ' this used the format in the existing sheet to execute the run
    ufMapDataTool.mapdata = sheetscan.scanRowsForKeysUntilConditionFound(ws, , 3, 2)
    
    ufMapDataTool.Show

End Sub

Sub testdata()

    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Sheets("testdataload")

    ' this used the format in the existing sheet to execute the run
    ufMapDataTool.dataheaders = sheetscan.scanColumnsConditionFound(ws)
    
    ufMapDataTool.Show

End Sub

Sub testnameedrange()
    c = 0
    For Each nme In ActiveWorkbook.Names
        If InStr(nme.Name, "_xlfn") = 0 Then
            c = c + 1
            With ActiveWorkbook.Sheets("myrangenames")
                .Cells(c, 1) = CStr(nme.Name)
                .Cells(c, 2) = "'" + CStr(nme.RefersTo)
            End With
        End If
    Next
End Sub

Sub testResetNamedranges()
    c = 0
    For Each rw In ActiveWorkbook.Sheets("myrangenames").UsedRange.Rows
        rw.Select
        c = c + 1
        ActiveWorkbook.Names.Add rw.Cells(1), Replace(rw.Cells(2), "'", "")
    Next
End Sub
