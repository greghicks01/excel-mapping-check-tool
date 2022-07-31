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

