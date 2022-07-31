Attribute VB_Name = "TestUIElement"
Sub testui()

    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Sheets("testmapload")

    ' this used the format in the existing sheet to execute the run
    UserForm1.mapdata = sheetscan.scanRowsForKeysUntilConditionFound(ws, , 3, 2)
    
    UserForm1.Show

End Sub
