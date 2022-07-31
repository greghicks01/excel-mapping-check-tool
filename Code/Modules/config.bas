Attribute VB_Name = "config"
Const sheetname = "config"
Const key = "Key"
Const value = "Value"
Const comment = "Comment"
Const headerstart = "A1"
Const datastart = "A2"
Private ws As Worksheet
Private r As Range

Function getConfigValue(key As String) As String
    
    Set ws = ActiveWorkbook.Sheets(sheetname)
    Set r = Range(datastart)
    
    ' get key row
    kn = sheetscan.scanRowsForKeyOrConditionFound(ws, key, , r.Row)
    ' scan for key
    v = sheetscan.scanColumnsForKeyOrConditionFound(ws, value)
    getConfigValue = ws.Cells(kn, v)
    
End Function
