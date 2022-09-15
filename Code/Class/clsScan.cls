VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum Direction
    down
    across
End Enum

Public currentRow, currentCol As Integer
Private tmpRow, tmpCol As Integer
Public control As Integer
Public dir As Direction
Public sheet As Worksheet
Public stopCondition As String
Private currentValue As String

Private Sub Class_Initialize()

    currentRow = 1
    currentCol = 1
    tmpRow = 1
    tmpCol = 1
    control = 0
    dir = down
    
End Sub

Property Get nextcell() As Boolean

    tmpRow = currentRow
    tmpCol = currentCol
    
    Select Case dir
        Case Is = down
            tmpRow = IIf(control <> 0, control, currentRow)
            
        Case Is = across
            tmpCol = IIf(control <> 0, control, currentCol)
            
    End Select
    
    currentValue = sh.Cells(tmpRow, tmpCol)
    nextcell = Not (currentValue = stopCondition)
    
    If Not nextcell Then Exit Property
    
    Select Case dir
        Case Is = down
            currentRow = currentRow + 1
            
        Case Is = across
            currentCol = currentCol + 1
            
    End Select
    
End Property

Property Get locateValue(ByVal v As String) As Integer

    stopCriteria = v
    
    While nextcell()
    Wend
    
    locateValue = IIf(Direction = down, tmpRow, tmpCol)
    
End Property
