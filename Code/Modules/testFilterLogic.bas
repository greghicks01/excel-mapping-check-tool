Attribute VB_Name = "testFilterLogic"
Sub testfilter()


    Debug.Print uiElement.filterValueByType("string", filterType.none) = True
    Debug.Print uiElement.filterValueByType("asc@gfd", filterType.none) = True
    
    Debug.Print uiElement.filterValueByType("asc@gfd", filterType.badchar) = True
    Debug.Print uiElement.filterValueByType("ascgfd", filterType.badchar) = False
    
    Debug.Print uiElement.filterValueByType("1234567890123456789012345678901234567890123456789012345678901", filterType.badlen) = True
    Debug.Print uiElement.filterValueByType("123456789012345678901234567890123456789012345678901234567890", filterType.badlen) = False
    
    Debug.Print uiElement.filterValueByType("12345678901234567890123456789012345678901234567890123456789@", filterType.badcharlen) = True
    Debug.Print uiElement.filterValueByType("123456789012345678901234567890123456789012345678901234567891", filterType.badcharlen) = False
    
End Sub

