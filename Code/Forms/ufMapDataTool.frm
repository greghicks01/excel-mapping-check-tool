VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufMapDataTool 
   Caption         =   "UserForm1"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   OleObjectBlob   =   "ufMapDataTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufMapDataTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Enum filterType
    none
    badchar
    badlen
    badcharlen
End Enum

Private Sub addFiltersToFrame()

End Sub

'====================================
' Properties
'====================================
Public Property Let mapdata(c As Collection)

    uiElement.loadListFromCollection Me.lbMapKeys, c
    
End Property

Public Property Let dataheaders(c As Collection)

    uiElement.loadListFromCollection Me.lbDataHeaders, c
    
End Property

Private Sub obBadChar_Click()

End Sub

Private Sub obBoth_Click()

End Sub

Private Sub obNone_Click()

End Sub

Private Sub obOverlength_Click()

End Sub

Private Sub filteredLoad(Optional ByVal filter As filterType = none)

End Sub


