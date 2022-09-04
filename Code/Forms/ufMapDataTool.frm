VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufMapDataTool 
   Caption         =   "UserForm1"
   ClientHeight    =   8745.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7245
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
    both
End Enum

'====================================
' Properties
'====================================
Public Property Let mapdata(c As Collection)

    uiElement.loadListFromCollection Me.lbMapKeys, c
    
End Property

Public Property Let dataheaders(c As Collection)

    uiElement.loadListFromCollection Me.lbDataHeaders, c
    
End Property

'==========
' Events
'==========
Private Sub obBadChar_Click()
    
End Sub

Private Sub obBoth_Click()

End Sub

Private Sub obNone_Click()

End Sub

Private Sub obOverlength_Click()

End Sub


'========
' Functions
'=========

Private Function getSelectedFilter() As filterType
    Dim oControl As control
    
    With Me.filterFrame
        For Each oControl In .Controls
            Select Case oControl.Name
                Case "obNone"
                    getSelectedFilter = none
                    Exit Function
                    
                Case "obBadChar"
                    getSelectedFilter = badchar
                    Exit Function
                    
                Case "obOverlength"
                    getSelectedFilter = badlen
                    Exit Function
                    
                Case "obBoth"
                    getSelectedFilter = both
                    Exit Function
                    
            End Select
        Next
    End With
End Function

