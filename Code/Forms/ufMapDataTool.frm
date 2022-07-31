VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufMapDataTool 
   Caption         =   "UserForm1"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3240
   OleObjectBlob   =   "ufMapDataTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufMapDataTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================
' Properties
'====================================
Public Property Let mapdata(c As Collection)

    uiElement.loadListFromCollection Me.lbMapKeys, c
    
End Property

