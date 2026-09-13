VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Excel.Application
Attribute App.VB_VarHelpID = -1


Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, _
                                       ByVal Target As Range, _
                                       Cancel As Boolean)

    If DP_HandleDatePickerDoubleClick(Target) Then
        Cancel = True
    End If

End Sub


