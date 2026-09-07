VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, _
                                            ByVal Target As Range, _
                                            Cancel As Boolean)

    If DP_HandleDatePickerDoubleClick(Target) Then
        Cancel = True ' Prevent other Excel actions
    End If

End Sub

