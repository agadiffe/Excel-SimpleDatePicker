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


Private Sub Workbook_SheetBeforeDoubleClick( _
    ByVal Sh As Object, _
    ByVal Target As Range, _
    Cancel As Boolean)

    ' Only process a single cell
    If Target.Cells.CountLarge <> 1 Then Exit Sub

    ' Do not open the date picker if the cell has a formula
    If Target.Cells(1).HasFormula Then Exit Sub

    ' Only open the date picker for supported date-formatted cells
    If Not IsShortDateCell(Target) Then Exit Sub

    ' Prevent Excel's normal double-click action
    Cancel = True

    ' Show the date picker
    DP_frmDatePicker.ShowPicker Target

End Sub


Private Function IsShortDateCell(ByVal Cell As Range) As Boolean

    Dim FormatType As Variant

    On Error GoTo NotDate

    FormatType = Cell.Parent.Evaluate( _
        "=CELL(""format""," & Cell.Address(False, False) & ")")

    ' Excel returns a format code beginning with "D"
    ' for supported date formats.
    '
    ' Long-date and custom date-format handling is not
    ' explicitly distinguished here.

    IsShortDateCell = _
        (Left$(UCase$(Trim$(CStr(FormatType))), 1) = "D")

    Exit Function

NotDate:

    IsShortDateCell = False

End Function

