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
    If Target.HasFormula Then Exit Sub

    ' Only open the date picker for supported date cells
    If Not IsDatePickerCell(Target) Then Exit Sub

    ' Prevent Excel's normal double-click action
    Cancel = True

    ' Show the date picker
    DP_frmDatePicker.ShowPicker Target

End Sub


Private Function IsDatePickerCell(ByVal Cell As Range) As Boolean

    Dim FormatType As String
    Dim NumFormat As String

    On Error GoTo NotDate

    ' Standard Excel date formats D1-D5
    FormatType = UCase$(Trim$(CStr( _
        Cell.Parent.Evaluate( _
            "=CELL(""format""," & Cell.Address(False, False) & ")"))))

    Select Case FormatType
        Case "D1", "D2", "D3", "D4", "D5"
            IsDatePickerCell = True
            Exit Function
    End Select

    ' Common Long Date formats
    NumFormat = LCase$(Cell.NumberFormat)

    ' Excel's system Long Date format
    If Left$(NumFormat, Len("[$-f800]")) = "[$-f800]" Then
        IsDatePickerCell = True
        Exit Function
    End If

    ' Common Long Date formats without a system-locale prefix
    Select Case NumFormat
        Case "dddd, mmmm d, yyyy", _
             "dddd, mmmm dd, yyyy", _
             "mmmm d, yyyy", _
             "d mmmm yyyy", _
             "dd mmmm yyyy", _
             "dddd d mmmm yyyy", _
             "dddd, d mmmm yyyy", _
             "dddd, dd mmmm yyyy", _
             "d. mmmm yyyy", _
             "dd. mmmm yyyy", _
             "d-mmmm-yyyy", _
             "dd-mmmm-yyyy"

            IsDatePickerCell = True
            Exit Function
    End Select

NotDate:
    IsDatePickerCell = False

End Function

