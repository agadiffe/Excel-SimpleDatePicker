Attribute VB_Name = "DP_modEntryPoint"
Option Explicit

' Project repository: https://github.com/agadiffe/Excel-SimpleDatePicker

'----------------------------------------
' Activation
'----------------------------------------

Public Function DP_TryShowDatePicker(ByVal Target As Range) As Boolean

    ' Only process a single cell
    If Target.Cells.CountLarge <> 1 Then Exit Function

    ' Do not open the date picker if the cell has a formula
    If Target.HasFormula Then Exit Function

    ' Only open the date picker for supported date cells
    If Not IsDatePickerCell(Target) Then Exit Function

    ' Existing date must be within the supported picker range
    If Not IsEmpty(Target) Then
        If DP_IsDateOutsideRange(CDate(Target.Value)) Then Exit Function
    End If

    ' Show the date picker
    DP_GetPickerForm(Excel.Application.ActiveWindow).ShowPicker Target

    DP_TryShowDatePicker = True

End Function


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

    ' Cell's number format string
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

