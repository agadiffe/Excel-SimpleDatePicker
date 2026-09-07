Attribute VB_Name = "DP_modHelpers"
Option Explicit

'----------------------------------------
' Date Picker
'----------------------------------------

Public Function DP_HandleDatePickerDoubleClick(ByVal Target As Range) As Boolean

    ' Only process a single cell
    If Target.Cells.CountLarge <> 1 Then Exit Function

    ' Do not open the date picker if the cell has a formula
    If Target.HasFormula Then Exit Function

    ' Only open the date picker for supported date cells
    If Not DP_IsDatePickerCell(Target) Then Exit Function

    ' Show the date picker
    DP_frmDatePicker.ShowPicker Target

    DP_HandleDatePickerDoubleClick = True

End Function


Public Function DP_IsDatePickerCell(ByVal Cell As Range) As Boolean

    Dim FormatType As String
    Dim NumFormat As String

    On Error GoTo NotDate

    ' Standard Excel date formats D1-D5
    FormatType = UCase$(Trim$(CStr( _
        Cell.Parent.Evaluate( _
            "=CELL(""format""," & Cell.Address(False, False) & ")"))))

    Select Case FormatType
        Case "D1", "D2", "D3", "D4", "D5"
            DP_IsDatePickerCell = True
            Exit Function
    End Select

    ' Common Long Date formats
    NumFormat = LCase$(Cell.NumberFormat)

    ' Excel's system Long Date format
    If Left$(NumFormat, Len("[$-f800]")) = "[$-f800]" Then
        DP_IsDatePickerCell = True
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

            DP_IsDatePickerCell = True
            Exit Function
    End Select

NotDate:
    DP_IsDatePickerCell = False

End Function


'----------------------------------------
' Picker controls
'----------------------------------------

Public Sub DP_ClearPickerControls(ByVal PickerForm As Object)

    Dim ControlIndex As Long

    For ControlIndex = PickerForm.Controls.Count - 1 To 0 Step -1
        PickerForm.Controls.Remove PickerForm.Controls(ControlIndex).Name
    Next ControlIndex

End Sub


Public Sub DP_InitializePicker(ByVal PickerForm As Object, _
                               ByVal FormWidth As Single, _
                               ByVal FormHeight As Single)

    With PickerForm
        .Width = FormWidth
        .Height = FormHeight

        .Width = .Width + (FormWidth - .InsideWidth)
        .Height = .Height + (FormHeight - .InsideHeight)

        .BackColor = DP_ColorBg()
    End With

End Sub


'----------------------------------------
' Layout
'----------------------------------------

Public Function DP_VerticalTextTop(ByVal ContainerTop As Single, _
                                   ByVal ContainerHeight As Single, _
                                   ByVal TextHeight As Single) As Single

    Const TEXT_VERTICAL_CORRECTION As Single = 0.033

    DP_VerticalTextTop = ContainerTop + _
                         (ContainerHeight - TextHeight) / 2 - _
                         ContainerHeight * TEXT_VERTICAL_CORRECTION

End Function


Public Function DP_GetFontScale(ByVal CurrentWidth As Single) As Single

    DP_GetFontScale = CurrentWidth / DP_DATEPICKER_BASE_WIDTH

End Function


Public Function DP_GridColumn(ByVal CellIndex As Long, _
                              ByVal GridColumns As Long) As Long

    DP_GridColumn = CellIndex Mod GridColumns

End Function


Public Function DP_GridRow(ByVal CellIndex As Long, _
                           ByVal GridColumns As Long) As Long

    DP_GridRow = CellIndex \ GridColumns

End Function


Public Function DP_GridLeft(ByVal ColumnIndex As Long, _
                            ByVal GridLeftPosition As Single, _
                            ByVal GridCellWidth As Single, _
                            ByVal CellWidth As Single) As Single

    DP_GridLeft = GridLeftPosition + _
                  ColumnIndex * GridCellWidth + _
                  (GridCellWidth - CellWidth) / 2

End Function


Public Function DP_GridTop(ByVal RowIndex As Long, _
                           ByVal GridTopPosition As Single, _
                           ByVal GridCellHeight As Single) As Single

    DP_GridTop = GridTopPosition + RowIndex * GridCellHeight

End Function


'----------------------------------------
' Reset hover
'----------------------------------------

Public Sub DP_ResetHoverCollection(ByVal Handlers As Collection)

    Dim Handler As Object

    For Each Handler In Handlers
        Handler.ResetHover
    Next Handler

End Sub


Public Sub DP_ResetHoverCollectionExcept(ByVal Handlers As Collection, _
                                         ByVal KeepHandler As Object)

    Dim Handler As Object

    For Each Handler In Handlers
        If Not Handler Is KeepHandler Then
            Handler.ResetHover
        End If
    Next Handler

End Sub

