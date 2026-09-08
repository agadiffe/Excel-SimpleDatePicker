VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DP_frmYearPicker 
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "DP_frmYearPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DP_frmYearPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' References
'----------------------------------------

Private Const YEAR_COUNT As Long = _
    DP_YEAR_GRID_COLUMNS * DP_YEAR_GRID_ROWS

Private ParentPicker As DP_frmMonthPicker
Private CurrentYear As Long

Private YearsInitialized As Boolean

Private ArrowHandlers As Collection
Private YearLabelHandlers As Collection
Private YearRangeLabelHandler As DP_CHeaderLabel
Private CurrentYearButtonHandler As DP_CActionButton


'----------------------------------------
' Year picker
'----------------------------------------

Public Sub ShowYears(ByVal Picker As DP_frmMonthPicker)

    Set ParentPicker = Picker

    CurrentYear = Picker.GetCurrentYear

    BuildYears
    DP_ShowPopupNextToCell Me, DP_frmDatePicker.GetTargetCell

End Sub


Private Sub BuildYears()

    If Not YearsInitialized Then
        InitializeYears
        YearsInitialized = True
    ElseIf DP_DarkMode <> (Me.BackColor = DP_ColorBgDark) Then
        InitializeYears
    Else
        RefreshYears
    End If

End Sub


Private Sub InitializeYears()

    Dim ArrowHandler As DP_CArrowLabel
    Dim FirstYear As Long
    Dim LastYear As Long

    Set YearLabelHandlers = New Collection
    Set ArrowHandlers = New Collection

    DP_ClearPickerControls Me
    DP_InitializePicker Me, DP_YEARPICKER_WIDTH, DP_YEARPICKER_HEIGHT

    ' Year range
    '--------------------

    FirstYear = GetFirstDisplayedYear(CurrentYear)
    LastYear = FirstYear + YEAR_COUNT - 1

    Set YearRangeLabelHandler = New DP_CHeaderLabel
    YearRangeLabelHandler.Setup Me, CStr(FirstYear) & " - " & CStr(LastYear)

    YearRangeLabelHandler.SetInteractive False

    ' Previous year range
    '--------------------

    Set ArrowHandler = New DP_CArrowLabel
    ArrowHandler.Setup Me, "PREV_RANGE_YEAR", ChrW(&H25B2), DP_ARROW_PREV_LEFT

    ArrowHandlers.Add ArrowHandler

    ' Next year range
    '--------------------

    Set ArrowHandler = New DP_CArrowLabel
    ArrowHandler.Setup Me, "NEXT_RANGE_YEAR", ChrW(&H25BC), DP_ARROW_NEXT_LEFT

    ArrowHandlers.Add ArrowHandler

    ' Years
    '--------------------

    CreateYearLabels

    ' Go To Current Year button
    '--------------------

    Set CurrentYearButtonHandler = New DP_CActionButton
    CurrentYearButtonHandler.Setup Me, "GO_TO_CURRENT_YEAR"

End Sub


Private Sub RefreshYears()

    ResetAllHover

    UpdateYearRangeHeader
    UpdateYearLabels
    UpdateArrowStates

End Sub


Private Sub UpdateYearRangeHeader()

    Dim FirstYear As Long
    Dim LastYear As Long

    FirstYear = GetFirstDisplayedYear(CurrentYear)
    LastYear = FirstYear + YEAR_COUNT - 1

    YearRangeLabelHandler.SetCaption CStr(FirstYear) & " - " & CStr(LastYear)

End Sub


'----------------------------------------
' Build Years helpers
'----------------------------------------

Private Sub CreateYearLabels()

    Dim YearHandler As DP_CPeriodLabel
    Dim CellIndex As Long
    Dim DisplayYear As Long
    Dim FirstYear As Long
    Dim CellDate As Date

    Dim ColumnIndex As Long
    Dim RowIndex As Long

    FirstYear = GetFirstDisplayedYear(CurrentYear)

    For CellIndex = 0 To YEAR_COUNT - 1

        DisplayYear = FirstYear + CellIndex
        CellDate = DateSerial(DisplayYear, 1, 1)

        ColumnIndex = DP_GridColumn(CellIndex, DP_YEAR_GRID_COLUMNS)
        RowIndex = DP_GridRow(CellIndex, DP_YEAR_GRID_COLUMNS)

        Set YearHandler = New DP_CPeriodLabel

        YearHandler.Setup Me, _
                          "YEAR", _
                          DisplayYear, _
                          CellDate, _
                          DP_GridLeft(ColumnIndex, _
                                      DP_PERIOD_GRID_LEFT, _
                                      DP_YEAR_GRID_CELL_WIDTH, _
                                      DP_YEAR_CELL_WIDTH), _
                          DP_GridTop(RowIndex, _
                                     DP_PERIOD_GRID_TOP, _
                                     DP_PERIOD_GRID_CELL_HEIGHT), _
                          DP_YEAR_CELL_WIDTH

        YearHandler.SetState IsCurrentYear(DisplayYear), _
                             IsSelectedYear(DisplayYear), _
                             False, _
                             IsPeriodOutsideRange(CellDate)

        YearLabelHandlers.Add YearHandler

    Next CellIndex

End Sub


Private Sub UpdateYearLabels()

    Dim CellIndex As Long
    Dim DisplayYear As Long
    Dim FirstYear As Long
    Dim CellDate As Date

    Dim YearHandler As DP_CPeriodLabel

    FirstYear = GetFirstDisplayedYear(CurrentYear)

    For CellIndex = 0 To YEAR_COUNT - 1

        DisplayYear = FirstYear + CellIndex
        CellDate = DateSerial(DisplayYear, 1, 1)

        Set YearHandler = YearLabelHandlers.Item(CellIndex + 1)

        YearHandler.UpdatePeriod CellDate, _
                                 DisplayYear, _
                                 IsCurrentYear(DisplayYear), _
                                 IsSelectedYear(DisplayYear), _
                                 False, _
                                 IsPeriodOutsideRange(CellDate)

    Next CellIndex

End Sub


Private Function IsCurrentYear(ByVal DisplayYear As Long) As Boolean

    IsCurrentYear = DisplayYear = Year(Date)

End Function


Private Function IsSelectedYear(ByVal DisplayYear As Long) As Boolean

    IsSelectedYear = DisplayYear = ParentPicker.GetCurrentYear

End Function


'----------------------------------------
' Year picker helpers
'----------------------------------------

Public Function IsPeriodOutsideRange(ByVal PeriodDate As Date) As Boolean

    IsPeriodOutsideRange = PeriodDate < DP_MinDate() Or _
                           PeriodDate > DP_MaxDate()

End Function


Private Function GetFirstDisplayedYear(ByVal SelectedYear As Long) As Long

    Dim MinYear As Long
    Dim Offset As Long

    MinYear = Year(DP_MinDate())
    Offset = (SelectedYear - MinYear) \ YEAR_COUNT

    GetFirstDisplayedYear = MinYear + Offset * YEAR_COUNT

End Function


Public Sub YearLabelClicked(ByVal YearNumber As Long)

    ParentPicker.SetYear YearNumber

    Me.Hide

End Sub


Public Sub GoToCurrentYear()

    CurrentYear = Year(Date)
    ParentPicker.SetYear CurrentYear

    Me.Hide

End Sub


'----------------------------------------
' Arrow
'----------------------------------------

Public Sub ArrowClicked(ByVal Action As String)

    If Not IsArrowEnabled(Action) Then Exit Sub

    Select Case Action
        Case "PREV_RANGE_YEAR"
            CurrentYear = CurrentYear - YEAR_COUNT
        Case "NEXT_RANGE_YEAR"
            CurrentYear = CurrentYear + YEAR_COUNT
    End Select

    BuildYears

End Sub


Public Function IsArrowEnabled(ByVal Action As String) As Boolean

    Dim FirstYear As Long
    Dim LastYear As Long

    FirstYear = GetFirstDisplayedYear(CurrentYear)
    LastYear = FirstYear + YEAR_COUNT - 1

    Select Case Action
        Case "PREV_RANGE_YEAR"
            IsArrowEnabled = FirstYear > Year(DP_MinDate())
        Case "NEXT_RANGE_YEAR"
            IsArrowEnabled = LastYear < Year(DP_MaxDate())
        Case Else
            IsArrowEnabled = False
    End Select

End Function


Private Sub UpdateArrowStates()

    ArrowHandlers.Item(1).SetState IsArrowEnabled("PREV_RANGE_YEAR")
    ArrowHandlers.Item(2).SetState IsArrowEnabled("NEXT_RANGE_YEAR")

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub ResetArrowHover()

    DP_ResetHoverCollection ArrowHandlers

End Sub


Public Sub ResetYearHover()

    DP_ResetHoverCollection YearLabelHandlers

End Sub


Public Sub ResetCurrentYearHover()

    If Not CurrentYearButtonHandler Is Nothing Then
        CurrentYearButtonHandler.ResetHover
    End If

End Sub


Public Sub ResetAllHover()

    ResetArrowHover
    ResetYearHover
    ResetCurrentYearHover

End Sub


Public Sub ResetYearHoverExcept(ByVal CurrentYearLabel As DP_CPeriodLabel)

    DP_ResetHoverCollectionExcept YearLabelHandlers, CurrentYearLabel

End Sub


'----------------------------------------
' Events
'----------------------------------------

Private Sub UserForm_MouseMove(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Single, _
                               ByVal Y As Single)

    ResetAllHover

End Sub

