VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DP_frmDatePicker 
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4365
   OleObjectBlob   =   "DP_frmDatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DP_frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' References
'----------------------------------------

Private Const CALENDAR_DAY_COUNT As Long = _
    DP_CALENDAR_COLUMNS * DP_CALENDAR_ROWS

Private TargetCell As Range
Private CurrentMonth As Date

Private CalendarInitialized As Boolean

Private MonthLabelHandler As DP_CHeaderLabel
Private ArrowHandlers As Collection
Private WeekdayHandlers As Collection
Private DayTextHandlers As Collection
Private TodayButtonHandler As DP_CActionButton


'----------------------------------------
' Date picker
'----------------------------------------

Public Sub ShowPicker(ByVal Cell As Range)

    Dim InitialDate As Date

    Set TargetCell = Cell

    If IsDate(Cell.Value) Then
        InitialDate = CDate(Cell.Value)
    Else
        InitialDate = Date
    End If

    CurrentMonth = DateSerial(Year(InitialDate), Month(InitialDate), 1)

    BuildCalendar
    DP_ShowPopupNextToCell Me, Cell

End Sub


Private Sub BuildCalendar()

    Dim CurrentDarkMode As Boolean

    CurrentDarkMode = DP_IsDarkMode()

    If Not CalendarInitialized Then
        DP_DarkMode = CurrentDarkMode
        InitializeCalendar
        CalendarInitialized = True
    ElseIf DP_DarkMode <> CurrentDarkMode Then
        DP_DarkMode = CurrentDarkMode
        InitializeCalendar
    Else
        RefreshCalendar
    End If

End Sub


Private Sub InitializeCalendar()

    Dim ArrowHandler As DP_CArrowLabel

    Set ArrowHandlers = New Collection
    Set WeekdayHandlers = New Collection
    Set DayTextHandlers = New Collection

    DP_ClearPickerControls Me
    DP_InitializePicker Me, DP_DATEPICKER_WIDTH, DP_DATEPICKER_HEIGHT

    ' Month / Year
    '--------------------

    Set MonthLabelHandler = New DP_CHeaderLabel
    MonthLabelHandler.Setup Me, Format(CurrentMonth, "mmmm yyyy")

    ' Previous month
    '--------------------

    Set ArrowHandler = New DP_CArrowLabel
    ArrowHandler.Setup Me, "PREV_MONTH", ChrW(&H25B2), DP_ARROW_PREV_LEFT

    ArrowHandlers.Add ArrowHandler

    ' Next month
    '--------------------

    Set ArrowHandler = New DP_CArrowLabel
    ArrowHandler.Setup Me, "NEXT_MONTH", ChrW(&H25BC), DP_ARROW_NEXT_LEFT

    ArrowHandlers.Add ArrowHandler

    ' Calendar dates
    '--------------------

    BuildWeekdaysHeader
    CreateCalendarDays

    ' Go To Today button
    '--------------------

    Set TodayButtonHandler = New DP_CActionButton
    TodayButtonHandler.Setup Me, "GO_TO_TODAY"

End Sub


Private Sub RefreshCalendar()

    ResetAllHover

    MonthLabelHandler.SetCaption Format(CurrentMonth, "mmmm yyyy")
    UpdateCalendarDays

End Sub


'----------------------------------------
' Build Calendar helpers
'----------------------------------------

Private Sub BuildWeekdaysHeader()

    Dim WeekdayHandler As DP_CWeekdayLabel
    Dim WeekdayIndex As Long

    For WeekdayIndex = 0 To 6

        Set WeekdayHandler = New DP_CWeekdayLabel
        WeekdayHandler.Setup Me, WeekdayIndex

        WeekdayHandlers.Add WeekdayHandler

    Next WeekdayIndex

End Sub


Private Sub CreateCalendarDays()

    Dim CellIndex As Long

    For CellIndex = 0 To CALENDAR_DAY_COUNT - 1
        CreateCalendarDay GetCalendarCellDate(CellIndex), CellIndex
    Next CellIndex

End Sub


Private Sub CreateCalendarDay(ByVal CellDate As Date, _
                              ByVal CellIndex As Long)

    Dim DayTextHandler As DP_CDayLabel

    Dim ColumnIndex As Long
    Dim RowIndex As Long

    ColumnIndex = DP_GridColumn(CellIndex, DP_CALENDAR_COLUMNS)
    RowIndex = DP_GridRow(CellIndex, DP_CALENDAR_COLUMNS)

    Set DayTextHandler = New DP_CDayLabel

    DayTextHandler.Setup Me, _
                         CellDate, _
                         DP_GridLeft(ColumnIndex, _
                                     DP_GRID_LEFT, _
                                     DP_GRID_CELL_WIDTH, _
                                     DP_CALENDAR_CELL_WIDTH), _
                         DP_GridTop(RowIndex + 1, _
                                    DP_GRID_TOP, _
                                    DP_GRID_CELL_HEIGHT)

    DayTextHandlers.Add DayTextHandler

End Sub


Private Sub UpdateCalendarDays()

    Dim CellIndex As Long
    Dim DayTextHandler As DP_CDayLabel

    For CellIndex = 0 To CALENDAR_DAY_COUNT - 1
        Set DayTextHandler = DayTextHandlers.Item(CellIndex + 1)
        DayTextHandler.UpdateDate GetCalendarCellDate(CellIndex)
    Next CellIndex

End Sub


Private Function GetCalendarCellDate(ByVal CellIndex As Long) As Date

    Dim FirstDayIndex As Long

    FirstDayIndex = Weekday(CurrentMonth, vbMonday)

    GetCalendarCellDate = CurrentMonth - FirstDayIndex + 1 + CellIndex

End Function


'----------------------------------------
' Date picker helpers
'----------------------------------------

Public Function GetTargetCell() As Range

    Set GetTargetCell = TargetCell

End Function


Public Function GetCurrentMonth() As Date

    GetCurrentMonth = CurrentMonth

End Function


Public Function IsSelected(ByVal CellDate As Date) As Boolean

    If Not IsDate(TargetCell.Value) Then Exit Function

    IsSelected = (CellDate = DateValue(TargetCell.Value))

End Function


Public Sub SetMonthYear(ByVal MonthNumber As Long, _
                        ByVal YearNumber As Long)

    CurrentMonth = DateSerial(YearNumber, MonthNumber, 1)

    BuildCalendar

End Sub


Public Sub HeaderClicked()

    Me.Hide
    DP_frmMonthPicker.ShowMonths Me
    Me.Show

End Sub


Public Sub ArrowClicked(ByVal Action As String)

    Select Case Action
        Case "PREV_MONTH"
            CurrentMonth = DateAdd("m", -1, CurrentMonth)
        Case "NEXT_MONTH"
            CurrentMonth = DateAdd("m", 1, CurrentMonth)
    End Select

    BuildCalendar

End Sub


Public Sub DayLabelClicked(ByVal SelectedDate As Date)

    Dim ExistingTime As Double

    If IsDate(TargetCell.Value) Then
        ExistingTime = TimeValue(TargetCell.Value)
        TargetCell.Value = DateValue(SelectedDate) + ExistingTime
    Else
        TargetCell.Value = SelectedDate
    End If

    Me.Hide

End Sub


Public Sub GoToToday()

    Dim ExistingTime As Double

    If IsDate(TargetCell.Value) Then
        ExistingTime = TimeValue(TargetCell.Value)
        TargetCell.Value = Date + ExistingTime
    Else
        TargetCell.Value = Date
    End If

    Me.Hide

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub ResetArrowHover()

    DP_ResetHoverCollection ArrowHandlers

End Sub


Public Sub ResetDayTextHover()

    DP_ResetHoverCollection DayTextHandlers

End Sub


Public Sub ResetHeaderHover()

    If Not MonthLabelHandler Is Nothing Then
        MonthLabelHandler.ResetHover
    End If

End Sub


Public Sub ResetTodayHover()

    If Not TodayButtonHandler Is Nothing Then
        TodayButtonHandler.ResetHover
    End If

End Sub


Public Sub ResetAllHover()

    ResetArrowHover
    ResetHeaderHover
    ResetDayTextHover
    ResetTodayHover

End Sub


Public Sub ResetDayHoverExcept(ByVal CurrentDay As DP_CDayLabel)

    DP_ResetHoverCollectionExcept DayTextHandlers, CurrentDay

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

