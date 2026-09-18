VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CViewCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements DP_IPickerView

'----------------------------------------
' Declarations
'----------------------------------------

Private PickerForm As DP_frmDatePicker

Private MonthHeaderHandler As DP_CHeaderPeriod
Private TodayButtonHandler As DP_CButtonShortcut

Private ArrowHandlers As Collection
Private WeekdayHandlers As Collection
Private DayHandlers As Collection

Private HoveredButton As Object

Private Const CALENDAR_DAY_COUNT As Long = _
    DP_CALENDAR_COLUMNS * DP_CALENDAR_ROWS


'----------------------------------------
' Initialization
'----------------------------------------

Private Sub DP_IPickerView_InitializeView(ByVal Form As DP_frmDatePicker)

    Dim ArrowHandler As DP_CHeaderArrow

    Set PickerForm = Form
    Set ArrowHandlers = New Collection
    Set WeekdayHandlers = New Collection
    Set DayHandlers = New Collection

    ' Header
    '--------------------

    Set MonthHeaderHandler = New DP_CHeaderPeriod
    MonthHeaderHandler.Initialize Me, Format(CurrentMonth, "mmmm yyyy")

    ' Previous month
    '--------------------

    Set ArrowHandler = New DP_CHeaderArrow
    ArrowHandler.Initialize Me, ChrW(&H25B2), _
                            DP_ARROW_PREV_LEFT, DP_NAV_PREVIOUS

    ArrowHandlers.Add ArrowHandler

    ' Next month
    '--------------------

    Set ArrowHandler = New DP_CHeaderArrow
    ArrowHandler.Initialize Me, ChrW(&H25BC), _
                            DP_ARROW_NEXT_LEFT, DP_NAV_NEXT

    ArrowHandlers.Add ArrowHandler

    ' Calendar
    '--------------------

    CreateWeekdays
    CreateDays

    ' Shortcut
    '--------------------

    Set TodayButtonHandler = New DP_CButtonShortcut
    TodayButtonHandler.Initialize _
        Me, _
        DP_GetDatePickerActionText("GO_TO_TODAY"), _
        DP_GetActionButtonTop(DP_SIZE_DATE)


End Sub


'----------------------------------------
' Grid
'----------------------------------------

Private Sub CreateWeekdays()

    Dim Handler As DP_CHeaderWeekday
    Dim Index As Long

    For Index = 0 To 6

        Set Handler = New DP_CHeaderWeekday
        Handler.Initialize Me, _
                           WeekdayName(Index + 1, True, DP_FIRST_DAY_OF_WEEK), _
                           DP_GridLeft(Index, DP_GRID_LEFT, _
                                       DP_GRID_CELL_WIDTH, DP_CALENDAR_CELL_WIDTH), _
                           DP_GridTop(0, DP_GRID_TOP, _
                                      DP_GRID_CELL_HEIGHT, DP_CALENDAR_CELL_HEIGHT)

        WeekdayHandlers.Add Handler

    Next Index

End Sub


Private Sub CreateDays()

    Dim Handler As DP_CButtonDay
    Dim Index As Long
    Dim CalendarStartDate As Date

    Dim ColumnIndex As Long
    Dim RowIndex As Long

    CalendarStartDate = CurrentMonth - Weekday(CurrentMonth, DP_FIRST_DAY_OF_WEEK) + 1

    For Index = 0 To CALENDAR_DAY_COUNT - 1

        ColumnIndex = DP_GridColumn(Index, DP_CALENDAR_COLUMNS)
        RowIndex = DP_GridRow(Index, DP_CALENDAR_COLUMNS)

        Set Handler = New DP_CButtonDay

        Handler.Initialize Me, _
                           CalendarStartDate + Index, _
                           CStr(Day(CalendarStartDate + Index)), _
                           DP_GridLeft(ColumnIndex, DP_GRID_LEFT, _
                                       DP_GRID_CELL_WIDTH, DP_CALENDAR_CELL_WIDTH), _
                           DP_GridTop(RowIndex + 1, DP_GRID_TOP, _
                                      DP_GRID_CELL_HEIGHT, DP_CALENDAR_CELL_HEIGHT)

        DayHandlers.Add Handler

    Next Index

End Sub


'----------------------------------------
' Update
'----------------------------------------

Private Sub DP_IPickerView_UpdateView()

    Dim Index As Long
    Dim CalendarStartDate As Date
    Dim Handler As DP_CButtonDay

    DP_IPickerView_ClearHoveredButton

    MonthHeaderHandler.SetCaption Format(CurrentMonth, "mmmm yyyy")

    CalendarStartDate = CurrentMonth - Weekday(CurrentMonth, DP_FIRST_DAY_OF_WEEK) + 1

    For Index = 0 To CALENDAR_DAY_COUNT - 1
        Set Handler = DayHandlers.Item(Index + 1)
        Handler.UpdateDate CalendarStartDate + Index, _
                           CStr(Day(CalendarStartDate + Index))
    Next Index

    DP_UpdateStateCollection ArrowHandlers

End Sub


Private Sub DP_IPickerView_RefreshTheme()

    ' Day handlers are handled by UpdateView.
    ' Shortcut button is handled by SetVisible.

    MonthHeaderHandler.RefreshTheme
    DP_RefreshThemeCollection ArrowHandlers
    DP_RefreshThemeCollection WeekdayHandlers

End Sub


'----------------------------------------
' Properties
'----------------------------------------

Private Property Get DP_IPickerView_ParentForm() As DP_frmDatePicker

    Set DP_IPickerView_ParentForm = PickerForm

End Property


Private Property Get CurrentMonth() As Date

    CurrentMonth = PickerForm.CurrentMonth

End Property


'----------------------------------------
' Queries
'----------------------------------------

Private Function DP_IPickerView_IsArrowEnabled(ByVal Direction As DP_NavigationDirection) As Boolean

    Dim TargetMonth As Date

    TargetMonth = DateAdd("m", Direction, CurrentMonth)

    If Direction = DP_NAV_PREVIOUS Then
        DP_IPickerView_IsArrowEnabled = TargetMonth >= DateSerial(Year(DP_MinDate()), Month(DP_MinDate()), 1)
    Else
        DP_IPickerView_IsArrowEnabled = TargetMonth <= DateSerial(Year(DP_MaxDate()), Month(DP_MaxDate()), 1)
    End If

End Function


Private Function DP_IPickerView_IsDateSelected(ByVal CellDate As Date) As Boolean

    DP_IPickerView_IsDateSelected = CellDate = DateValue(PickerForm.InitialCellDate)

End Function


Private Function DP_IPickerView_IsCurrentPeriod(ByVal CellDate As Date) As Boolean

    DP_IPickerView_IsCurrentPeriod = CellDate = DateValue(Date)

End Function


Private Function DP_IPickerView_IsSecondaryPeriod(ByVal CellDate As Date) As Boolean
    
    DP_IPickerView_IsSecondaryPeriod = Month(CellDate) <> Month(CurrentMonth)

End Function


'----------------------------------------
' Click
'----------------------------------------

Private Sub DP_IPickerView_HeaderClicked()

    PickerForm.HeaderClicked DP_SIZE_MONTH

End Sub


Private Sub DP_IPickerView_ArrowClicked(ByVal Direction As DP_NavigationDirection)

    If Not DP_IPickerView_IsArrowEnabled(Direction) Then Exit Sub

    PickerForm.ChangePeriod Direction, "m"

End Sub


Private Sub DP_IPickerView_GridButtonClicked(ByVal SelectedDate As Date)

    If DP_IsDateOutsideRange(SelectedDate) Then Exit Sub

    PickerForm.DaySelected SelectedDate

End Sub


Private Sub DP_IPickerView_ActionButtonClicked()

    PickerForm.DaySelected Date

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Private Sub DP_IPickerView_SetHoveredButton(ByVal Button As Object)

    If HoveredButton Is Button Then Exit Sub

    If Not HoveredButton Is Nothing Then
        HoveredButton.SetHovered False
    End If

    Set HoveredButton = Button
    HoveredButton.SetHovered True

End Sub


Private Sub DP_IPickerView_ClearHoveredButton()

    If HoveredButton Is Nothing Then Exit Sub

    HoveredButton.SetHovered False
    Set HoveredButton = Nothing

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Private Sub DP_IPickerView_SetVisible(ByVal IsVisible As Boolean)

    MonthHeaderHandler.SetVisible IsVisible
    TodayButtonHandler.SetVisible IsVisible

    DP_SetVisibleCollection ArrowHandlers, IsVisible
    DP_SetVisibleCollection WeekdayHandlers, IsVisible
    DP_SetVisibleCollection DayHandlers, IsVisible

    If IsVisible Then
        TodayButtonHandler.Recreate _
            DP_GetDatePickerActionText("GO_TO_TODAY"), _
            DP_GetActionButtonTop(DP_SIZE_DATE)
    End If

End Sub

