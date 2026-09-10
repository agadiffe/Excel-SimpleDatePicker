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

'----------------------------------------
' Declarations
'----------------------------------------

Private PickerForm As DP_frmDatePicker

Private MonthHeaderHandler As DP_CHeaderPeriod
Private TodayButtonHandler As DP_CButtonShortcut

Private ArrowHandlers As Collection
Private WeekdayHandlers As Collection
Private DayHandlers As Collection

Private Const CALENDAR_DAY_COUNT As Long = _
    DP_CALENDAR_COLUMNS * DP_CALENDAR_ROWS


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub Attach(ByVal Form As DP_frmDatePicker)

    Set PickerForm = Form
    InitializeView

End Sub


Private Sub InitializeView()

    Dim ArrowHandler As DP_CHeaderArrow

    Set ArrowHandlers = New Collection
    Set WeekdayHandlers = New Collection
    Set DayHandlers = New Collection

    ' Header
    '--------------------

    Set MonthHeaderHandler = New DP_CHeaderPeriod
    MonthHeaderHandler.Initialize Me, Format(PickerForm.CurrentMonth, "mmmm yyyy")

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
        Me, DP_GetDatePickerActionText("GO_TO_TODAY")


End Sub


'----------------------------------------
' Grid
'----------------------------------------

Private Sub CreateWeekdays()

    Dim Handler As DP_CHeaderWeekday
    Dim Index As Long

    For Index = 0 To 6

        Set Handler = New DP_CHeaderWeekday
        Handler.Initialize Me, Index

        WeekdayHandlers.Add Handler

    Next Index

End Sub


Private Sub CreateDays()

    Dim Index As Long

    For Index = 0 To CALENDAR_DAY_COUNT - 1
        CreateDay Index
    Next Index

End Sub


Private Sub CreateDay(ByVal CellIndex As Long)

    Dim Handler As DP_CButtonDay

    Dim ColumnIndex As Long
    Dim RowIndex As Long

    ColumnIndex = DP_GridColumn(CellIndex, DP_CALENDAR_COLUMNS)
    RowIndex = DP_GridRow(CellIndex, DP_CALENDAR_COLUMNS)

    Set Handler = New DP_CButtonDay

    Handler.Initialize Me, _
                       GetCalendarCellDate(CellIndex), _
                       DP_GridLeft(ColumnIndex, DP_GRID_LEFT, _
                                   DP_GRID_CELL_WIDTH, DP_CALENDAR_CELL_WIDTH), _
                       DP_GridTop(RowIndex + 1, DP_GRID_TOP, _
                                  DP_GRID_CELL_HEIGHT, DP_CALENDAR_CELL_HEIGHT)

    DayHandlers.Add Handler

End Sub


Private Function GetCalendarCellDate(ByVal CellIndex As Long) As Date

    Dim FirstDayIndex As Long

    FirstDayIndex = Weekday(PickerForm.CurrentMonth, vbMonday)

    GetCalendarCellDate = PickerForm.CurrentMonth - FirstDayIndex + 1 + CellIndex

End Function


'----------------------------------------
' Update
'----------------------------------------

Public Sub UpdateView()

    Dim Index As Long
    Dim Handler As DP_CButtonDay

    MonthHeaderHandler.SetCaption Format(PickerForm.CurrentMonth, "mmmm yyyy")

    For Index = 0 To CALENDAR_DAY_COUNT - 1
        Set Handler = DayHandlers.Item(Index + 1)
        Handler.UpdateDate GetCalendarCellDate(Index)
    Next Index

    DP_UpdateStateCollection ArrowHandlers
    TodayButtonHandler.Recreate DP_GetDatePickerActionText("GO_TO_TODAY")

End Sub


Public Sub RefreshTheme()

    ' Day handlers and shortcut button are handled by UpdateView.

    MonthHeaderHandler.RefreshTheme
    DP_RefreshThemeCollection ArrowHandlers
    DP_RefreshThemeCollection WeekdayHandlers

End Sub


'----------------------------------------
' Properties
'----------------------------------------

Public Property Get ParentForm() As DP_frmDatePicker

    Set ParentForm = PickerForm

End Property


'----------------------------------------
' Queries
'----------------------------------------

Public Function IsArrowEnabled(ByVal Direction As DP_NavigationDirection) As Boolean

    Dim NewDate As Date

    NewDate = DateAdd("m", Direction, PickerForm.CurrentMonth)

    If Direction = DP_NAV_PREVIOUS Then
        IsArrowEnabled = NewDate >= DateSerial(Year(DP_MinDate()), Month(DP_MinDate()), 1)
    Else
        IsArrowEnabled = NewDate <= DateSerial(Year(DP_MaxDate()), Month(DP_MaxDate()), 1)
    End If

End Function


Public Function IsDayOutsideRange(ByVal DayDate As Date) As Boolean

    IsDayOutsideRange = PickerForm.IsDayOutsideRange(DayDate)

End Function


Public Function IsDaySelected(ByVal CellDate As Date) As Boolean

    IsDaySelected = PickerForm.IsDaySelected(CellDate)

End Function


'----------------------------------------
' Click
'----------------------------------------

Public Sub HeaderClicked()

    PickerForm.CalendarHeaderClicked

End Sub


Public Sub ArrowClicked(ByVal Direction As DP_NavigationDirection)

    If Not IsArrowEnabled(Direction) Then Exit Sub

    PickerForm.ChangeMonth Direction

End Sub


Public Sub DayClicked(ByVal SelectedDate As Date)

    If IsDayOutsideRange(SelectedDate) Then Exit Sub

    PickerForm.DaySelected SelectedDate

End Sub


Public Sub ActionButtonClicked()

    PickerForm.DaySelected Date

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub ResetArrowHover()

    DP_ResetHoverCollection ArrowHandlers

End Sub


Public Sub ResetDayHover()

    DP_ResetHoverCollection DayHandlers

End Sub


Public Sub ResetDayHoverExcept(ByVal CurrentDay As DP_CButtonDay)

    DP_ResetHoverCollectionExcept DayHandlers, CurrentDay

End Sub


Public Sub ResetHeaderHover()

    If Not MonthHeaderHandler Is Nothing Then
        MonthHeaderHandler.ResetHover
    End If

End Sub


Public Sub ResetTodayHover()

    If Not TodayButtonHandler Is Nothing Then
        TodayButtonHandler.ResetHover
    End If

End Sub


Public Sub ResetAllHover()

    ResetArrowHover
    ResetDayHover
    ResetHeaderHover
    ResetTodayHover

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    Dim Handler As Object

    If Not MonthHeaderHandler Is Nothing Then
        MonthHeaderHandler.SetVisible IsVisible
    End If

    For Each Handler In ArrowHandlers
        Handler.SetVisible IsVisible
    Next Handler

    For Each Handler In WeekdayHandlers
        Handler.SetVisible IsVisible
    Next Handler

    For Each Handler In DayHandlers
        Handler.SetVisible IsVisible
    Next Handler

    If Not TodayButtonHandler Is Nothing Then
        TodayButtonHandler.SetVisible IsVisible
    End If

End Sub

