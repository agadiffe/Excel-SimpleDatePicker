VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CViewMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Private PickerForm As DP_frmDatePicker

Private YearHeaderHandler As DP_CHeaderPeriod
Private CurrentMonthButtonHandler As DP_CButtonShortcut

Private ArrowHandlers As Collection
Private MonthHandlers As Collection

Private Const MONTH_COUNT As Long = _
    DP_MONTH_GRID_COLUMNS * DP_MONTH_GRID_ROWS


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
    Set MonthHandlers = New Collection

    ' Header
    '--------------------

    Set YearHeaderHandler = New DP_CHeaderPeriod
    YearHeaderHandler.Initialize Me, CStr(CurrentYear)

    ' Previous year
    '--------------------

    Set ArrowHandler = New DP_CHeaderArrow
    ArrowHandler.Initialize Me, ChrW(&H25B2), _
                            DP_ARROW_PREV_LEFT, DP_NAV_PREVIOUS

    ArrowHandlers.Add ArrowHandler

    ' Next year
    '--------------------

    Set ArrowHandler = New DP_CHeaderArrow
    ArrowHandler.Initialize Me, ChrW(&H25BC), _
                            DP_ARROW_NEXT_LEFT, DP_NAV_NEXT

    ArrowHandlers.Add ArrowHandler

    ' Months
    '--------------------

    CreateMonths

    ' Shortcut
    '--------------------

    Set CurrentMonthButtonHandler = New DP_CButtonShortcut
    CurrentMonthButtonHandler.Initialize _
        Me, DP_GetDatePickerActionText("GO_TO_CURRENT_MONTH")

End Sub


'----------------------------------------
' Grid
'----------------------------------------

Private Sub CreateMonths()

    Dim Handler As DP_CButtonPeriod
    Dim Index As Long
    Dim MonthNumber As Long
    Dim CellDate As Date

    Dim ColumnIndex As Long
    Dim RowIndex As Long

    For Index = 0 To MONTH_COUNT - 1

        MonthNumber = Index + 1
        CellDate = GetMonthDate(Index)
        ColumnIndex = DP_GridColumn(Index, DP_MONTH_GRID_COLUMNS)
        RowIndex = DP_GridRow(Index, DP_MONTH_GRID_COLUMNS)

        Set Handler = New DP_CButtonPeriod

        Handler.Initialize Me, _
                           GetMonthCaption(MonthNumber), _
                           MonthNumber, _
                           CellDate, _
                           DP_GridLeft(ColumnIndex, DP_PERIOD_GRID_LEFT, _
                                       DP_MONTH_GRID_CELL_WIDTH, DP_MONTH_CELL_WIDTH), _
                           DP_GridTop(RowIndex, DP_PERIOD_GRID_TOP, _
                                      DP_PERIOD_GRID_CELL_HEIGHT, DP_PERIOD_CELL_HEIGHT), _
                           DP_MONTH_CELL_WIDTH

        MonthHandlers.Add Handler

    Next Index

End Sub


'----------------------------------------
' Update
'----------------------------------------

Public Sub UpdateView()

    Dim Index As Long
    Dim MonthNumber As Long
    Dim CellDate As Date

    Dim Handler As DP_CButtonPeriod

    YearHeaderHandler.SetCaption CStr(CurrentYear)

    For Index = 0 To MONTH_COUNT - 1

        MonthNumber = Index + 1
        CellDate = GetMonthDate(Index)

        Set Handler = MonthHandlers.Item(Index + 1)
        Handler.UpdatePeriod GetMonthCaption(MonthNumber), _
                             CellDate, _
                             MonthNumber

    Next Index

    DP_UpdateStateCollection ArrowHandlers
    CurrentMonthButtonHandler.Recreate _
        DP_GetDatePickerActionText("GO_TO_CURRENT_MONTH")

End Sub


Public Sub RefreshTheme()

    ' Month handlers and shortcut button are handled by UpdateView.

    YearHeaderHandler.RefreshTheme
    DP_RefreshThemeCollection ArrowHandlers

End Sub


'----------------------------------------
' Properties
'----------------------------------------

Public Property Get ParentForm() As DP_frmDatePicker

    Set ParentForm = PickerForm

End Property


Private Property Get CurrentYear() As Long

    CurrentYear = Year(PickerForm.CurrentMonth)

End Property


'----------------------------------------
' Queries
'----------------------------------------

Public Function IsArrowEnabled(ByVal Direction As DP_NavigationDirection) As Boolean

    Dim TargetYear As Long

    TargetYear = CurrentYear + Direction

    If Direction = DP_NAV_PREVIOUS Then
        IsArrowEnabled = TargetYear >= Year(DP_MinDate())
    Else
        IsArrowEnabled = TargetYear <= Year(DP_MaxDate())
    End If

End Function


Public Function IsPeriodOutsideRange(ByVal PeriodDate As Date) As Boolean

    IsPeriodOutsideRange = PickerForm.IsPeriodOutsideRange(PeriodDate)

End Function


Public Function IsPeriodSelected(ByVal CellDate As Date) As Boolean

    IsPeriodSelected = CellDate = PickerForm.SelectedMonth

End Function


Public Function IsCurrentPeriod(ByVal MonthNumber As Long) As Boolean

    IsCurrentPeriod = MonthNumber = Month(Date) And CurrentYear = Year(Date)

End Function


Public Function IsSecondaryPeriod(ByVal PeriodDate As Date) As Boolean

    IsSecondaryPeriod = Year(PeriodDate) <> CurrentYear

End Function


'----------------------------------------
' Month
'----------------------------------------

Private Function GetMonthCaption(ByVal MonthNumber As Long) As String

    If DP_MONTH_GRID_COLUMNS > 3 Then
        GetMonthCaption = Format(DateSerial(CurrentYear, MonthNumber, 1), "mmm")
    Else
        GetMonthCaption = Format(DateSerial(CurrentYear, MonthNumber, 1), "mmmm")
    End If

End Function


Private Function GetMonthDate(ByVal CellIndex As Long) As Date

    GetMonthDate = DateSerial(CurrentYear, CellIndex + 1, 1)

End Function


'----------------------------------------
' Click
'----------------------------------------

Public Sub HeaderClicked()

    PickerForm.MonthHeaderClicked

End Sub


Public Sub ArrowClicked(ByVal Direction As DP_NavigationDirection)

    If Not IsArrowEnabled(Direction) Then Exit Sub

    PickerForm.ChangeYear Direction

End Sub


Public Sub PeriodClicked(ByVal MonthNumber As Long)

    If IsPeriodOutsideRange(DateSerial(CurrentYear, MonthNumber, 1)) Then Exit Sub

    PickerForm.MonthSelected MonthNumber, CurrentYear

End Sub


Public Sub ActionButtonClicked()

    PickerForm.MonthSelected Month(Date), Year(Date)

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub ResetArrowHover()

    DP_ResetHoverCollection ArrowHandlers

End Sub


Public Sub ResetMonthHover()

    DP_ResetHoverCollection MonthHandlers

End Sub


Public Sub ResetPeriodHoverExcept(ByVal CurrentMonthLabel As DP_CButtonPeriod)

    DP_ResetHoverCollectionExcept MonthHandlers, CurrentMonthLabel
    ResetHeaderHover

End Sub


Private Sub ResetHeaderHover()

    If Not YearHeaderHandler Is Nothing Then
        YearHeaderHandler.ResetHover
    End If

End Sub


Public Sub ResetActionButtonHover()

    If Not CurrentMonthButtonHandler Is Nothing Then
        CurrentMonthButtonHandler.ResetHover
    End If

End Sub


Public Sub ResetAllHover()

    ResetArrowHover
    ResetMonthHover
    ResetHeaderHover
    ResetActionButtonHover

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    Dim Handler As Object

    If Not YearHeaderHandler Is Nothing Then
        YearHeaderHandler.SetVisible IsVisible
    End If

    For Each Handler In ArrowHandlers
        Handler.SetVisible IsVisible
    Next Handler

    For Each Handler In MonthHandlers
        Handler.SetVisible IsVisible
    Next Handler

    If Not CurrentMonthButtonHandler Is Nothing Then
        CurrentMonthButtonHandler.SetVisible IsVisible
    End If

End Sub

