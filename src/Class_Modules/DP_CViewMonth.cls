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

Implements DP_IPickerView

'----------------------------------------
' Declarations
'----------------------------------------

Private PickerForm As DP_frmDatePicker

Private YearHeaderHandler As DP_CHeaderPeriod
Private CurrentMonthButtonHandler As DP_CButtonShortcut

Private ArrowHandlers As Collection
Private MonthHandlers As Collection

Private HoveredButton As Object

Private Const MONTH_COUNT As Long = _
    DP_MONTH_GRID_COLUMNS * DP_MONTH_GRID_ROWS


'----------------------------------------
' Initialization
'----------------------------------------

Private Sub DP_IPickerView_InitializeView(ByVal Form As DP_frmDatePicker)

    Dim ArrowHandler As DP_CHeaderArrow

    Set PickerForm = Form
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
        Me, _
        DP_GetDatePickerActionText("GO_TO_CURRENT_MONTH"), _
        DP_GetActionButtonTop(DP_SIZE_MONTH)

End Sub


'----------------------------------------
' Grid
'----------------------------------------

Private Sub CreateMonths()

    Dim Handler As DP_CButtonPeriod
    Dim Index As Long

    Dim ColumnIndex As Long
    Dim RowIndex As Long

    For Index = 0 To MONTH_COUNT - 1

        ColumnIndex = DP_GridColumn(Index, DP_MONTH_GRID_COLUMNS)
        RowIndex = DP_GridRow(Index, DP_MONTH_GRID_COLUMNS)

        Set Handler = New DP_CButtonPeriod

        Handler.Initialize Me, _
                           DP_BUTTON_TYPE_MONTH, _
                           DateSerial(CurrentYear, Index + 1, 1), _
                           MonthName((Index Mod 12) + 1, DP_MONTH_GRID_COLUMNS > 3), _
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

Private Sub DP_IPickerView_UpdateView()

    Dim Index As Long
    Dim Handler As DP_CButtonPeriod

    DP_IPickerView_ClearHoveredButton

    YearHeaderHandler.SetCaption CStr(CurrentYear)

    For Index = 0 To MONTH_COUNT - 1
        Set Handler = MonthHandlers.Item(Index + 1)
        Handler.UpdatePeriod DateSerial(CurrentYear, Index + 1, 1), _
                             MonthName((Index Mod 12) + 1, DP_MONTH_GRID_COLUMNS > 3)
    Next Index

    DP_UpdateStateCollection ArrowHandlers

End Sub


Private Sub DP_IPickerView_RefreshTheme()

    DP_ClearButtonPictureCache

    ' Month handlers are handled by UpdateView.
    ' Shortcut button is handled by SetVisible.

    YearHeaderHandler.RefreshTheme
    DP_RefreshThemeCollection ArrowHandlers

End Sub


'----------------------------------------
' Properties
'----------------------------------------

Private Property Get DP_IPickerView_ParentForm() As DP_frmDatePicker

    Set DP_IPickerView_ParentForm = PickerForm

End Property


Private Property Get CurrentYear() As Long

    CurrentYear = Year(PickerForm.CurrentMonth)

End Property


'----------------------------------------
' Queries
'----------------------------------------

Private Function DP_IPickerView_IsArrowEnabled(ByVal Direction As DP_NavigationDirection) As Boolean

    Dim TargetYear As Long

    TargetYear = CurrentYear + Direction

    If Direction = DP_NAV_PREVIOUS Then
        DP_IPickerView_IsArrowEnabled = TargetYear >= Year(DP_MinDate())
    Else
        DP_IPickerView_IsArrowEnabled = TargetYear <= Year(DP_MaxDate())
    End If

End Function


Private Function DP_IPickerView_IsDateSelected(ByVal CellDate As Date) As Boolean

    DP_IPickerView_IsDateSelected = CellDate = PickerForm.SelectedMonth

End Function


Private Function DP_IPickerView_IsCurrentPeriod(ByVal CellDate As Date) As Boolean

    DP_IPickerView_IsCurrentPeriod = Month(CellDate) = Month(Date) And CurrentYear = Year(Date)

End Function


Private Function DP_IPickerView_IsSecondaryPeriod(ByVal CellDate As Date) As Boolean

    DP_IPickerView_IsSecondaryPeriod = Year(CellDate) <> CurrentYear

End Function


'----------------------------------------
' Click
'----------------------------------------

Private Sub DP_IPickerView_HeaderClicked()

    PickerForm.HeaderClicked DP_SIZE_YEAR

End Sub


Private Sub DP_IPickerView_ArrowClicked(ByVal Direction As DP_NavigationDirection)

    If Not DP_IPickerView_IsArrowEnabled(Direction) Then Exit Sub

    PickerForm.ChangePeriod Direction, "yyyy"

End Sub


Private Sub DP_IPickerView_GridButtonClicked(ByVal SelectedDate As Date)

    If DP_IsDateOutsideRange(SelectedDate) Then Exit Sub

    PickerForm.MonthSelected Month(SelectedDate), Year(SelectedDate)

End Sub


Private Sub DP_IPickerView_ActionButtonClicked()

    PickerForm.MonthSelected Month(Date), Year(Date)

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

    YearHeaderHandler.SetVisible IsVisible
    CurrentMonthButtonHandler.SetVisible IsVisible

    DP_SetVisibleCollection ArrowHandlers, IsVisible
    DP_SetVisibleCollection MonthHandlers, IsVisible

    If IsVisible Then
        CurrentMonthButtonHandler.Recreate _
            DP_GetDatePickerActionText("GO_TO_CURRENT_MONTH"), _
            DP_GetActionButtonTop(DP_SIZE_MONTH)
    End If

End Sub

