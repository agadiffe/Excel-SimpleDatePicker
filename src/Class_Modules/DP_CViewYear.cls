VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CViewYear"
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

Private YearRangeHeaderHandler As DP_CHeaderPeriod
Private CurrentYearButtonHandler As DP_CButtonShortcut

Private ArrowHandlers As Collection
Private YearHandlers As Collection

Private HoveredButton As Object

Private Const YEAR_COUNT As Long = _
    DP_YEAR_GRID_COLUMNS * DP_YEAR_GRID_ROWS


'----------------------------------------
' Initialization
'----------------------------------------

Private Sub DP_IPickerView_InitializeView(ByVal Form As DP_frmDatePicker)

    Dim ArrowHandler As DP_CHeaderArrow

    Set PickerForm = Form
    Set ArrowHandlers = New Collection
    Set YearHandlers = New Collection

    ' Header
    '--------------------

    Set YearRangeHeaderHandler = New DP_CHeaderPeriod
    YearRangeHeaderHandler.Initialize Me, _
        CStr(FirstDisplayedYear) & " - " & CStr(LastDisplayedYear)
    YearRangeHeaderHandler.SetInteractive False

    ' Previous range
    '--------------------

    Set ArrowHandler = New DP_CHeaderArrow
    ArrowHandler.Initialize Me, ChrW(&H25B2), _
                            DP_ARROW_PREV_LEFT, DP_NAV_PREVIOUS

    ArrowHandlers.Add ArrowHandler

    ' Next range
    '--------------------

    Set ArrowHandler = New DP_CHeaderArrow
    ArrowHandler.Initialize Me, ChrW(&H25BC), _
                            DP_ARROW_NEXT_LEFT, DP_NAV_NEXT

    ArrowHandlers.Add ArrowHandler

    ' Years
    '--------------------

    CreateYears

    ' Shortcut
    '--------------------

    Set CurrentYearButtonHandler = New DP_CButtonShortcut
    CurrentYearButtonHandler.Initialize _
        Me, _
        DP_GetDatePickerActionText("GO_TO_CURRENT_YEAR"), _
        DP_GetActionButtonTop(DP_SIZE_YEAR)

End Sub


'----------------------------------------
' Grid
'----------------------------------------

Private Sub CreateYears()

    Dim Handler As DP_CButtonPeriod
    Dim Index As Long
    Dim DisplayedYear As Long
    Dim FirstYear As Long

    Dim ColumnIndex As Long
    Dim RowIndex As Long

    FirstYear = FirstDisplayedYear

    For Index = 0 To YEAR_COUNT - 1

        DisplayedYear = FirstYear + Index

        ColumnIndex = DP_GridColumn(Index, DP_YEAR_GRID_COLUMNS)
        RowIndex = DP_GridRow(Index, DP_YEAR_GRID_COLUMNS)

        Set Handler = New DP_CButtonPeriod

        Handler.Initialize Me, _
                           CStr(DisplayedYear), _
                           DateSerial(DisplayedYear, 1, 1), _
                           DP_GridLeft(ColumnIndex, DP_PERIOD_GRID_LEFT, _
                                       DP_YEAR_GRID_CELL_WIDTH, DP_YEAR_CELL_WIDTH), _
                           DP_GridTop(RowIndex, DP_PERIOD_GRID_TOP, _
                                      DP_PERIOD_GRID_CELL_HEIGHT, DP_PERIOD_CELL_HEIGHT), _
                           DP_YEAR_CELL_WIDTH

        YearHandlers.Add Handler

    Next Index

End Sub


'----------------------------------------
' Update
'----------------------------------------

Private Sub DP_IPickerView_UpdateView()

    Dim Index As Long
    Dim DisplayedYear As Long
    Dim FirstYear As Long

    Dim Handler As DP_CButtonPeriod

    DP_IPickerView_ClearHoveredButton

    FirstYear = FirstDisplayedYear

    YearRangeHeaderHandler.SetCaption _
        CStr(FirstDisplayedYear) & " - " & CStr(LastDisplayedYear)

    For Index = 0 To YEAR_COUNT - 1

        DisplayedYear = FirstYear + Index

        Set Handler = YearHandlers.Item(Index + 1)
        Handler.UpdatePeriod CStr(DisplayedYear), DateSerial(DisplayedYear, 1, 1)

    Next Index

    DP_UpdateStateCollection ArrowHandlers

End Sub


Private Sub DP_IPickerView_RefreshTheme()

    ' Year handlers are handled by UpdateView.
    ' Shortcut button is handled by SetVisible.

    YearRangeHeaderHandler.RefreshTheme
    DP_RefreshThemeCollection ArrowHandlers

End Sub


'----------------------------------------
' Properties
'----------------------------------------

Private Property Get DP_IPickerView_ParentForm() As DP_frmDatePicker

    Set DP_IPickerView_ParentForm = PickerForm

End Property


Private Property Get FirstDisplayedYear() As Long

    Dim MinYear As Long
    Dim Offset As Long

    MinYear = Year(DP_MinDate())
    Offset = (Year(PickerForm.CurrentMonth) - MinYear) \ YEAR_COUNT
    FirstDisplayedYear = MinYear + Offset * YEAR_COUNT

End Property


Private Property Get LastDisplayedYear() As Long

    LastDisplayedYear = FirstDisplayedYear + YEAR_COUNT - 1

End Property


'----------------------------------------
' Queries
'----------------------------------------

Private Function DP_IPickerView_IsArrowEnabled(ByVal Direction As DP_NavigationDirection) As Boolean

    If Direction = DP_NAV_PREVIOUS Then
        DP_IPickerView_IsArrowEnabled = FirstDisplayedYear > Year(DP_MinDate())
    Else
        DP_IPickerView_IsArrowEnabled = LastDisplayedYear < Year(DP_MaxDate())
    End If

End Function


Private Function DP_IPickerView_IsDateSelected(ByVal CellDate As Date) As Boolean

    DP_IPickerView_IsDateSelected = Year(CellDate) = PickerForm.SelectedYear

End Function


Private Function DP_IPickerView_IsCurrentPeriod(ByVal CellDate As Date) As Boolean

    DP_IPickerView_IsCurrentPeriod = Year(CellDate) = Year(Date)

End Function


Private Function DP_IPickerView_IsSecondaryPeriod(ByVal PeriodDate As Date) As Boolean

    DP_IPickerView_IsSecondaryPeriod = False

End Function


'----------------------------------------
' Click
'----------------------------------------

Private Sub DP_IPickerView_HeaderClicked()

    ' Latest view level. Header has no further navigation target.

End Sub


Private Sub DP_IPickerView_ArrowClicked(ByVal Direction As DP_NavigationDirection)

    If Not DP_IPickerView_IsArrowEnabled(Direction) Then Exit Sub

    PickerForm.ChangePeriod Direction * YEAR_COUNT, "yyyy"

End Sub


Private Sub DP_IPickerView_GridButtonClicked(ByVal SelectedDate As Date)

    If DP_IsDateOutsideRange(SelectedDate) Then Exit Sub

    PickerForm.YearSelected Year(SelectedDate)

End Sub


Private Sub DP_IPickerView_ActionButtonClicked()

    PickerForm.YearSelected Year(Date)

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

    YearRangeHeaderHandler.SetVisible IsVisible
    CurrentYearButtonHandler.SetVisible IsVisible

    DP_SetVisibleCollection ArrowHandlers, IsVisible
    DP_SetVisibleCollection YearHandlers, IsVisible

    If IsVisible Then
        CurrentYearButtonHandler.Recreate _
            DP_GetDatePickerActionText("GO_TO_CURRENT_YEAR"), _
            DP_GetActionButtonTop(DP_SIZE_YEAR)
    End If

End Sub

