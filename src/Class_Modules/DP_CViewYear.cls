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

'----------------------------------------
' Declarations
'----------------------------------------

Private PickerForm As DP_frmDatePicker

Private YearRangeHeaderHandler As DP_CHeaderPeriod
Private CurrentYearButtonHandler As DP_CButtonShortcut

Private ArrowHandlers As Collection
Private YearHandlers As Collection

Private Const YEAR_COUNT As Long = _
    DP_YEAR_GRID_COLUMNS * DP_YEAR_GRID_ROWS


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
    Set YearHandlers = New Collection

    ' Header
    '--------------------

    Set YearRangeHeaderHandler = New DP_CHeaderPeriod
    YearRangeHeaderHandler.Initialize Me, YearRangeCaption
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

    CreateYearLabels

    ' Shortcut
    '--------------------

    Set CurrentYearButtonHandler = New DP_CButtonShortcut
    CurrentYearButtonHandler.Initialize _
        Me, DP_GetDatePickerActionText("GO_TO_CURRENT_YEAR")

End Sub


'----------------------------------------
' Grid
'----------------------------------------

Private Sub CreateYearLabels()

    Dim Handler As DP_CButtonPeriod
    Dim Index As Long
    Dim DisplayYear As Long
    Dim FirstYear As Long
    Dim CellDate As Date

    Dim ColumnIndex As Long
    Dim RowIndex As Long

    FirstYear = FirstDisplayedYear

    For Index = 0 To YEAR_COUNT - 1

        DisplayYear = FirstYear + Index
        CellDate = DateSerial(DisplayYear, 1, 1)

        ColumnIndex = DP_GridColumn(Index, DP_YEAR_GRID_COLUMNS)
        RowIndex = DP_GridRow(Index, DP_YEAR_GRID_COLUMNS)

        Set Handler = New DP_CButtonPeriod

        Handler.Initialize Me, _
                           CStr(DisplayYear), _
                           DisplayYear, _
                           CellDate, _
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

Public Sub UpdateView()

    Dim Index As Long
    Dim DisplayYear As Long
    Dim FirstYear As Long
    Dim CellDate As Date

    Dim Handler As DP_CButtonPeriod

    FirstYear = FirstDisplayedYear

    YearRangeHeaderHandler.SetCaption YearRangeCaption

    For Index = 0 To YEAR_COUNT - 1

        DisplayYear = FirstYear + Index
        CellDate = DateSerial(DisplayYear, 1, 1)

        Set Handler = YearHandlers.Item(Index + 1)
        Handler.UpdatePeriod CStr(DisplayYear), _
                             CellDate, _
                             DisplayYear

    Next Index

    DP_UpdateStateCollection ArrowHandlers
    CurrentYearButtonHandler.Recreate _
        DP_GetDatePickerActionText("GO_TO_CURRENT_YEAR")

End Sub


Public Sub RefreshTheme()

    ' Year handlers and shortcut button are handled by UpdateView.

    YearRangeHeaderHandler.RefreshTheme
    DP_RefreshThemeCollection ArrowHandlers

End Sub


'----------------------------------------
' Properties
'----------------------------------------

Public Property Get ParentForm() As DP_frmDatePicker

    Set ParentForm = PickerForm

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


Private Property Get YearRangeCaption() As String

    YearRangeCaption = CStr(FirstDisplayedYear) & _
                       " - " & _
                       CStr(LastDisplayedYear)

End Property


'----------------------------------------
' Queries
'----------------------------------------

Public Function IsArrowEnabled(ByVal Direction As DP_NavigationDirection) As Boolean

    If Direction = DP_NAV_PREVIOUS Then
        IsArrowEnabled = FirstDisplayedYear > Year(DP_MinDate())
    Else
        IsArrowEnabled = LastDisplayedYear < Year(DP_MaxDate())
    End If

End Function


Public Function IsPeriodOutsideRange(ByVal PeriodDate As Date) As Boolean

    IsPeriodOutsideRange = PickerForm.IsPeriodOutsideRange(PeriodDate)

End Function


Public Function IsPeriodSelected(ByVal CellDate As Date) As Boolean

    IsPeriodSelected = Year(CellDate) = PickerForm.SelectedYear

End Function


Public Function IsCurrentPeriod(ByVal DisplayYear As Long) As Boolean

    IsCurrentPeriod = DisplayYear = Year(Date)

End Function


Public Function IsSecondaryPeriod(ByVal PeriodDate As Date) As Boolean

    IsSecondaryPeriod = False

End Function


'----------------------------------------
' Click
'----------------------------------------

Public Sub ArrowClicked(ByVal Direction As DP_NavigationDirection)

    If Not IsArrowEnabled(Direction) Then Exit Sub

    PickerForm.ChangeYearRange Direction * YEAR_COUNT

End Sub


Public Sub PeriodClicked(ByVal YearNumber As Long)

    If IsPeriodOutsideRange(DateSerial(YearNumber, 1, 1)) Then Exit Sub

    PickerForm.YearSelected YearNumber

End Sub


Public Sub ActionButtonClicked()

    PickerForm.YearSelected Year(Date)

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub ResetArrowHover()

    DP_ResetHoverCollection ArrowHandlers

End Sub


Public Sub ResetYearHover()

    DP_ResetHoverCollection YearHandlers

End Sub


Public Sub ResetPeriodHoverExcept(ByVal CurrentYearLabel As DP_CButtonPeriod)

    DP_ResetHoverCollectionExcept YearHandlers, CurrentYearLabel

End Sub


Public Sub ResetActionButtonHover()

    If Not CurrentYearButtonHandler Is Nothing Then
        CurrentYearButtonHandler.ResetHover
    End If

End Sub


Public Sub ResetAllHover()

    ResetArrowHover
    ResetYearHover
    ResetActionButtonHover

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    Dim Handler As Object

    If Not YearRangeHeaderHandler Is Nothing Then
        YearRangeHeaderHandler.SetVisible IsVisible
    End If

    For Each Handler In ArrowHandlers
        Handler.SetVisible IsVisible
    Next Handler

    For Each Handler In YearHandlers
        Handler.SetVisible IsVisible
    Next Handler

    If Not CurrentYearButtonHandler Is Nothing Then
        CurrentYearButtonHandler.SetVisible IsVisible
    End If

End Sub

