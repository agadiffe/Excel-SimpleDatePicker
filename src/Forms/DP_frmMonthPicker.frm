VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DP_frmMonthPicker 
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "DP_frmMonthPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DP_frmMonthPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' References
'----------------------------------------

Private Const MONTH_COUNT As Long = _
    MONTH_GRID_COLUMNS * MONTH_GRID_ROWS

Private ParentPicker As DP_frmDatePicker
Private CurrentYear As Long

Private MonthsInitialized As Boolean

Private ArrowHandlers As Collection
Private MonthLabelHandlers As Collection
Private YearLabelHandler As DP_CHeaderLabel
Private CurrentMonthButtonHandler As DP_CActionButton


'----------------------------------------
' Month picker
'----------------------------------------

Public Sub ShowMonths(ByVal Picker As DP_frmDatePicker)

    Set ParentPicker = Picker

    CurrentYear = Year(ParentPicker.GetCurrentMonth)

    BuildMonths
    ShowPopupNextToCell Me, DP_frmDatePicker.GetTargetCell

End Sub


Private Sub BuildMonths()

    If Not MonthsInitialized Then
        InitializeMonths
        MonthsInitialized = True
    Else
        RefreshMonths
    End If

End Sub


Private Sub InitializeMonths()

    Dim ArrowHandler As DP_CArrowLabel

    Set MonthLabelHandlers = New Collection
    Set ArrowHandlers = New Collection

    ClearPickerControls Me
    InitializePicker Me, MONTHPICKER_WIDTH, MONTHPICKER_HEIGHT

    ' Year
    '--------------------

    Set YearLabelHandler = New DP_CHeaderLabel
    YearLabelHandler.Setup Me, CStr(CurrentYear)

    ' Previous year
    '--------------------

    Set ArrowHandler = New DP_CArrowLabel
    ArrowHandler.Setup Me, "PREV_YEAR", ChrW(&H25B2), ARROW_PREV_LEFT

    ArrowHandlers.Add ArrowHandler

    ' Next year
    '--------------------

    Set ArrowHandler = New DP_CArrowLabel
    ArrowHandler.Setup Me, "NEXT_YEAR", ChrW(&H25BC), ARROW_NEXT_LEFT

    ArrowHandlers.Add ArrowHandler

    ' Months
    '--------------------

    CreateMonthLabels

    ' Go To Current Month button
    '--------------------

    Set CurrentMonthButtonHandler = New DP_CActionButton
    CurrentMonthButtonHandler.Setup Me, "GO_TO_CURRENT_MONTH"

End Sub


Private Sub RefreshMonths()

    ResetAllHover

    YearLabelHandler.SetCaption CStr(CurrentYear)

    UpdateMonthLabels

End Sub


'----------------------------------------
' Build Months helpers
'----------------------------------------

Private Sub CreateMonthLabels()

    Dim MonthHandler As DP_CPeriodLabel
    Dim CellIndex As Long
    Dim MonthNumber As Long
    Dim CellDate As Date

    Dim ColumnIndex As Long
    Dim RowIndex As Long

    For CellIndex = 0 To MONTH_COUNT - 1

        MonthNumber = CellIndex + 1
        CellDate = GetMonthDate(CellIndex)

        ColumnIndex = GridColumn(CellIndex, MONTH_GRID_COLUMNS)
        RowIndex = GridRow(CellIndex, MONTH_GRID_COLUMNS)

        Set MonthHandler = New DP_CPeriodLabel

        MonthHandler.Setup Me, _
                           "MONTH", _
                           MonthNumber, _
                           CellDate, _
                           GridLeft(ColumnIndex, PERIOD_GRID_LEFT, MONTH_GRID_CELL_WIDTH), _
                           GridTop(RowIndex, PERIOD_GRID_TOP, PERIOD_CELL_HEIGHT), _
                           MONTH_CELL_WIDTH

        MonthHandler.SetState IsCurrentMonth(MonthNumber), _
                              IsSelectedMonth(MonthNumber), _
                              Year(CellDate) <> CurrentYear

        MonthLabelHandlers.Add MonthHandler

    Next CellIndex

End Sub


Private Sub UpdateMonthLabels()

    Dim CellIndex As Long
    Dim MonthNumber As Long
    Dim CellDate As Date

    Dim MonthHandler As DP_CPeriodLabel

    For CellIndex = 0 To MONTH_COUNT - 1

        MonthNumber = CellIndex + 1
        CellDate = GetMonthDate(CellIndex)

        Set MonthHandler = MonthLabelHandlers.Item(CellIndex + 1)

        MonthHandler.UpdatePeriod CellDate, _
                                  MonthNumber, _
                                  IsCurrentMonth(MonthNumber), _
                                  IsSelectedMonth(MonthNumber)

    Next CellIndex

End Sub


Private Function GetMonthDate(ByVal CellIndex As Long) As Date

    GetMonthDate = DateSerial(CurrentYear, CellIndex + 1, 1)

End Function


Private Function IsCurrentMonth(ByVal MonthNumber As Long) As Boolean

    IsCurrentMonth = MonthNumber = Month(Date) And _
                     CurrentYear = Year(Date)

End Function


Private Function IsSelectedMonth(ByVal MonthNumber As Long) As Boolean

    IsSelectedMonth = MonthNumber = Month(ParentPicker.GetCurrentMonth) And _
                      CurrentYear = Year(ParentPicker.GetCurrentMonth)

End Function


'----------------------------------------
' Month picker helpers
'----------------------------------------

Public Sub MonthLabelClicked(ByVal MonthNumber As Long)

    ParentPicker.SetMonthYear MonthNumber, CurrentYear

    Me.Hide

End Sub


Public Sub HeaderClicked()

    Me.Hide
    DP_frmYearPicker.ShowYears Me

    Unload DP_frmYearPicker
    Me.Show

End Sub


Public Sub ArrowClicked(ByVal Action As String)

    Select Case Action
        Case "PREV_YEAR"
            CurrentYear = CurrentYear - 1
        Case "NEXT_YEAR"
            CurrentYear = CurrentYear + 1
    End Select

    BuildMonths

End Sub


Public Function GetCurrentYear() As Long

    GetCurrentYear = CurrentYear

End Function


Public Sub SetYear(ByVal YearNumber As Long)

    CurrentYear = YearNumber
    ParentPicker.SetMonthYear Month(ParentPicker.GetCurrentMonth), CurrentYear

    BuildMonths

End Sub


Public Sub GoToCurrentMonth()

    CurrentYear = Year(Date)
    ParentPicker.SetMonthYear Month(Date), CurrentYear

    Me.Hide

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub ResetArrowHover()

    ResetHoverCollection ArrowHandlers

End Sub


Public Sub ResetMonthHover()

    ResetHoverCollection MonthLabelHandlers

End Sub


Public Sub ResetHeaderHover()

    If Not YearLabelHandler Is Nothing Then
        YearLabelHandler.ResetHover
    End If

End Sub


Public Sub ResetCurrentMonthHover()

    If Not CurrentMonthButtonHandler Is Nothing Then
        CurrentMonthButtonHandler.ResetHover
    End If

End Sub


Public Sub ResetAllHover()

    ResetArrowHover
    ResetHeaderHover
    ResetMonthHover
    ResetCurrentMonthHover

End Sub


Public Sub ResetMonthHoverExcept(ByVal CurrentMonth As DP_CPeriodLabel)

    ResetHoverCollectionExcept MonthLabelHandlers, CurrentMonth

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


Private Sub UserForm_Activate()

    RemoveUserFormTitleBar Me.Caption

End Sub

