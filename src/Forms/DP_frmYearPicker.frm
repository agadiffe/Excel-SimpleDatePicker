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
    ShowPopupNextToCell Me, DP_frmDatePicker.GetTargetCell

End Sub


Private Sub BuildYears()

    If Not YearsInitialized Then
        InitializeYears
        YearsInitialized = True
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

    ClearPickerControls Me
    InitializePicker Me, YEARPICKER_WIDTH, YEARPICKER_HEIGHT

    ' Year range
    '--------------------

    FirstYear = GetFirstDisplayedYear(CurrentYear)
    LastYear = FirstYear + YEAR_BLOCK_SIZE - 1

    Set YearRangeLabelHandler = New DP_CHeaderLabel
    YearRangeLabelHandler.Setup Me, CStr(FirstYear) & " - " & CStr(LastYear)

    YearRangeLabelHandler.SetInteractive False

    ' Previous year range
    '--------------------

    Set ArrowHandler = New DP_CArrowLabel
    ArrowHandler.Setup Me, "PREV_RANGE_YEAR", ChrW(&H25B2), ARROW_PREV_LEFT

    ArrowHandlers.Add ArrowHandler

    ' Next year range
    '--------------------

    Set ArrowHandler = New DP_CArrowLabel
    ArrowHandler.Setup Me, "NEXT_RANGE_YEAR", ChrW(&H25BC), ARROW_NEXT_LEFT

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

End Sub


Private Sub UpdateYearRangeHeader()

    Dim FirstYear As Long
    Dim LastYear As Long

    FirstYear = GetFirstDisplayedYear(CurrentYear)
    LastYear = FirstYear + YEAR_BLOCK_SIZE - 1

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

    For CellIndex = 0 To YEAR_BLOCK_SIZE - 1

        DisplayYear = FirstYear + CellIndex
        CellDate = DateSerial(DisplayYear, 1, 1)

        ColumnIndex = GridColumn(CellIndex, YEAR_GRID_COLUMNS)
        RowIndex = GridRow(CellIndex, YEAR_GRID_COLUMNS)

        Set YearHandler = New DP_CPeriodLabel

        YearHandler.Setup Me, _
                          "YEAR", _
                          DisplayYear, _
                          CellDate, _
                          GridLeft(ColumnIndex, PERIOD_GRID_LEFT, YEAR_GRID_CELL_WIDTH), _
                          GridTop(RowIndex, PERIOD_GRID_TOP, PERIOD_CELL_HEIGHT), _
                          YEAR_CELL_WIDTH

        YearHandler.SetState IsCurrentYear(DisplayYear), _
                             IsSelectedYear(DisplayYear)

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

    For CellIndex = 0 To YEAR_BLOCK_SIZE - 1

        DisplayYear = FirstYear + CellIndex
        CellDate = DateSerial(DisplayYear, 1, 1)

        Set YearHandler = YearLabelHandlers.Item(CellIndex + 1)

        YearHandler.UpdatePeriod CellDate, _
                                 DisplayYear, _
                                 IsCurrentYear(DisplayYear), _
                                 IsSelectedYear(DisplayYear)

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

Private Function GetFirstDisplayedYear(ByVal SelectedYear As Long) As Long

    GetFirstDisplayedYear = ((SelectedYear - 1) \ YEAR_BLOCK_SIZE) * YEAR_BLOCK_SIZE + 1

End Function


Public Sub YearLabelClicked(ByVal YearNumber As Long)

    ParentPicker.SetYear YearNumber

    Me.Hide

End Sub


Public Sub ArrowClicked(ByVal Action As String)

    Select Case Action
        Case "PREV_RANGE_YEAR"
            CurrentYear = CurrentYear - YEAR_BLOCK_SIZE
        Case "NEXT_RANGE_YEAR"
            CurrentYear = CurrentYear + YEAR_BLOCK_SIZE
    End Select

    BuildYears

End Sub


Public Sub GoToCurrentYear()

    CurrentYear = Year(Date)
    ParentPicker.SetYear CurrentYear

    Me.Hide

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub ResetArrowHover()

    ResetHoverCollection ArrowHandlers

End Sub


Public Sub ResetYearHover()

    ResetHoverCollection YearLabelHandlers

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

    ResetHoverCollectionExcept YearLabelHandlers, CurrentYearLabel

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

