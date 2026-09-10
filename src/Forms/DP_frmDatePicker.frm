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
' Declarations
'----------------------------------------

Private TargetCell As Range

Private CurrentPickerSizeValue As DP_PickerSize

Private InitialCellDate As Date
Private CurrentMonthValue As Date
Private SelectedMonthValue As Date
Private SelectedYearValue As Long

Private CalendarView As DP_CViewCalendar
Private MonthView As DP_CViewMonth
Private YearView As DP_CViewYear


'----------------------------------------
' Initialization
'----------------------------------------

Private Sub UserForm_Initialize()

    DP_InitializePicker Me, DP_DATEPICKER_WIDTH, DP_DATEPICKER_HEIGHT

    DP_IsDarkMode = DP_IsDarkThemeActive()
    Me.BackColor = DP_ColorBg()

    Set CalendarView = New DP_CViewCalendar
    Set MonthView = New DP_CViewMonth
    Set YearView = New DP_CViewYear

    CalendarView.Attach Me
    MonthView.Attach Me
    YearView.Attach Me

End Sub


Public Sub ShowPicker(ByVal Cell As Range)

    Set TargetCell = Cell

    If IsDate(Cell.Value) Then
        InitialCellDate = CDate(Cell.Value)
    Else
        InitialCellDate = Date
    End If

    SelectedMonthValue = DateSerial(Year(InitialCellDate), Month(InitialCellDate), 1)
    SelectedYearValue = Year(SelectedMonthValue)
    CurrentMonthValue = SelectedMonthValue

    RefreshPickerTheme
    ShowPickerView DP_SIZE_DATE
    DP_ShowPopupNextToCell Me, Cell

End Sub


Private Sub RefreshPickerTheme()

    Dim CurrentDarkMode As Boolean

    CurrentDarkMode = DP_IsDarkThemeActive()

    If CurrentDarkMode = DP_IsDarkMode Then Exit Sub

    DP_IsDarkMode = CurrentDarkMode

    Me.BackColor = DP_ColorBg()

    CalendarView.RefreshTheme
    MonthView.RefreshTheme
    YearView.RefreshTheme

End Sub


'----------------------------------------
' View
'----------------------------------------

Private Sub ShowPickerView(ByVal PickerSize As DP_PickerSize)

    ResetAllHover

    DP_SetPickerSize Me, PickerSize
    CurrentPickerSizeValue = PickerSize

    CalendarView.SetVisible PickerSize = DP_SIZE_DATE
    MonthView.SetVisible PickerSize = DP_SIZE_MONTH
    YearView.SetVisible PickerSize = DP_SIZE_YEAR

    Select Case PickerSize
        Case DP_SIZE_DATE
            CalendarView.UpdateView
        Case DP_SIZE_MONTH
            MonthView.UpdateView
        Case DP_SIZE_YEAR
            YearView.UpdateView
    End Select

End Sub


'----------------------------------------
' Properties
'----------------------------------------

Public Property Get CurrentPickerSize() As DP_PickerSize

    CurrentPickerSize = CurrentPickerSizeValue

End Property


Public Property Get SelectedMonth() As Date

    SelectedMonth = SelectedMonthValue

End Property


Public Property Get SelectedYear() As Long

    SelectedYear = SelectedYearValue

End Property


Public Property Get CurrentMonth() As Date

    CurrentMonth = CurrentMonthValue

End Property


'----------------------------------------
' Queries
'----------------------------------------

Public Function IsDayOutsideRange(ByVal DayDate As Date) As Boolean

    IsDayOutsideRange = DayDate < DP_MinDate() Or DayDate > DP_MaxDate()

End Function


Public Function IsPeriodOutsideRange(ByVal PeriodDate As Date) As Boolean

    IsPeriodOutsideRange = PeriodDate < DP_MinDate() Or PeriodDate > DP_MaxDate()

End Function


Public Function IsDaySelected(ByVal CellDate As Date) As Boolean

    IsDaySelected = CellDate = DateValue(InitialCellDate)

End Function


'----------------------------------------
' Date Selection
'----------------------------------------

Public Sub DaySelected(ByVal SelectedDate As Date)

    SetTargetDate SelectedDate
    Me.Hide

End Sub


Public Sub MonthSelected(ByVal MonthNumber As Long, _
                         ByVal YearNumber As Long)

    SetMonthYear MonthNumber, YearNumber
    ShowPickerView DP_SIZE_DATE

End Sub


Public Sub YearSelected(ByVal YearNumber As Long)

    SetYear YearNumber
    ShowPickerView DP_SIZE_MONTH

End Sub


Private Sub SetTargetDate(ByVal NewDate As Date)

    Dim ExistingTime As Double

    If IsDate(InitialCellDate) Then
        ExistingTime = TimeValue(InitialCellDate)
        TargetCell.Value = DateValue(NewDate) + ExistingTime
    Else
        TargetCell.Value = NewDate
    End If

End Sub


Public Sub SetMonthYear(ByVal MonthNumber As Long, _
                        ByVal YearNumber As Long)

    SelectedYearValue = YearNumber
    SelectedMonthValue = DateSerial(YearNumber, MonthNumber, 1)
    CurrentMonthValue = SelectedMonthValue

End Sub


Public Sub SetYear(ByVal YearNumber As Long)

    SelectedYearValue = YearNumber
    CurrentMonthValue = DateSerial(YearNumber, Month(CurrentMonthValue), 1)

End Sub


'----------------------------------------
' Navigation
'----------------------------------------

Public Sub ChangeMonth(ByVal Amount As Long)

    CurrentMonthValue = DateAdd("m", Amount, CurrentMonthValue)
    CalendarView.UpdateView

End Sub


Public Sub ChangeYear(ByVal Amount As Long)

    CurrentMonthValue = DateAdd("yyyy", Amount, CurrentMonthValue)
    MonthView.UpdateView

End Sub


Public Sub ChangeYearRange(ByVal Amount As Long)

    CurrentMonthValue = DateAdd("yyyy", Amount, CurrentMonthValue)
    YearView.UpdateView

End Sub


Public Sub CalendarHeaderClicked()

    ShowPickerView DP_SIZE_MONTH

End Sub


Public Sub MonthHeaderClicked()

    ShowPickerView DP_SIZE_YEAR

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub ResetAllHover()

    Select Case CurrentPickerSizeValue
        Case DP_SIZE_DATE
            CalendarView.ResetAllHover
        Case DP_SIZE_MONTH
            MonthView.ResetAllHover
        Case DP_SIZE_YEAR
            YearView.ResetAllHover
    End Select

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

