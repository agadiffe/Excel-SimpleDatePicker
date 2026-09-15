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

Private CurrentPickerSize As DP_PickerSize

Private InitialCellDateValue As Date
Private CurrentMonthValue As Date
Private SelectedMonthValue As Date
Private SelectedYearValue As Long

Private CalendarView As DP_IPickerView
Private MonthView As DP_IPickerView
Private YearView As DP_IPickerView


'----------------------------------------
' Initialization
'----------------------------------------

Private Sub UserForm_Initialize()

    DP_IsDarkMode = DP_IsDarkThemeActive()

    DP_InitializePicker Me, DP_DATEPICKER_WIDTH, DP_DATEPICKER_HEIGHT

    Set CalendarView = New DP_CViewCalendar
    Set MonthView = New DP_CViewMonth
    Set YearView = New DP_CViewYear

    CalendarView.InitializeView Me
    MonthView.InitializeView Me
    YearView.InitializeView Me

End Sub


Public Sub ShowPicker(ByVal Cell As Range)

    Set TargetCell = Cell

    If IsDate(TargetCell.Value) Then
        InitialCellDateValue = CDate(TargetCell.Value)
    Else
        InitialCellDateValue = Date
    End If

    SelectedMonthValue = DateSerial(Year(InitialCellDateValue), Month(InitialCellDateValue), 1)
    SelectedYearValue = Year(SelectedMonthValue)
    CurrentMonthValue = SelectedMonthValue

    RefreshPickerTheme
    ShowPickerView DP_SIZE_DATE
    DP_ShowPopupNextToCell Me, TargetCell

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

    Dim PickerView As DP_IPickerView

    Set PickerView = GetPickerView(PickerSize)

    PickerView.ResetAllHover

    DP_SetPickerSize Me, PickerSize
    CurrentPickerSize = PickerSize

    CalendarView.SetVisible PickerSize = DP_SIZE_DATE
    MonthView.SetVisible PickerSize = DP_SIZE_MONTH
    YearView.SetVisible PickerSize = DP_SIZE_YEAR

    PickerView.UpdateView

End Sub


Private Function GetPickerView(ByVal PickerSize As DP_PickerSize) As DP_IPickerView

    Select Case PickerSize
        Case DP_SIZE_DATE
            Set GetPickerView = CalendarView
        Case DP_SIZE_MONTH
            Set GetPickerView = MonthView
        Case DP_SIZE_YEAR
            Set GetPickerView = YearView
    End Select

End Function


'----------------------------------------
' Properties
'----------------------------------------

Public Property Get SelectedMonth() As Date

    SelectedMonth = SelectedMonthValue

End Property


Public Property Get SelectedYear() As Long

    SelectedYear = SelectedYearValue

End Property


Public Property Get CurrentMonth() As Date

    CurrentMonth = CurrentMonthValue

End Property


Public Property Get InitialCellDate() As Date

    InitialCellDate = InitialCellDateValue

End Property


'----------------------------------------
' Date Selection
'----------------------------------------

Public Sub DaySelected(ByVal SelectedDate As Date)

    Dim ExistingTime As Double

    If IsDate(InitialCellDateValue) Then
        ExistingTime = TimeValue(InitialCellDateValue)
        TargetCell.Value = DateValue(SelectedDate) + ExistingTime
    Else
        TargetCell.Value = SelectedDate
    End If

    Me.Hide

End Sub


Public Sub MonthSelected(ByVal MonthNumber As Long, _
                         ByVal YearNumber As Long)

    SelectedYearValue = YearNumber
    SelectedMonthValue = DateSerial(SelectedYearValue, MonthNumber, 1)
    CurrentMonthValue = SelectedMonthValue

    ShowPickerView DP_SIZE_DATE

End Sub


Public Sub YearSelected(ByVal YearNumber As Long)

    SelectedYearValue = YearNumber
    CurrentMonthValue = DateSerial(SelectedYearValue, Month(CurrentMonthValue), 1)

    ShowPickerView DP_SIZE_MONTH

End Sub


'----------------------------------------
' Navigation
'----------------------------------------

Public Sub ChangePeriod(ByVal Amount As Long, _
                        ByVal Interval As String)

    CurrentMonthValue = DateAdd(Interval, Amount, CurrentMonthValue)
    GetPickerView(CurrentPickerSize).UpdateView

End Sub


Public Sub HeaderClicked(ByVal PickerSize As DP_PickerSize)

    ShowPickerView PickerSize

End Sub


'----------------------------------------
' Events
'----------------------------------------

Private Sub UserForm_MouseMove(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal x As Single, _
                               ByVal Y As Single)

    GetPickerView(CurrentPickerSize).ResetAllHover

End Sub

