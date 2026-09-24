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
Private TargetWindowValue As Excel.Window

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

    #If Not Mac Then
        Me.Caption = DP_CreateWindowCaption()
    #End If

    DP_IsDarkMode = DP_IsDarkThemeActive()

    DP_InitializePicker Me

    Set CalendarView = New DP_CViewCalendar
    Set MonthView = New DP_CViewMonth
    Set YearView = New DP_CViewYear

    CalendarView.InitializeView Me
    MonthView.InitializeView Me
    YearView.InitializeView Me

End Sub


'----------------------------------------
' View management
'----------------------------------------

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

    RefreshFormTheme
    SwitchPickerView DP_SIZE_DATE
    DP_ShowPopupNextToCell Me, TargetCell, TargetWindowValue

End Sub


Private Sub RefreshFormTheme()

    Dim CurrentDarkMode As Boolean

    CurrentDarkMode = DP_IsDarkThemeActive()

    If CurrentDarkMode = DP_IsDarkMode Then Exit Sub

    DP_IsDarkMode = CurrentDarkMode
    Me.BackColor = DP_ColorBg()

    DP_ClearPictureCache
    DP_ApplyPickerBorder Me, DP_SIZE_DATE

End Sub


Private Sub SwitchPickerView(ByVal PickerSize As DP_PickerSize)

    Dim PickerView As DP_IPickerView

    Set PickerView = GetPickerView(PickerSize)

    PickerView.ClearHoveredButton

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

' Context
'--------------------

Public Property Set TargetWindow(ByVal TargetWindow As Excel.Window)

    Set TargetWindowValue = TargetWindow

End Property


Public Property Get TargetWindow() As Excel.Window

    Set TargetWindow = TargetWindowValue

End Property


' State
'--------------------

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
' Selection
'----------------------------------------

Public Sub DaySelected(ByVal SelectedDate As Date)

    Dim ExistingTime As Double

    ExistingTime = TimeValue(InitialCellDateValue)
    TargetCell.Value = DateValue(SelectedDate) + ExistingTime

    Me.Hide

End Sub


Public Sub MonthSelected(ByVal MonthNumber As Long, _
                         ByVal YearNumber As Long)

    SelectedYearValue = YearNumber
    SelectedMonthValue = DateSerial(SelectedYearValue, MonthNumber, 1)
    CurrentMonthValue = SelectedMonthValue

    SwitchPickerView DP_SIZE_DATE

End Sub


Public Sub YearSelected(ByVal YearNumber As Long)

    SelectedYearValue = YearNumber
    CurrentMonthValue = DateSerial(SelectedYearValue, Month(CurrentMonthValue), 1)

    SwitchPickerView DP_SIZE_MONTH

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

    SwitchPickerView PickerSize

End Sub


'----------------------------------------
' Events
'----------------------------------------

Private Sub UserForm_MouseMove(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Single, _
                               ByVal Y As Single)

    GetPickerView(CurrentPickerSize).ClearHoveredButton

End Sub


Private Sub UserForm_Terminate()

    DP_RemovePickerForm Me

End Sub

