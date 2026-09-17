VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_IPickerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' Initialization
'----------------------------------------

Sub InitializeView(ByVal Form As DP_frmDatePicker)

'----------------------------------------
' View
'----------------------------------------

Sub UpdateView()
Sub RefreshTheme()
Sub SetVisible(ByVal IsVisible As Boolean)
Sub SetHoveredButton(ByVal Button As Object)
Sub ClearHoveredButton()

'----------------------------------------
' Properties
'----------------------------------------

Property Get ParentForm() As DP_frmDatePicker

'----------------------------------------
' Queries
'----------------------------------------

Function IsArrowEnabled(ByVal Direction As DP_NavigationDirection) As Boolean
Function IsDateSelected(ByVal CellDate As Date) As Boolean
Function IsCurrentPeriod(ByVal CellDate As Date) As Boolean
Function IsSecondaryPeriod(ByVal CellDate As Date) As Boolean

'----------------------------------------
' Interaction
'----------------------------------------

Sub ArrowClicked(ByVal Direction As DP_NavigationDirection)
Sub HeaderClicked()
Sub ActionButtonClicked()
Sub GridButtonClicked(ByVal SelectedDate As Date)

