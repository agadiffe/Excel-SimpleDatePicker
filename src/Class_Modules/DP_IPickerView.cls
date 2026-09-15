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

Public Sub InitializeView(ByVal Form As DP_frmDatePicker)
End Sub


'----------------------------------------
' View
'----------------------------------------

Public Sub UpdateView()
End Sub

Public Sub RefreshTheme()
End Sub

Public Sub SetVisible(ByVal IsVisible As Boolean)
End Sub

Public Sub ResetAllHover()
End Sub


'----------------------------------------
' Properties
'----------------------------------------

Public Property Get ParentForm() As DP_frmDatePicker
End Property


'----------------------------------------
' Queries
'----------------------------------------

Public Function IsArrowEnabled(ByVal Direction As DP_NavigationDirection) As Boolean
End Function

Public Function IsDateSelected(ByVal CellDate As Date) As Boolean
End Function

Public Function IsCurrentPeriod(ByVal CellDate As Date) As Boolean
End Function

Public Function IsSecondaryPeriod(ByVal CellDate As Date) As Boolean
End Function


'----------------------------------------
' Interaction
'----------------------------------------

Public Sub ArrowClicked(ByVal Direction As DP_NavigationDirection)
End Sub

Public Sub HeaderClicked()
End Sub

Public Sub ActionButtonClicked()
End Sub

Public Sub GridButtonClicked(ByVal SelectedDate As Date)
End Sub

