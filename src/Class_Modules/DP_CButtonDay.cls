VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CButtonDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Private WithEvents DayBackground As MSForms.Label
Attribute DayBackground.VB_VarHelpID = -1
Private WithEvents DayText As MSForms.Label
Attribute DayText.VB_VarHelpID = -1

Private ParentView As DP_IPickerView
Private CellDate As Date

Private IsHovered As Boolean
Private IsPressed As Boolean


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub Initialize(ByVal View As DP_CViewCalendar, _
                      ByVal CellDateValue As Date, _
                      ByVal LeftPosition As Single, _
                      ByVal TopPosition As Single)

    Set ParentView = View
    CellDate = CellDateValue

    Set DayBackground = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With DayBackground
        .Left = LeftPosition
        .Top = TopPosition
        .Width = DP_CALENDAR_CELL_WIDTH
        .Height = DP_CALENDAR_CELL_HEIGHT
    End With

    Set DayText = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With DayText
        .Caption = CStr(Day(CellDate))
        .Font.Size = DP_FONT_SIZE_DAY * DP_GetFontScale(DP_DATEPICKER_WIDTH)
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
        .WordWrap = False

        .AutoSize = True
        .Top = DP_VerticalTextTop(DayBackground.Top, _
                                  DayBackground.Height, _
                                  .Height)
        .AutoSize = False

        .Left = DayBackground.Left
        .Width = DayBackground.Width
        .Height = DayBackground.Top + DayBackground.Height - .Top
    End With

    Render

End Sub


'----------------------------------------
' Render
'----------------------------------------

Public Sub UpdateDate(ByVal CellDateValue As Date)

    CellDate = CellDateValue
    DayText.Caption = CStr(Day(CellDate))

    Render

End Sub


Private Sub Render()

    Dim BackgroundColor As Long
    Dim TextColor As Long
    Dim IsBold As Boolean

    ' Default
    '--------------------

    BackgroundColor = DP_ColorBg()
    TextColor = DP_ColorText()
    IsBold = False

    ' Other month
    '--------------------

    If ParentView.IsSecondaryPeriod(CellDate) Then
        TextColor = DP_ColorTextSecondary()
    End If

    ' Disabled
    '--------------------

    If DP_IsDateOutsideRange(CellDate) Then
        TextColor = DP_ColorTextDisabled()
    End If

    ' Today
    '--------------------

    If ParentView.IsCurrentPeriod(CellDate) Then
        BackgroundColor = DP_ColorBgToday()
        TextColor = DP_ColorTextToday()
    End If

    ' Selected
    '--------------------

    If ParentView.IsDateSelected(CellDate) Then
        BackgroundColor = DP_ColorBgSelected()
        TextColor = DP_ColorTextSelected()
        IsBold = True
    End If

    ' Hover
    '--------------------

    If IsHovered Then
        If ParentView.IsDateSelected(CellDate) Then
            BackgroundColor = DP_ColorBgSelectedHover()
        Else
            BackgroundColor = DP_ColorBgHover()
        End If
    End If

    ' Pressed
    '--------------------

    If IsPressed Then
        If ParentView.IsDateSelected(CellDate) Then
            BackgroundColor = DP_ColorBgSelectedPressed()
        Else
            BackgroundColor = DP_ColorBgPressed()
        End If
    End If

    ' Apply
    '--------------------

    DayBackground.BackColor = BackgroundColor
    DayText.ForeColor = TextColor
    DayText.Font.Bold = IsBold

End Sub


'----------------------------------------
' Properties
'----------------------------------------

Private Property Let Hovered(ByVal Value As Boolean)

    If IsHovered = Value Then Exit Property

    IsHovered = Value
    Render

End Property


Private Property Let Pressed(ByVal Value As Boolean)

    If IsPressed = Value Then Exit Property

    IsPressed = Value
    Render

End Property


'----------------------------------------
' Click
'----------------------------------------

Private Sub HandleDayClick()

    If DP_IsDateOutsideRange(CellDate) Then Exit Sub

    ParentView.GridButtonClicked CellDate

End Sub


Private Sub DayBackground_Click()

    HandleDayClick

End Sub


Private Sub DayText_Click()

    HandleDayClick

End Sub


'----------------------------------------
' Pressed
'----------------------------------------

' Mouse down
'--------------------

Private Sub HandleMouseDown()

    If DP_IsDateOutsideRange(CellDate) Then Exit Sub

    Pressed = True

End Sub


Private Sub DayBackground_MouseDown(ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)

    HandleMouseDown

End Sub


Private Sub DayText_MouseDown(ByVal Button As Integer, _
                              ByVal Shift As Integer, _
                              ByVal X As Single, _
                              ByVal Y As Single)

    HandleMouseDown

End Sub


' Mouse up
'--------------------

Private Sub DayBackground_MouseUp(ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)

    Pressed = False

End Sub


Private Sub DayText_MouseUp(ByVal Button As Integer, _
                            ByVal Shift As Integer, _
                            ByVal X As Single, _
                            ByVal Y As Single)

    Pressed = False

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub SetHovered(ByVal Value As Boolean)

    Hovered = Value

End Sub


Private Sub HandleMouseHover()

    If DP_IsDateOutsideRange(CellDate) Then
        ParentView.ClearHoveredButton
    Else
        ParentView.SetHoveredButton Me
    End If

End Sub


Private Sub DayBackground_MouseMove(ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)

    HandleMouseHover

End Sub


Private Sub DayText_MouseMove(ByVal Button As Integer, _
                              ByVal Shift As Integer, _
                              ByVal X As Single, _
                              ByVal Y As Single)

    HandleMouseHover

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    DayBackground.Visible = IsVisible
    DayText.Visible = IsVisible

End Sub

