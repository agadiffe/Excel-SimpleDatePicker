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

Private ParentView As DP_CViewCalendar
Private CellDate As Date

Private IsHovered As Boolean
Private IsPressed As Boolean


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub Initialize(ByVal View As DP_CViewCalendar, _
                      ByVal CalendarCellDate As Date, _
                      ByVal LeftPosition As Single, _
                      ByVal TopPosition As Single)

    Set ParentView = View
    CellDate = CalendarCellDate

    Set DayBackground = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With DayBackground
        .Left = LeftPosition
        .Top = TopPosition
        .Width = DP_CALENDAR_CELL_WIDTH
        .Height = DP_CALENDAR_CELL_HEIGHT
    End With

    Set DayText = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With DayText
        .Caption = GetDayTextCaption()
        .Font.Size = DP_FONT_SIZE_DAY * _
                     DP_GetFontScale(DP_DATEPICKER_WIDTH)
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


Private Function GetDayTextCaption() As String

    GetDayTextCaption = CStr(Day(CellDate))

End Function


'----------------------------------------
' Render
'----------------------------------------

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

    If Month(CellDate) <> Month(ParentView.ParentForm.CurrentMonth) Then
        TextColor = DP_ColorTextSecondary()
    End If

    ' Disabled
    '--------------------

    If ParentView.IsDayOutsideRange(CellDate) Then
        TextColor = DP_ColorTextDisabled()
    End If

    ' Today
    '--------------------

    If CellDate = DateValue(Date) Then
        BackgroundColor = DP_ColorBgToday()
        TextColor = DP_ColorTextToday()
    End If

    ' Selected
    '--------------------

    If ParentView.IsDaySelected(CellDate) Then
        BackgroundColor = DP_ColorBgSelected()
        TextColor = DP_ColorTextSelected()
        IsBold = True
    End If

    ' Hover
    '--------------------

    If IsHovered Then
        If ParentView.IsDaySelected(CellDate) Then
            BackgroundColor = DP_ColorBgSelectedHover()
        Else
            BackgroundColor = DP_ColorBgHover()
        End If
    End If

    ' Pressed
    '--------------------

    If IsPressed Then
        If ParentView.IsDaySelected(CellDate) Then
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


Public Sub UpdateDate(ByVal CalendarCellDate As Date)

    CellDate = CalendarCellDate
    DayText.Caption = GetDayTextCaption()

    Render

End Sub


'----------------------------------------
' Click
'----------------------------------------

Private Sub DayBackground_Click()

    HandleDayClick

End Sub


Private Sub DayText_Click()

    HandleDayClick

End Sub


Private Sub HandleDayClick()

    If ParentView.IsDayOutsideRange(CellDate) Then Exit Sub

    ParentView.DayClicked CellDate

End Sub


'----------------------------------------
' Pressed
'----------------------------------------

' Common
'--------------------

Private Sub SetPressed(ByVal Value As Boolean)

    If IsPressed = Value Then Exit Sub

    IsPressed = Value
    Render

End Sub


' Mouse down
'--------------------

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


Private Sub HandleMouseDown()

    If ParentView.IsDayOutsideRange(CellDate) Then Exit Sub

    SetPressed True

End Sub


' Mouse up
'--------------------

Private Sub DayBackground_MouseUp(ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)

    SetPressed False

End Sub


Private Sub DayText_MouseUp(ByVal Button As Integer, _
                            ByVal Shift As Integer, _
                            ByVal X As Single, _
                            ByVal Y As Single)

    SetPressed False

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Private Sub DayBackground_MouseMove(ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)

    SetHover

End Sub


Private Sub DayText_MouseMove(ByVal Button As Integer, _
                              ByVal Shift As Integer, _
                              ByVal X As Single, _
                              ByVal Y As Single)

    SetHover

End Sub


Private Sub SetHover()

    ParentView.ResetArrowHover
    ParentView.ResetHeaderHover
    ParentView.ResetTodayHover
    ParentView.ResetDayHoverExcept Me

    If Not ParentView.IsDayOutsideRange(CellDate) Then
        SetHovered True
    End If

End Sub


Private Sub SetHovered(ByVal Value As Boolean)

    If IsHovered = Value Then Exit Sub

    IsHovered = Value
    Render

End Sub


Public Sub ResetHover()

    SetHovered False

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    DayBackground.Visible = IsVisible
    DayText.Visible = IsVisible

End Sub

