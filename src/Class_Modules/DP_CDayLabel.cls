VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CDayLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' References
'----------------------------------------

Private WithEvents DayLabel As MSForms.Label
Attribute DayLabel.VB_VarHelpID = -1
Private WithEvents DayTextLabel As MSForms.Label
Attribute DayTextLabel.VB_VarHelpID = -1

Private ParentForm As DP_frmDatePicker
Private CellDate As Date

Private HoverHandler As DP_CHoverHandler


'----------------------------------------
' Setup
'----------------------------------------

Public Sub Setup(ByVal Form As DP_frmDatePicker, _
                 ByVal CalendarCellDate As Date, _
                 ByVal LeftPosition As Single, _
                 ByVal TopPosition As Single)

    Set ParentForm = Form
    CellDate = CalendarCellDate

    ' Outer clickable square
    '--------------------

    Set DayLabel = ParentForm.Controls.Add("Forms.Label.1")

    With DayLabel
        .Left = LeftPosition
        .Top = TopPosition
        .Width = DP_CALENDAR_CELL_WIDTH
        .Height = DP_CALENDAR_CELL_HEIGHT
    End With

    ' Day number
    '--------------------

    Set DayTextLabel = ParentForm.Controls.Add("Forms.Label.1")

    With DayTextLabel
        .Caption = GetDayTextCaption()
        .Font.Size = DP_FONT_SIZE_DAY * DP_GetFontScale(ParentForm.Width)
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
        .WordWrap = False
        .AutoSize = True

        .Top = DP_VerticalTextTop(DayLabel.Top, _
                                  DayLabel.Height, _
                                  .Height)

        .AutoSize = False
        .Left = DayLabel.Left
        .Width = DayLabel.Width
        .Height = DayLabel.Top + DayLabel.Height - .Top
    End With

    ' Hover
    '--------------------

    Set HoverHandler = New DP_CHoverHandler
    HoverHandler.Setup DayTextLabel, DayLabel

    ' Initial state
    '--------------------

    UpdateDate CalendarCellDate

End Sub


'----------------------------------------
' Update
'----------------------------------------

Public Sub UpdateDate(ByVal CalendarCellDate As Date)

    CellDate = CalendarCellDate

    DayTextLabel.Caption = GetDayTextCaption()

    ' Default state
    '--------------------

    DayLabel.BackColor = DP_ColorBg()
    DayTextLabel.ForeColor = DP_ColorText()
    DayTextLabel.Font.Bold = False

    ' Other month
    '--------------------

    If Month(CellDate) <> Month(ParentForm.GetCurrentMonth) Then
        DayTextLabel.ForeColor = DP_ColorTextSecondary()
    End If

    ' Today
    '--------------------

    If CellDate = DateValue(Date) Then
        DayLabel.BackColor = DP_ColorBgToday()
        DayTextLabel.ForeColor = DP_ColorTextToday()
        DayTextLabel.Font.Bold = True
    End If

    ' Selected date
    '--------------------

    If ParentForm.HasSelectedDate(CellDate) Then
        DayLabel.BackColor = DP_ColorBgSelected()
        DayTextLabel.ForeColor = DP_ColorTextSelected()
        DayTextLabel.Font.Bold = True
    End If

End Sub


'----------------------------------------
' Day
'----------------------------------------

Private Sub DayLabel_Click()

    ParentForm.DayLabelClicked CellDate

End Sub


Private Sub DayTextLabel_Click()

    ParentForm.DayLabelClicked CellDate

End Sub


Private Function GetDayTextCaption() As String

    GetDayTextCaption = CStr(Day(CellDate))

End Function


'----------------------------------------
' Hover
'----------------------------------------

Private Sub DayLabel_MouseMove(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Single, _
                               ByVal Y As Single)

    SetHover

End Sub


Private Sub DayTextLabel_MouseMove(ByVal Button As Integer, _
                                   ByVal Shift As Integer, _
                                   ByVal X As Single, _
                                   ByVal Y As Single)

    SetHover

End Sub


Private Sub SetHover()

    ParentForm.ResetArrowHover
    ParentForm.ResetHeaderHover
    ParentForm.ResetTodayHover
    ParentForm.ResetDayHoverExcept Me

    HoverHandler.SetBackgroundHover

End Sub


Public Sub ResetHover()

    HoverHandler.ResetHover

End Sub

