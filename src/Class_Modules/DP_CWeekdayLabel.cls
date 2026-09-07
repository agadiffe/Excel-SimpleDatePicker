VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CWeekdayLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' References
'----------------------------------------

Private WithEvents WeekdayLabel As MSForms.Label
Attribute WeekdayLabel.VB_VarHelpID = -1

Private ParentForm As DP_frmDatePicker


'----------------------------------------
' Setup
'----------------------------------------

Public Sub Setup(ByVal Form As DP_frmDatePicker, _
                 ByVal WeekdayIndex As Long)

    Set ParentForm = Form

    Set WeekdayLabel = ParentForm.Controls.Add("Forms.Label.1")

    With WeekdayLabel
        .Caption = Format(DateSerial(2029, 1, 1 + WeekdayIndex), "ddd")
        .Left = DP_GridLeft(WeekdayIndex, _
                            DP_GRID_LEFT, _
                            DP_GRID_CELL_WIDTH, _
                            DP_CALENDAR_CELL_WIDTH)
        .ForeColor = DP_ColorTextHeader()
        .Font.Bold = True
        .Font.Size = DP_FONT_SIZE_WEEKDAY * DP_GetFontScale(ParentForm.InsideWidth)
        .TextAlign = fmTextAlignCenter
        .AutoSize = True

        .Top = DP_VerticalTextTop(DP_GridTop(0, DP_GRID_TOP, DP_GRID_CELL_HEIGHT), _
                                  DP_CALENDAR_CELL_HEIGHT, _
                                  .Height)

        .AutoSize = False
        .Width = DP_CALENDAR_CELL_WIDTH
        .Height = DP_CALENDAR_CELL_HEIGHT
    End With

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Private Sub WeekdayLabel_MouseMove(ByVal Button As Integer, _
                                   ByVal Shift As Integer, _
                                   ByVal X As Single, _
                                   ByVal Y As Single)

    ParentForm.ResetAllHover

End Sub

