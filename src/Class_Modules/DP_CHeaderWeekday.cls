VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CHeaderWeekday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Private WithEvents WeekdayBackground As MSForms.Label
Attribute WeekdayBackground.VB_VarHelpID = -1
Private WithEvents WeekdayText As MSForms.Label
Attribute WeekdayText.VB_VarHelpID = -1

Private ParentView As DP_IPickerView


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub Initialize(ByVal View As DP_CViewCalendar, _
                      ByVal Caption As String, _
                      ByVal LeftPosition As Single, _
                      ByVal TopPosition As Single)

    Set ParentView = View

    Set WeekdayBackground = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With WeekdayBackground
        .Left = LeftPosition
        .Top = TopPosition
        .Width = DP_CALENDAR_CELL_WIDTH
        .Height = DP_CALENDAR_CELL_HEIGHT
    End With

    Set WeekdayText = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With WeekdayText
        .Caption = Caption
        .Font.Size = DP_FONT_SIZE_WEEKDAY * DP_GetFontScale(DP_DATEPICKER_WIDTH)
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False

        .AutoSize = True
        .Top = DP_VerticalTextTop(WeekdayBackground.Top, _
                                  WeekdayBackground.Height, _
                                  .Height)
        .AutoSize = False

        .Left = WeekdayBackground.Left
        .Width = WeekdayBackground.Width
        .Height = WeekdayBackground.Top + WeekdayBackground.Height - .Top
    End With

    Render

End Sub


'----------------------------------------
' Render
'----------------------------------------

Public Sub RefreshTheme()

    Render

End Sub


Private Sub Render()

    WeekdayBackground.BackColor = DP_ColorBg()
    WeekdayText.ForeColor = DP_ColorTextHeader()

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Private Sub WeekdayBackground_MouseMove(ByVal Button As Integer, _
                                        ByVal Shift As Integer, _
                                        ByVal X As Single, _
                                        ByVal Y As Single)

    ParentView.ClearHoveredButton

End Sub


Private Sub WeekdayText_MouseMove(ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)

    ParentView.ClearHoveredButton

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    WeekdayBackground.Visible = IsVisible
    WeekdayText.Visible = IsVisible

End Sub


