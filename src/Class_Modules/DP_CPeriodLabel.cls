VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CPeriodLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' References
'----------------------------------------

Private WithEvents PeriodLabel As MSForms.Label
Attribute PeriodLabel.VB_VarHelpID = -1
Private WithEvents PeriodTextLabel As MSForms.Label
Attribute PeriodTextLabel.VB_VarHelpID = -1

Private ParentForm As Object
Private PeriodKind As String
Private PeriodValue As Long
Private PeriodDate As Date

Private HoverHandler As DP_CHoverHandler


'----------------------------------------
' Setup
'----------------------------------------

Public Sub Setup(ByVal Form As Object, _
                 ByVal Kind As String, _
                 ByVal Value As Long, _
                 ByVal CellDate As Date, _
                 ByVal LeftPosition As Single, _
                 ByVal TopPosition As Single, _
                 ByVal CellWidth As Single)

    Set ParentForm = Form

    PeriodKind = Kind
    PeriodValue = Value
    PeriodDate = CellDate

    Set PeriodLabel = ParentForm.Controls.Add("Forms.Label.1")

    With PeriodLabel
        .Left = LeftPosition
        .Top = TopPosition
        .Width = CellWidth
        .Height = PERIOD_CELL_HEIGHT
    End With

    Set PeriodTextLabel = ParentForm.Controls.Add("Forms.Label.1")

    With PeriodTextLabel
        .Caption = GetPeriodCaption()
        .Font.Size = FONT_SIZE_PERIOD * GetFontScale(ParentForm.Width)
        .TextAlign = fmTextAlignCenter
        .ForeColor = ColorText()
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
        .AutoSize = True

        .Top = VerticalTextTop(PeriodLabel.Top, _
                               PeriodLabel.Height, _
                               .Height)

        .AutoSize = False
        .Left = PeriodLabel.Left
        .Width = PeriodLabel.Width
        .Height = PeriodLabel.Height
    End With

    Set HoverHandler = New DP_CHoverHandler
    HoverHandler.Setup PeriodTextLabel, PeriodLabel

End Sub


'----------------------------------------
' Period
'----------------------------------------

Public Sub UpdatePeriod(ByVal CellDate As Date, _
                        ByVal Value As Long, _
                        ByVal IsCurrent As Boolean, _
                        ByVal IsSelected As Boolean)

    PeriodDate = CellDate
    PeriodValue = Value

    PeriodTextLabel.Caption = GetPeriodCaption()

    SetState IsCurrent, IsSelected

End Sub


Private Function GetPeriodCaption() As String

    Select Case PeriodKind
        Case "MONTH"
            If MONTH_GRID_COLUMNS > 3 Then
                GetPeriodCaption = Format(PeriodDate, "mmm")
            Else
                GetPeriodCaption = Format(PeriodDate, "mmmm")
            End If
        Case "YEAR"
            GetPeriodCaption = Format(PeriodDate, "yyyy")
    End Select

End Function


'----------------------------------------
' State
'----------------------------------------

Public Sub SetState(ByVal IsCurrent As Boolean, _
                    ByVal IsSelected As Boolean)

    ' Default state
    '--------------------

    PeriodLabel.BackColor = ColorBg()
    PeriodTextLabel.ForeColor = ColorText()
    PeriodTextLabel.Font.Bold = False

    ' Other year
    '--------------------

    If PeriodKind = "MONTH" And _
           Year(PeriodDate) <> ParentForm.GetCurrentYear Then
        PeriodTextLabel.ForeColor = ColorTextSecondary()
    End If

    ' Current
    '--------------------

    If IsCurrent Then
        PeriodLabel.BackColor = ColorBgToday()
        PeriodTextLabel.ForeColor = ColorTextToday()
        PeriodTextLabel.Font.Bold = True
    End If

    ' Selected
    '--------------------

    If IsSelected Then
        PeriodLabel.BackColor = ColorBgSelected()
        PeriodTextLabel.ForeColor = ColorTextSelected()
        PeriodTextLabel.Font.Bold = True
    End If

End Sub


'----------------------------------------
' Period click
'----------------------------------------

Private Sub PeriodLabel_Click()

    HandlePeriodClick

End Sub


Private Sub PeriodTextLabel_Click()

    HandlePeriodClick

End Sub


Private Sub HandlePeriodClick()

    Select Case PeriodKind
        Case "MONTH"
            ParentForm.MonthLabelClicked PeriodValue
        Case "YEAR"
            ParentForm.YearLabelClicked PeriodValue
    End Select

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Private Sub PeriodLabel_MouseMove(ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)

    SetHover

End Sub


Private Sub PeriodTextLabel_MouseMove(ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)

    SetHover

End Sub


Private Sub SetHover()

    ParentForm.ResetArrowHover

    Select Case PeriodKind
        Case "MONTH"
            ParentForm.ResetMonthHoverExcept Me
            ParentForm.ResetCurrentMonthHover
            ParentForm.ResetHeaderHover
        Case "YEAR"
            ParentForm.ResetYearHoverExcept Me
            ParentForm.ResetCurrentYearHover
    End Select

    HoverHandler.SetBackgroundHover

End Sub


Public Sub ResetHover()

    HoverHandler.ResetHover

End Sub

