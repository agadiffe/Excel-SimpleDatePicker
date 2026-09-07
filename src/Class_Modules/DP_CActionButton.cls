VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CActionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' References
'----------------------------------------

Private WithEvents ButtonLabel As MSForms.Label
Attribute ButtonLabel.VB_VarHelpID = -1
Private WithEvents ButtonTextLabel As MSForms.Label
Attribute ButtonTextLabel.VB_VarHelpID = -1

Private ParentForm As Object
Private ActionKey As String

Private HoverHandler As DP_CHoverHandler


'----------------------------------------
' Setup
'----------------------------------------

Public Sub Setup(ByVal Form As Object, _
                 ByVal Action As String)

    Dim ButtonWidth As Single

    Set ParentForm = Form
    ActionKey = Action

    Set ButtonLabel = ParentForm.Controls.Add("Forms.Label.1")

    With ButtonLabel
        .Height = DP_ACTION_BUTTON_HEIGHT
        .Top = ParentForm.InsideHeight - .Height - DP_ACTION_BUTTON_BOTTOM_MARGIN
    End With

    Set ButtonTextLabel = ParentForm.Controls.Add("Forms.Label.1")

    With ButtonTextLabel
        .Caption = DP_GetDatePickerActionText(ActionKey)
        .Font.Size = DP_FONT_SIZE_ACTION_BUTTON * DP_GetFontScale(ParentForm.InsideWidth)
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .ForeColor = DP_ColorTextSecondary()
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
        .AutoSize = True

        .Top = DP_VerticalTextTop(ButtonLabel.Top, _
                                  ButtonLabel.Height, _
                                  .Height)
    End With

    ButtonWidth = ButtonTextLabel.Width + DP_DATEPICKER_WIDTH * 0.1

    With ButtonLabel
        .Width = ButtonWidth
        .Left = DP_DATEPICKER_WIDTH - DP_ACTION_BUTTON_RIGHT_MARGIN - .Width
    End With

    With ButtonTextLabel
        .AutoSize = False
        .Left = ButtonLabel.Left
        .Width = ButtonLabel.Width
        .Height = ButtonLabel.Top + ButtonLabel.Height - .Top
    End With

    Set HoverHandler = New DP_CHoverHandler
    HoverHandler.Setup ButtonLabel

End Sub


'----------------------------------------
' Button action
'----------------------------------------

Private Sub ButtonLabel_Click()

    ExecuteAction

End Sub


Private Sub ButtonTextLabel_Click()

    ExecuteAction

End Sub


Private Sub ExecuteAction()

    Select Case ActionKey
        Case "GO_TO_TODAY"
            ParentForm.GoToToday
        Case "GO_TO_CURRENT_MONTH"
            ParentForm.GoToCurrentMonth
        Case "GO_TO_CURRENT_YEAR"
            ParentForm.GoToCurrentYear
    End Select

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Private Sub ButtonLabel_MouseMove(ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)

    SetHover

End Sub


Private Sub ButtonTextLabel_MouseMove(ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)

    SetHover

End Sub


Private Sub SetHover()

    ParentForm.ResetAllHover
    HoverHandler.SetHover

End Sub


Public Sub ResetHover()

    HoverHandler.ResetHover

End Sub

