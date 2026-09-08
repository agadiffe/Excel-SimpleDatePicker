VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CArrowLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' References
'----------------------------------------

Private WithEvents ArrowLabel As MSForms.Label
Attribute ArrowLabel.VB_VarHelpID = -1
Private WithEvents ArrowTextLabel As MSForms.Label
Attribute ArrowTextLabel.VB_VarHelpID = -1

Private ParentForm As Object
Private ActionKey As String

Private HoverHandler As DP_CHoverHandler


'----------------------------------------
' Setup
'----------------------------------------

Public Sub Setup(ByVal Form As Object, _
                 ByVal Action As String, _
                 ByVal ArrowCaption As String, _
                 ByVal LeftPosition As Single)

    Set ParentForm = Form
    ActionKey = Action

    Set ArrowLabel = ParentForm.Controls.Add("Forms.Label.1")

    With ArrowLabel
        .Left = LeftPosition
        .Top = DP_ARROW_TOP
        .Width = DP_ARROW_WIDTH
        .Height = DP_ARROW_HEIGHT
    End With

    Set ArrowTextLabel = ParentForm.Controls.Add("Forms.Label.1")

    With ArrowTextLabel
        .Caption = ArrowCaption
        .Font.Size = DP_FONT_SIZE_ARROW * DP_GetFontScale(ParentForm.InsideWidth)
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .AutoSize = True

        .Top = DP_VerticalTextTop(ArrowLabel.Top, _
                                  ArrowLabel.Height, _
                                  .Height)

        .AutoSize = False
        .Left = ArrowLabel.Left
        .Width = ArrowLabel.Width
        .Height = ArrowLabel.Top + ArrowLabel.Height - .Top
    End With

    Set HoverHandler = New DP_CHoverHandler
    HoverHandler.Setup ArrowLabel

    SetState ParentForm.IsArrowEnabled(ActionKey)

End Sub


'----------------------------------------
' Arrow
'----------------------------------------

Public Sub SetState(ByVal IsEnabled As Boolean)

    If IsEnabled Then
        ArrowTextLabel.ForeColor = DP_ColorTextHeader()
    Else
        ArrowTextLabel.ForeColor = DP_ColorTextDisabled()
    End If

End Sub


Private Sub ArrowLabel_Click()

    HandleArrowClick

End Sub


Private Sub ArrowTextLabel_Click()

    HandleArrowClick

End Sub


Private Sub HandleArrowClick()

    If Not ParentForm.IsArrowEnabled(ActionKey) Then Exit Sub

    ParentForm.ArrowClicked ActionKey

End Sub


Private Sub ArrowLabel_MouseDown(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    HandleArrowMouseDown

End Sub


Private Sub ArrowTextLabel_MouseDown(ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single)

    HandleArrowMouseDown

End Sub


Private Sub HandleArrowMouseDown()

    If Not ParentForm.IsArrowEnabled(ActionKey) Then Exit Sub

    ArrowLabel.BackColor = DP_ColorBgArrowPressed()

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Private Sub ArrowLabel_MouseMove(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    SetHover

End Sub


Private Sub ArrowTextLabel_MouseMove(ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single)

    SetHover

End Sub


Private Sub SetHover()

    ParentForm.ResetAllHover

    If ParentForm.IsArrowEnabled(ActionKey) Then
        HoverHandler.SetHover
    End If

End Sub


Public Sub ResetHover()

    HoverHandler.ResetHover

End Sub

