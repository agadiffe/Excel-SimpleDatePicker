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
        .Top = ARROW_TOP
        .Width = ARROW_WIDTH
        .Height = ARROW_HEIGHT
    End With

    Set ArrowTextLabel = ParentForm.Controls.Add("Forms.Label.1")

    With ArrowTextLabel
        .Caption = ArrowCaption
        .Font.Size = FONT_SIZE_ARROW * GetFontScale(ParentForm.Width)
        .TextAlign = fmTextAlignCenter
        .ForeColor = ColorTextHeader()
        .BackStyle = fmBackStyleTransparent
        .AutoSize = True

        .Top = VerticalTextTop(ArrowLabel.Top, _
                               ArrowLabel.Height, _
                               .Height)

        .AutoSize = False
        .Left = ArrowLabel.Left
        .Width = ArrowLabel.Width
        .Height = ArrowLabel.Height
    End With

    Set HoverHandler = New DP_CHoverHandler
    HoverHandler.Setup ArrowLabel

End Sub


'----------------------------------------
' Arrow
'----------------------------------------

Private Sub ArrowLabel_Click()

    ParentForm.ArrowClicked ActionKey

End Sub


Private Sub ArrowTextLabel_Click()

    ParentForm.ArrowClicked ActionKey

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
    HoverHandler.SetHover

End Sub


Public Sub ResetHover()

    HoverHandler.ResetHover

End Sub

