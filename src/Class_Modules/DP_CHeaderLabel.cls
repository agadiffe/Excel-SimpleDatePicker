VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CHeaderLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' References
'----------------------------------------

Private WithEvents HeaderLabel As MSForms.Label
Attribute HeaderLabel.VB_VarHelpID = -1
Private WithEvents HeaderTextLabel As MSForms.Label
Attribute HeaderTextLabel.VB_VarHelpID = -1

Private ParentForm As Object
Private Interactive As Boolean

Private HoverHandler As DP_CHoverHandler


'----------------------------------------
' Setup
'----------------------------------------

Public Sub Setup(ByVal Form As Object, _
                 ByVal HeaderCaption As String)

    Set ParentForm = Form
    Interactive = True

    Set HeaderLabel = ParentForm.Controls.Add("Forms.Label.1")

    With HeaderLabel
        .Left = DP_HEADER_LEFT
        .Top = DP_HEADER_TOP
        .Width = DP_HEADER_WIDTH
        .Height = DP_HEADER_HEIGHT
    End With

    Set HeaderTextLabel = ParentForm.Controls.Add("Forms.Label.1")

    With HeaderTextLabel
        .Caption = HeaderCaption
        .Font.Size = DP_FONT_SIZE_HEADER * DP_GetFontScale(ParentForm.Width)
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .ForeColor = DP_ColorTextHeader()
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
        .AutoSize = True

        .Top = DP_VerticalTextTop(HeaderLabel.Top, _
                                  HeaderLabel.Height, _
                                  .Height)

        .AutoSize = False
        .Left = HeaderLabel.Left
        .Width = HeaderLabel.Width
        .Height = HeaderLabel.Top + HeaderLabel.Height - .Top
    End With

    Set HoverHandler = New DP_CHoverHandler
    HoverHandler.Setup HeaderLabel

End Sub


Public Sub SetCaption(ByVal Caption As String)

    HeaderTextLabel.Caption = Caption

End Sub


Public Sub SetInteractive(ByVal Value As Boolean)

    Interactive = Value

End Sub


'----------------------------------------
' Header
'----------------------------------------

Private Sub HeaderLabel_Click()

    HandleHeaderClick

End Sub


Private Sub HeaderTextLabel_Click()

    HandleHeaderClick

End Sub


Private Sub HandleHeaderClick()

    If Not Interactive Then Exit Sub

    ParentForm.HeaderClicked

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Private Sub HeaderLabel_MouseMove(ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)

    SetHover

End Sub


Private Sub HeaderTextLabel_MouseMove(ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)

    SetHover

End Sub


Private Sub SetHover()

    ParentForm.ResetAllHover

    If Interactive Then
        HoverHandler.SetHover
    End If

End Sub


Public Sub ResetHover()

    HoverHandler.ResetHover

End Sub

