VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CButtonShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Private WithEvents ButtonBackground As MSForms.Image
Attribute ButtonBackground.VB_VarHelpID = -1
Private WithEvents ButtonText As MSForms.Label
Attribute ButtonText.VB_VarHelpID = -1

Private ParentView As DP_IPickerView

Private IsHovered As Boolean
Private IsPressed As Boolean


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub Initialize(ByVal View As DP_IPickerView, _
                      ByVal Caption As String, _
                      ByVal TopPosition As Single)

    Dim ButtonWidth As Single
    Set ParentView = View

    Set ButtonBackground = ParentView.ParentForm.Controls.Add("Forms.Image.1")

    With ButtonBackground
        .Height = DP_ACTION_BUTTON_HEIGHT
        .Top = TopPosition
        .BorderStyle = fmBorderStyleNone
        .PictureSizeMode = fmPictureSizeModeStretch
    End With

    Set ButtonText = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With ButtonText
        .Caption = Caption
        .Font.Size = DP_FONT_SIZE_ACTION_BUTTON * DP_GetFontScale(DP_DATEPICKER_WIDTH)
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False

        .AutoSize = True
        .Top = DP_VerticalTextTop(ButtonBackground.Top, _
                                  ButtonBackground.Height, _
                                  .Height)
    End With

    ButtonWidth = ButtonText.Width + DP_DATEPICKER_WIDTH * 0.07

    With ButtonBackground
        .Width = ButtonWidth
        .Left = DP_DATEPICKER_WIDTH - DP_ACTION_BUTTON_RIGHT_MARGIN - .Width
    End With

    With ButtonText
        .AutoSize = False
        .Left = ButtonBackground.Left
        .Width = ButtonBackground.Width
        .Height = ButtonBackground.Top + ButtonBackground.Height - .Top
    End With

    Render

End Sub


Public Sub Recreate(ByVal Caption As String, _
                    ByVal TopPosition As Single)

    ' Recreate because the button position depends on the
    ' active view's grid height.

    ButtonBackground.Parent.Controls.Remove ButtonBackground.Name
    ButtonText.Parent.Controls.Remove ButtonText.Name

    Set ButtonBackground = Nothing
    Set ButtonText = Nothing

    Initialize ParentView, Caption, TopPosition

End Sub


'----------------------------------------
' Render
'----------------------------------------

Private Sub Render()

    Dim BackgroundColor As Long

    If IsPressed Then
        BackgroundColor = DP_ColorBgPressed()
    ElseIf IsHovered Then
        BackgroundColor = DP_ColorBgHover()
    Else
        BackgroundColor = DP_ColorBg()
    End If

    ButtonBackground.BackColor = BackgroundColor
    ButtonText.ForeColor = DP_ColorTextSecondary()

    If DP_BUTTON_NO_SHAPE Then
        ' Plain Image used as a colored background.
        ButtonBackground.Picture = Nothing
    Else
        ' Render the actual button shape.
        Set ButtonBackground.Picture = _
            DP_GetCachedPicture(DP_PICTURE_TYPE_SHORTCUT, _
                                BackgroundColor, _
                                DP_ColorBg(), _
                                ButtonBackground.Width)

        ' MSForms Image/Label overlap bug: reset Z-order.
        ButtonBackground.ZOrder 1
        ButtonText.ZOrder 0
    End If

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

Private Sub ButtonBackground_Click()

    ParentView.ActionButtonClicked

End Sub


Private Sub ButtonText_Click()

    ParentView.ActionButtonClicked

End Sub


'----------------------------------------
' Pressed
'----------------------------------------

' Mouse down
'--------------------

Private Sub HandleMouseDown(ByVal Button As Integer)

    If Button <> 1 Then Exit Sub

    Pressed = True

End Sub


Private Sub ButtonBackground_MouseDown(ByVal Button As Integer, _
                                       ByVal Shift As Integer, _
                                       ByVal X As Single, _
                                       ByVal Y As Single)

    HandleMouseDown Button

End Sub


Private Sub ButtonText_MouseDown(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    HandleMouseDown Button

End Sub


' Mouse up
'--------------------

Private Sub ButtonBackground_MouseUp(ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single)

    Pressed = False

End Sub


Private Sub ButtonText_MouseUp(ByVal Button As Integer, _
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

    ParentView.SetHoveredButton Me

End Sub


Private Sub ButtonBackground_MouseMove(ByVal Button As Integer, _
                                       ByVal Shift As Integer, _
                                       ByVal X As Single, _
                                       ByVal Y As Single)

    HandleMouseHover

End Sub


Private Sub ButtonText_MouseMove(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    HandleMouseHover

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    ButtonBackground.Visible = IsVisible
    ButtonText.Visible = IsVisible

End Sub

