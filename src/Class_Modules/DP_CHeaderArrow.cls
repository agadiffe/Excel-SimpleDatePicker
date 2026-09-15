VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CHeaderArrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Private WithEvents ArrowBackground As MSForms.Label
Attribute ArrowBackground.VB_VarHelpID = -1
Private WithEvents ArrowText As MSForms.Label
Attribute ArrowText.VB_VarHelpID = -1

Private ParentView As DP_IPickerView
Private Direction As DP_NavigationDirection

Private IsHovered As Boolean
Private IsPressed As Boolean


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub Initialize(ByVal View As DP_IPickerView, _
                      ByVal Caption As String, _
                      ByVal LeftPosition As Single, _
                      ByVal DirectionValue As DP_NavigationDirection)

    Set ParentView = View
    Direction = DirectionValue

    Set ArrowBackground = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With ArrowBackground
        .Left = LeftPosition
        .Top = DP_ARROW_TOP
        .Width = DP_ARROW_WIDTH
        .Height = DP_ARROW_HEIGHT
    End With

    Set ArrowText = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With ArrowText
        .Caption = Caption
        .Font.Size = DP_FONT_SIZE_ARROW * DP_GetFontScale(DP_DATEPICKER_WIDTH)
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent

        .AutoSize = True
        .Top = DP_VerticalTextTop(ArrowBackground.Top, _
                                  ArrowBackground.Height, _
                                  .Height)
        .AutoSize = False

        .Left = ArrowBackground.Left
        .Width = ArrowBackground.Width
        .Height = ArrowBackground.Top + ArrowBackground.Height - .Top
    End With

    Render

End Sub


'----------------------------------------
' Render
'----------------------------------------

Public Sub UpdateState()

    Render

End Sub


Public Sub RefreshTheme()

    Render

End Sub


Private Sub Render()

    If Not ParentView.IsArrowEnabled(Direction) Then
        ArrowBackground.BackColor = DP_ColorBg()
        ArrowText.ForeColor = DP_ColorTextDisabled()
    ElseIf IsPressed Then
        ArrowBackground.BackColor = DP_ColorBgPressed()
        ArrowText.ForeColor = DP_ColorTextHeader()
    ElseIf IsHovered Then
        ArrowBackground.BackColor = DP_ColorBgHover()
        ArrowText.ForeColor = DP_ColorTextHeader()
    Else
        ArrowBackground.BackColor = DP_ColorBg()
        ArrowText.ForeColor = DP_ColorTextHeader()
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

Private Sub HandleClick()

    If Not ParentView.IsArrowEnabled(Direction) Then Exit Sub

    ParentView.ArrowClicked Direction

End Sub


Private Sub ArrowBackground_Click()

    HandleClick

End Sub


Private Sub ArrowText_Click()

    HandleClick

End Sub


'----------------------------------------
' Pressed
'----------------------------------------

' Mouse down
'--------------------

Private Sub HandleMouseDown()

    If Not ParentView.IsArrowEnabled(Direction) Then Exit Sub

    Pressed = True

End Sub


Private Sub ArrowBackground_MouseDown(ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)

    HandleMouseDown

End Sub


Private Sub ArrowText_MouseDown(ByVal Button As Integer, _
                                ByVal Shift As Integer, _
                                ByVal X As Single, _
                                ByVal Y As Single)

    HandleMouseDown

End Sub


' Mouse up
'--------------------

Private Sub ArrowBackground_MouseUp(ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)

    Pressed = False

End Sub


Private Sub ArrowText_MouseUp(ByVal Button As Integer, _
                              ByVal Shift As Integer, _
                              ByVal X As Single, _
                              ByVal Y As Single)

    Pressed = False

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub ResetHover()

    Hovered = False

End Sub


Private Sub HandleMouseHover()

    ParentView.ResetAllHover

    If ParentView.IsArrowEnabled(Direction) Then
        Hovered = True
    End If

End Sub


Private Sub ArrowBackground_MouseMove(ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)

    HandleMouseHover

End Sub


Private Sub ArrowText_MouseMove(ByVal Button As Integer, _
                                ByVal Shift As Integer, _
                                ByVal X As Single, _
                                ByVal Y As Single)

    HandleMouseHover

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    ArrowBackground.Visible = IsVisible
    ArrowText.Visible = IsVisible

End Sub

