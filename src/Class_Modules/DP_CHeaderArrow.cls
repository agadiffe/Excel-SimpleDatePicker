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

Private WithEvents ArrowBackground As MSForms.Image
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

    Set ArrowBackground = ParentView.ParentForm.Controls.Add("Forms.Image.1")

    With ArrowBackground
        .Left = LeftPosition
        .Top = DP_ARROW_TOP
        .Width = DP_ARROW_WIDTH
        .Height = DP_ARROW_HEIGHT
        .BorderStyle = fmBorderStyleNone
        .PictureSizeMode = fmPictureSizeModeStretch
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

    Dim BackgroundColor As Long

    ArrowText.ForeColor = DP_ColorTextHeader()

    If Not ParentView.IsArrowEnabled(Direction) Then
        BackgroundColor = DP_ColorBg()
        ArrowText.ForeColor = DP_ColorTextDisabled()
    ElseIf IsPressed Then
        BackgroundColor = DP_ColorBgPressed()
    ElseIf IsHovered Then
        BackgroundColor = DP_ColorBgHover()
    Else
        BackgroundColor = DP_ColorBg()
    End If

    ArrowBackground.BackColor = BackgroundColor

    #If Mac Then
        ' No GDI+ on Mac.
        ArrowBackground.Picture = Nothing
    #Else
        If DP_BUTTON_NO_SHAPE Then
            ' Plain Image used as a colored background.
            ArrowBackground.Picture = Nothing
        Else
            ' Render the actual button shape.
            Set ArrowBackground.Picture = _
                DP_GetCachedButtonPicture(DP_BUTTON_TYPE_ARROW, _
                                          BackgroundColor, _
                                          DP_ColorBg())

            ' MSForms Image/Label overlap bug: reset Z-order.
            ArrowBackground.ZOrder 1
            ArrowText.ZOrder 0
        End If
    #End If

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

Private Sub HandleMouseDown(ByVal Button As Integer)

    If Button <> 1 Then Exit Sub
    If Not ParentView.IsArrowEnabled(Direction) Then Exit Sub

    Pressed = True

End Sub


Private Sub ArrowBackground_MouseDown(ByVal Button As Integer, _
                                      ByVal Shift As Integer, _
                                      ByVal X As Single, _
                                      ByVal Y As Single)

    HandleMouseDown Button

End Sub


Private Sub ArrowText_MouseDown(ByVal Button As Integer, _
                                ByVal Shift As Integer, _
                                ByVal X As Single, _
                                ByVal Y As Single)

    HandleMouseDown Button

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

Public Sub SetHovered(ByVal Value As Boolean)

    Hovered = Value

End Sub


Private Sub HandleMouseHover()

    If ParentView.IsArrowEnabled(Direction) Then
        ParentView.SetHoveredButton Me
    Else
        ParentView.ClearHoveredButton
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

