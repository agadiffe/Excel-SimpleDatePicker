VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CButtonDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Private WithEvents DayBackground As MSForms.Image
Attribute DayBackground.VB_VarHelpID = -1
Private WithEvents DayText As MSForms.Label
Attribute DayText.VB_VarHelpID = -1

Private ParentView As DP_IPickerView
Private CellDate As Date

Private IsHovered As Boolean
Private IsPressed As Boolean
Private MouseDownInside As Boolean


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub Initialize(ByVal View As DP_CViewCalendar, _
                      ByVal CellDateValue As Date, _
                      ByVal Caption As String, _
                      ByVal LeftPosition As Single, _
                      ByVal TopPosition As Single)

    Dim HorizontalCorrection As Single

    Set ParentView = View
    CellDate = CellDateValue

    Set DayBackground = ParentView.ParentForm.Controls.Add("Forms.Image.1")

    With DayBackground
        .Left = LeftPosition
        .Top = TopPosition
        .Width = DP_CALENDAR_CELL_WIDTH
        .Height = DP_CALENDAR_CELL_HEIGHT
        .BorderStyle = fmBorderStyleNone
        .PictureSizeMode = fmPictureSizeModeStretch
    End With

    HorizontalCorrection = DayBackground.Width * 0.03

    Set DayText = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With DayText
        .Caption = Caption
        .Font.Size = DP_FONT_SIZE_DAY * DP_GetFontScale(DP_DATEPICKER_WIDTH)
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
        .WordWrap = False

        .AutoSize = True
        .Top = DP_VerticalTextTop(DayBackground.Top, _
                                  DayBackground.Height, _
                                  .Height)
        .AutoSize = False

        .Left = DayBackground.Left + HorizontalCorrection
        .Width = DayBackground.Width - HorizontalCorrection
        .Height = DayBackground.Top + DayBackground.Height - .Top
    End With

    Render

End Sub


'----------------------------------------
' Render
'----------------------------------------

Public Sub UpdateDate(ByVal CellDateValue As Date, _
                      ByVal Caption As String)

    CellDate = CellDateValue
    DayText.Caption = Caption

    Render

End Sub


Private Sub Render()

    Dim BackgroundColor As Long
    Dim TextColor As Long
    Dim IsBold As Boolean

    ' Default
    '--------------------

    BackgroundColor = DP_ColorBg()
    TextColor = DP_ColorText()
    IsBold = False

    ' Other month
    '--------------------

    If ParentView.IsSecondaryPeriod(CellDate) Then
        TextColor = DP_ColorTextSecondary()
    End If

    ' Disabled
    '--------------------

    If DP_IsDateOutsideRange(CellDate) Then
        TextColor = DP_ColorTextDisabled()
    End If

    ' Today
    '--------------------

    If ParentView.IsCurrentPeriod(CellDate) Then
        BackgroundColor = DP_ColorBgToday()
        TextColor = DP_ColorTextToday()
    End If

    ' Selected
    '--------------------

    If ParentView.IsDateSelected(CellDate) Then
        BackgroundColor = DP_ColorBgSelected()
        TextColor = DP_ColorTextSelected()
        IsBold = True
    End If

    ' Hover
    '--------------------

    If IsHovered Then
        If ParentView.IsDateSelected(CellDate) Then
            BackgroundColor = DP_ColorBgSelectedHover()
        Else
            BackgroundColor = DP_ColorBgHover()
        End If
    End If

    ' Pressed
    '--------------------

    If IsPressed Then
        If ParentView.IsDateSelected(CellDate) Then
            BackgroundColor = DP_ColorBgSelectedPressed()
        Else
            BackgroundColor = DP_ColorBgPressed()
        End If
    End If

    ' Apply
    '--------------------

    DayBackground.BackColor = BackgroundColor

    If DP_BUTTON_NO_SHAPE Then
        DayBackground.Picture = Nothing
    Else
        ' Render the actual button shape.
        Set DayBackground.Picture = _
            DP_GetCachedButtonPicture(DP_BUTTON_TYPE_DAY, _
                                      BackgroundColor, _
                                      DP_ColorBg())

        ' MSForms Image/Label overlap bug: reset Z-order.
        DayBackground.ZOrder 1
        DayText.ZOrder 0
    End If

    DayText.ForeColor = TextColor
    DayText.Font.Bold = IsBold

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
' Hit testing
'----------------------------------------

Private Function IsInsideButton(ByVal X As Single, _
                                ByVal Y As Single) As Boolean

    If DP_BUTTON_NO_SHAPE Or DP_DAY_BUTTON_SHAPE = DP_BUTTON_SHAPE_RECTANGLE Then
        IsInsideButton = True
        Exit Function
    End If

    ' Circle

    Dim CenterX As Single
    Dim CenterY As Single
    Dim Radius As Single

    CenterX = DayBackground.Width / 2
    CenterY = DayBackground.Height / 2

    Radius = DayBackground.Width

    If DayBackground.Height < Radius Then
        Radius = DayBackground.Height
    End If

    Radius = Radius / 2

    IsInsideButton = ((X - CenterX) ^ 2 + (Y - CenterY) ^ 2) <= Radius ^ 2

End Function


'----------------------------------------
' Click
'----------------------------------------

Private Sub HandleDayClick()

    If MouseDownInside Then
        If Not DP_IsDateOutsideRange(CellDate) Then
            ParentView.GridButtonClicked CellDate
        End If
    End If

    MouseDownInside = False

End Sub


Private Sub DayBackground_Click()

    HandleDayClick

End Sub


Private Sub DayText_Click()

    HandleDayClick

End Sub


'----------------------------------------
' Pressed
'----------------------------------------

' Mouse down
'--------------------

Private Sub HandleMouseDown(ByVal Button As Integer, _
                            ByVal X As Single, _
                            ByVal Y As Single)

    If Button <> 1 Then Exit Sub
    If Not IsInsideButton(X, Y) Then Exit Sub
    If DP_IsDateOutsideRange(CellDate) Then Exit Sub

    Pressed = True

End Sub


Private Sub DayBackground_MouseDown(ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)

    HandleMouseDown Button, X, Y

End Sub


Private Sub DayText_MouseDown(ByVal Button As Integer, _
                              ByVal Shift As Integer, _
                              ByVal X As Single, _
                              ByVal Y As Single)

    HandleMouseDown Button, _
                    DayText.Left + X - DayBackground.Left, _
                    DayText.Top + Y - DayBackground.Top

End Sub


' Mouse up
'--------------------

Private Sub HandleMouseUp(ByVal X As Single, _
                          ByVal Y As Single)

    MouseDownInside = IsInsideButton(X, Y)
    Pressed = False

End Sub


Private Sub DayBackground_MouseUp(ByVal Button As Integer, _
                                  ByVal Shift As Integer, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)

    HandleMouseUp X, Y

End Sub


Private Sub DayText_MouseUp(ByVal Button As Integer, _
                            ByVal Shift As Integer, _
                            ByVal X As Single, _
                            ByVal Y As Single)

    HandleMouseUp DayText.Left + X - DayBackground.Left, _
                  DayText.Top + Y - DayBackground.Top

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub SetHovered(ByVal Value As Boolean)

    Hovered = Value

End Sub


Private Sub HandleMouseHover(ByVal X As Single, _
                             ByVal Y As Single)

    If DP_IsDateOutsideRange(CellDate) Or Not IsInsideButton(X, Y) Then
        ParentView.ClearHoveredButton
    Else
        ParentView.SetHoveredButton Me
    End If

End Sub


Private Sub DayBackground_MouseMove(ByVal Button As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal X As Single, _
                                    ByVal Y As Single)

    HandleMouseHover X, Y

End Sub


Private Sub DayText_MouseMove(ByVal Button As Integer, _
                              ByVal Shift As Integer, _
                              ByVal X As Single, _
                              ByVal Y As Single)

    HandleMouseHover DayText.Left + X - DayBackground.Left, _
                     DayText.Top + Y - DayBackground.Top

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    DayBackground.Visible = IsVisible
    DayText.Visible = IsVisible

End Sub

