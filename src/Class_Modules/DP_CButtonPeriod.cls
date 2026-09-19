VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CButtonPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Private WithEvents PeriodBackground As MSForms.Image
Attribute PeriodBackground.VB_VarHelpID = -1
Private WithEvents PeriodText As MSForms.Label
Attribute PeriodText.VB_VarHelpID = -1

Private ParentView As DP_IPickerView
Private CellDate As Date
Private ButtonType As DP_ButtonType

Private IsHovered As Boolean
Private IsPressed As Boolean


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub Initialize(ByVal View As DP_IPickerView, _
                      ByVal ButtonTypeValue As DP_ButtonType, _
                      ByVal CellDateValue As Date, _
                      ByVal Caption As String, _
                      ByVal LeftPosition As Single, _
                      ByVal TopPosition As Single, _
                      ByVal CellWidth As Single)

    Set ParentView = View
    ButtonType = ButtonTypeValue
    CellDate = CellDateValue

    Set PeriodBackground = ParentView.ParentForm.Controls.Add("Forms.Image.1")

    With PeriodBackground
        .Left = LeftPosition
        .Top = TopPosition
        .Width = CellWidth
        .Height = DP_PERIOD_CELL_HEIGHT
        .BorderStyle = fmBorderStyleNone
        .PictureSizeMode = fmPictureSizeModeStretch
    End With

    Set PeriodText = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With PeriodText
        .Caption = Caption
        .Font.Size = DP_FONT_SIZE_PERIOD * DP_GetFontScale(DP_DATEPICKER_WIDTH)
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False

        .AutoSize = True
        .Top = DP_VerticalTextTop(PeriodBackground.Top, _
                                  PeriodBackground.Height, _
                                  .Height)
        .AutoSize = False

        .Left = PeriodBackground.Left
        .Width = PeriodBackground.Width
        .Height = PeriodBackground.Top + PeriodBackground.Height - .Top
    End With

    Render

End Sub


'----------------------------------------
' Render
'----------------------------------------

Public Sub UpdatePeriod(ByVal CellDateValue As Date, _
                        ByVal Caption As String)

    CellDate = CellDateValue
    PeriodText.Caption = Caption

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

    ' Other period
    '--------------------

    If ParentView.IsSecondaryPeriod(CellDate) Then
        TextColor = DP_ColorTextSecondary()
    End If

    ' Disabled
    '--------------------

    If DP_IsDateOutsideRange(CellDate) Then
        TextColor = DP_ColorTextDisabled()
    End If

    ' Current
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

    PeriodBackground.BackColor = BackgroundColor

    #If Mac Then
        ' No GDI+ on Mac.
        PeriodBackground.Picture = Nothing
    #Else
        If DP_BUTTON_NO_SHAPE Then
            PeriodBackground.Picture = Nothing
        Else
            ' Render the actual button shape.
            Set PeriodBackground.Picture = _
                DP_GetCachedButtonPicture(ButtonType, _
                                          BackgroundColor, _
                                          DP_ColorBg())

            ' MSForms Image/Label overlap bug: reset Z-order.
            PeriodBackground.ZOrder 1
            PeriodText.ZOrder 0
        End If
    #End If

    PeriodText.ForeColor = TextColor
    PeriodText.Font.Bold = IsBold

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

Private Sub HandlePeriodClick()

    If DP_IsDateOutsideRange(CellDate) Then Exit Sub

    ParentView.GridButtonClicked CellDate

End Sub


Private Sub PeriodBackground_Click()

    HandlePeriodClick

End Sub


Private Sub PeriodText_Click()

    HandlePeriodClick

End Sub


'----------------------------------------
' Pressed
'----------------------------------------

' Mouse down
'--------------------

Private Sub HandleMouseDown(ByVal Button As Integer)

    If Button <> 1 Then Exit Sub
    If DP_IsDateOutsideRange(CellDate) Then Exit Sub

    Pressed = True

End Sub


Private Sub PeriodBackground_MouseDown(ByVal Button As Integer, _
                                       ByVal Shift As Integer, _
                                       ByVal X As Single, _
                                       ByVal Y As Single)

    HandleMouseDown Button

End Sub


Private Sub PeriodText_MouseDown(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    HandleMouseDown Button

End Sub


' Mouse up
'--------------------

Private Sub PeriodBackground_MouseUp(ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single)

    Pressed = False

End Sub


Private Sub PeriodText_MouseUp(ByVal Button As Integer, _
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

    If DP_IsDateOutsideRange(CellDate) Then
        ParentView.ClearHoveredButton
    Else
        ParentView.SetHoveredButton Me
    End If

End Sub


Private Sub PeriodBackground_MouseMove(ByVal Button As Integer, _
                                       ByVal Shift As Integer, _
                                       ByVal X As Single, _
                                       ByVal Y As Single)

    HandleMouseHover

End Sub


Private Sub PeriodText_MouseMove(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    HandleMouseHover

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    PeriodBackground.Visible = IsVisible
    PeriodText.Visible = IsVisible

End Sub

