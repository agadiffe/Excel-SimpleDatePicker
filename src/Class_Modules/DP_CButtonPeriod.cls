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

Private WithEvents PeriodBackground As MSForms.Label
Attribute PeriodBackground.VB_VarHelpID = -1
Private WithEvents PeriodText As MSForms.Label
Attribute PeriodText.VB_VarHelpID = -1

Private ParentView As Object

Private PeriodValue As Long
Private PeriodDate As Date

Private IsHovered As Boolean
Private IsPressed As Boolean


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub Initialize(ByVal View As Object, _
                      ByVal Caption As String, _
                      ByVal Value As Long, _
                      ByVal CellDate As Date, _
                      ByVal LeftPosition As Single, _
                      ByVal TopPosition As Single, _
                      ByVal CellWidth As Single)

    Set ParentView = View

    PeriodDate = CellDate
    PeriodValue = Value

    Set PeriodBackground = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With PeriodBackground
        .Left = LeftPosition
        .Top = TopPosition
        .Width = CellWidth
        .Height = DP_PERIOD_CELL_HEIGHT
    End With

    Set PeriodText = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With PeriodText
        .Caption = Caption
        .Font.Size = DP_FONT_SIZE_PERIOD * _
                     DP_GetFontScale(DP_DATEPICKER_WIDTH)
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

    If ParentView.IsSecondaryPeriod(PeriodDate) Then
        TextColor = DP_ColorTextSecondary()
    End If

    ' Disabled
    '--------------------

    If ParentView.IsPeriodOutsideRange(PeriodDate) Then
        TextColor = DP_ColorTextDisabled()
    End If

    ' Current
    '--------------------

    If ParentView.IsCurrentPeriod(PeriodValue) Then
        BackgroundColor = DP_ColorBgToday()
        TextColor = DP_ColorTextToday()
    End If

    ' Selected
    '--------------------

    If ParentView.IsPeriodSelected(PeriodDate) Then
        BackgroundColor = DP_ColorBgSelected()
        TextColor = DP_ColorTextSelected()
        IsBold = True
    End If

    ' Hover
    '--------------------

    If IsHovered Then
        If ParentView.IsPeriodSelected(PeriodDate) Then
            BackgroundColor = DP_ColorBgSelectedHover()
        Else
            BackgroundColor = DP_ColorBgHover()
        End If
    End If

    ' Pressed
    '--------------------

    If IsPressed Then
        If ParentView.IsPeriodSelected(PeriodDate) Then
            BackgroundColor = DP_ColorBgSelectedPressed()
        Else
            BackgroundColor = DP_ColorBgPressed()
        End If
    End If

    ' Apply
    '--------------------

    PeriodBackground.BackColor = BackgroundColor
    PeriodText.ForeColor = TextColor
    PeriodText.Font.Bold = IsBold

End Sub


Public Sub UpdatePeriod(ByVal Caption As String, _
                        ByVal CellDate As Date, _
                        ByVal Value As Long)

    PeriodDate = CellDate
    PeriodValue = Value

    PeriodText.Caption = Caption

    Render

End Sub


'----------------------------------------
' Click
'----------------------------------------

Private Sub PeriodBackground_Click()

    HandlePeriodClick

End Sub


Private Sub PeriodText_Click()

    HandlePeriodClick

End Sub


Private Sub HandlePeriodClick()

    If ParentView.IsPeriodOutsideRange(PeriodDate) Then Exit Sub

    ParentView.PeriodClicked PeriodValue

End Sub


'----------------------------------------
' Pressed
'----------------------------------------

' Common
'--------------------

Private Sub SetPressed(ByVal Value As Boolean)

    If IsPressed = Value Then Exit Sub

    IsPressed = Value
    Render

End Sub


' Mouse down
'--------------------

Private Sub PeriodBackground_MouseDown(ByVal Button As Integer, _
                                       ByVal Shift As Integer, _
                                       ByVal X As Single, _
                                       ByVal Y As Single)

    HandleMouseDown

End Sub


Private Sub PeriodText_MouseDown(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    HandleMouseDown

End Sub


Private Sub HandleMouseDown()

    If ParentView.IsPeriodOutsideRange(PeriodDate) Then Exit Sub

    SetPressed True

End Sub


' Mouse up
'--------------------

Private Sub PeriodBackground_MouseUp(ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single)

    SetPressed False

End Sub


Private Sub PeriodText_MouseUp(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Single, _
                               ByVal Y As Single)

    SetPressed False

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Private Sub PeriodBackground_MouseMove(ByVal Button As Integer, _
                                       ByVal Shift As Integer, _
                                       ByVal X As Single, _
                                       ByVal Y As Single)

    SetHover

End Sub


Private Sub PeriodText_MouseMove(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    SetHover

End Sub


Private Sub SetHover()

    ' ResetHeaderHover (Month) is handled in ResetPeriodHoverExcept.

    ParentView.ResetArrowHover
    ParentView.ResetPeriodHoverExcept Me
    ParentView.ResetActionButtonHover

    If Not ParentView.IsPeriodOutsideRange(PeriodDate) Then
        SetHovered True
    End If

End Sub


Private Sub SetHovered(ByVal Value As Boolean)

    If IsHovered = Value Then Exit Sub

    IsHovered = Value
    Render

End Sub


Public Sub ResetHover()

    SetHovered False

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    PeriodBackground.Visible = IsVisible
    PeriodText.Visible = IsVisible

End Sub

