VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CHeaderPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Private WithEvents HeaderBackground As MSForms.Image
Attribute HeaderBackground.VB_VarHelpID = -1
Private WithEvents HeaderText As MSForms.Label
Attribute HeaderText.VB_VarHelpID = -1

Private ParentView As DP_IPickerView

Private Interactive As Boolean
Private IsHovered As Boolean
Private IsPressed As Boolean


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub Initialize(ByVal View As DP_IPickerView, _
                      ByVal Caption As String)

    Set ParentView = View
    Interactive = True

    Set HeaderBackground = ParentView.ParentForm.Controls.Add("Forms.Image.1")

    With HeaderBackground
        .Left = DP_HEADER_LEFT
        .Top = DP_HEADER_TOP
        .Width = DP_HEADER_WIDTH
        .Height = DP_HEADER_HEIGHT
        .BorderStyle = fmBorderStyleNone
        .PictureSizeMode = fmPictureSizeModeStretch
    End With

    Set HeaderText = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With HeaderText
        .Caption = Caption
        .Font.Size = DP_FONT_SIZE_HEADER * DP_GetFontScale(DP_DATEPICKER_WIDTH)
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False

        .AutoSize = True
        .Top = DP_VerticalTextTop(HeaderBackground.Top, _
                                  HeaderBackground.Height, _
                                  .Height)
        .AutoSize = False

        .Left = HeaderBackground.Left
        .Width = HeaderBackground.Width
        .Height = HeaderBackground.Top + HeaderBackground.Height - .Top
    End With

    Render

End Sub


Public Sub SetCaption(ByVal Caption As String)

    HeaderText.Caption = Caption

End Sub


Public Sub SetInteractive(ByVal Value As Boolean)

    Interactive = Value

End Sub


'----------------------------------------
' Render
'----------------------------------------

Public Sub RefreshTheme()

    Render

End Sub


Private Sub Render()

    Dim BackgroundColor As Long

    BackgroundColor = DP_ColorBg()

    If Interactive Then
        If IsPressed Then
            BackgroundColor = DP_ColorBgPressed()
        ElseIf IsHovered Then
            BackgroundColor = DP_ColorBgHover()
        End If
    End If

    HeaderBackground.BackColor = BackgroundColor
    HeaderText.ForeColor = DP_ColorTextHeader()

    #If Mac Then
        ' No GDI+ on Mac.
        HeaderBackground.Picture = Nothing
    #Else
        If DP_BUTTON_NO_SHAPE Then
            ' Plain Image used as a colored background.
            HeaderBackground.Picture = Nothing
        Else
            ' Render the actual button shape.
            Set HeaderBackground.Picture = _
                DP_GetCachedButtonPicture(DP_BUTTON_TYPE_HEADER, _
                                          BackgroundColor, _
                                          DP_ColorBg())

            ' MSForms Image/Label overlap bug: reset Z-order.
            HeaderBackground.ZOrder 1
            HeaderText.ZOrder 0
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

    If Not Interactive Then Exit Sub

    ParentView.HeaderClicked

End Sub


Private Sub HeaderBackground_Click()

    HandleClick

End Sub


Private Sub HeaderText_Click()

    HandleClick

End Sub


'----------------------------------------
' Pressed
'----------------------------------------

' Mouse down
'--------------------

Private Sub HandleMouseDown(ByVal Button As Integer)

    If Button <> 1 Then Exit Sub
    If Not Interactive Then Exit Sub

    Pressed = True

End Sub


Private Sub HeaderBackground_MouseDown(ByVal Button As Integer, _
                                       ByVal Shift As Integer, _
                                       ByVal X As Single, _
                                       ByVal Y As Single)

    HandleMouseDown Button

End Sub


Private Sub HeaderText_MouseDown(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    HandleMouseDown Button

End Sub


' Mouse up
'--------------------

Private Sub HeaderBackground_MouseUp(ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single)

    Pressed = False

End Sub


Private Sub HeaderText_MouseUp(ByVal Button As Integer, _
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

    If Interactive Then
        ParentView.SetHoveredButton Me
    Else
        ParentView.ClearHoveredButton
    End If

End Sub


Private Sub HeaderBackground_MouseMove(ByVal Button As Integer, _
                                       ByVal Shift As Integer, _
                                       ByVal X As Single, _
                                       ByVal Y As Single)

    HandleMouseHover

End Sub


Private Sub HeaderText_MouseMove(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    HandleMouseHover

End Sub


'----------------------------------------
' Visibility
'----------------------------------------

Public Sub SetVisible(ByVal IsVisible As Boolean)

    HeaderBackground.Visible = IsVisible
    HeaderText.Visible = IsVisible

End Sub

