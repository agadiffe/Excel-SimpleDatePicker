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

Private WithEvents HeaderBackground As MSForms.Label
Attribute HeaderBackground.VB_VarHelpID = -1
Private WithEvents HeaderText As MSForms.Label
Attribute HeaderText.VB_VarHelpID = -1

Private ParentView As Object

Private Interactive As Boolean
Private IsHovered As Boolean
Private IsPressed As Boolean


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub Initialize(ByVal View As Object, _
                      ByVal HeaderCaption As String)

    Set ParentView = View

    Interactive = True

    Set HeaderBackground = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With HeaderBackground
        .Left = DP_HEADER_LEFT
        .Top = DP_HEADER_TOP
        .Width = DP_HEADER_WIDTH
        .Height = DP_HEADER_HEIGHT
    End With

    Set HeaderText = ParentView.ParentForm.Controls.Add("Forms.Label.1")

    With HeaderText
        .Caption = HeaderCaption
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


'----------------------------------------
' Render
'----------------------------------------

Private Sub Render()

    HeaderBackground.BackColor = DP_ColorBg()
    HeaderText.ForeColor = DP_ColorTextHeader()

    If Interactive Then
        If IsPressed Then
            HeaderBackground.BackColor = DP_ColorBgPressed()
        ElseIf IsHovered Then
            HeaderBackground.BackColor = DP_ColorBgHover()
        End If
    End If

End Sub


Public Sub RefreshTheme()

    Render

End Sub


'----------------------------------------
' Header
'----------------------------------------

Public Sub SetCaption(ByVal Caption As String)

    HeaderText.Caption = Caption

End Sub


Public Sub SetInteractive(ByVal Value As Boolean)

    Interactive = Value

End Sub


'----------------------------------------
' Click
'----------------------------------------

Private Sub HeaderBackground_Click()

    HandleClick

End Sub


Private Sub HeaderText_Click()

    HandleClick

End Sub


Private Sub HandleClick()

    If Not Interactive Then Exit Sub

    ParentView.HeaderClicked

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

Private Sub HeaderBackground_MouseDown(ByVal Button As Integer, _
                                       ByVal Shift As Integer, _
                                       ByVal X As Single, _
                                       ByVal Y As Single)

    HandleMouseDown

End Sub


Private Sub HeaderText_MouseDown(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    HandleMouseDown

End Sub


Private Sub HandleMouseDown()

    If Not Interactive Then Exit Sub

    SetPressed True

End Sub


' Mouse up
'--------------------

Private Sub HeaderBackground_MouseUp(ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single)

    SetPressed False

End Sub


Private Sub HeaderText_MouseUp(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Single, _
                               ByVal Y As Single)

    SetPressed False

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Private Sub HeaderBackground_MouseMove(ByVal Button As Integer, _
                                       ByVal Shift As Integer, _
                                       ByVal X As Single, _
                                       ByVal Y As Single)

    SetHover

End Sub


Private Sub HeaderText_MouseMove(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)

    SetHover

End Sub


Private Sub SetHover()

    ParentView.ResetAllHover

    If Interactive Then
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

    HeaderBackground.Visible = IsVisible
    HeaderText.Visible = IsVisible

End Sub

