VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DP_CHoverHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------
' References
'----------------------------------------

Private HoverLabel As MSForms.Label
Private BackgroundLabel As MSForms.Label

Private NormalBackColor As Long
Private Hovered As Boolean


'----------------------------------------
' Setup
'----------------------------------------

Public Sub Setup(ByVal HoverLabelControl As MSForms.Label, _
                 Optional ByVal BackgroundLabelControl As MSForms.Label)

    Set HoverLabel = HoverLabelControl

    If Not BackgroundLabelControl Is Nothing Then
        Set BackgroundLabel = BackgroundLabelControl
    End If

End Sub


'----------------------------------------
' Hover
'----------------------------------------

Public Sub SetHover()

    If Not Hovered Then
        NormalBackColor = HoverLabel.BackColor
        HoverLabel.BackColor = ColorBgHover()
        Hovered = True
    End If

End Sub


Public Sub SetBackgroundHover()

    If Not Hovered Then
        NormalBackColor = BackgroundLabel.BackColor
        Hovered = True
    End If

    ApplyHoverColor

End Sub


Private Sub ApplyHoverColor()

    If NormalBackColor = ColorBgSelected() Then
        BackgroundLabel.BackColor = ColorBgSelectedHover()
    Else
        BackgroundLabel.BackColor = ColorBgHover()
    End If

End Sub


Public Sub ResetHover()

    If Hovered Then

        If Not BackgroundLabel Is Nothing Then
            BackgroundLabel.BackColor = NormalBackColor
        Else
            HoverLabel.BackColor = NormalBackColor
        End If

        Hovered = False

    End If

End Sub

