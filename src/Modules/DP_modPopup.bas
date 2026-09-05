Attribute VB_Name = "DP_modPopup"
Option Explicit


' Opens and positions a popup form next to the specified Excel cell,
' automatically adjusting its position to keep it within the Excel window.


Private Const POPUP_GAP As Single = 0
Private Const POPUP_MARGIN As Single = 10
Private Const CALIBRATION_POINTS_BASE As Double = 7200#


Public Sub ShowPopupNextToCell(ByVal Popup As Object, _
                               ByVal Target As Range)

    Dim Position As Variant

    Position = GetPopupPosition(Target)

    With Popup
        .StartUpPosition = 0
        .Left = Position(0)
        .Top = Position(1)
    End With

    AdjustPopupPosition Popup, Target
    Popup.Show

End Sub


Private Function GetPopupPosition(ByVal Target As Range) As Variant

    Dim CellRight As Double
    Dim CellTop As Double

    With ActiveWindow
        CellRight = .ActivePane.PointsToScreenPixelsX(Target.Left + Target.Width) * GetPointsPerPixelX
        CellTop = .ActivePane.PointsToScreenPixelsY(Target.Top) * GetPointsPerPixelY
    End With

    GetPopupPosition = Array(CellRight + POPUP_GAP, CellTop)

End Function


Private Sub AdjustPopupPosition(ByVal Popup As Object, _
                                ByVal Target As Range)

    Dim ExcelRight As Double
    Dim ExcelBottom As Double

    Dim CellLeft As Double
    Dim CellTop As Double

    ExcelRight = Application.Left + Application.Width - POPUP_MARGIN
    ExcelBottom = Application.Top + Application.Height - POPUP_MARGIN

    CellLeft = GetCellScreenLeft(Target)
    CellTop = GetCellScreenTop(Target)

    ' Horizontal
    '--------------------

    If Popup.Left + Popup.Width > ExcelRight Then
        Popup.Left = CellLeft - Popup.Width - POPUP_GAP
    End If

    If Popup.Left < Application.Left + POPUP_MARGIN Then
        Popup.Left = Application.Left + POPUP_MARGIN
    End If

    ' Vertical
    '--------------------

    If Popup.Top + Popup.Height > ExcelBottom Then
        Popup.Top = ExcelBottom - Popup.Height
    End If

    If Popup.Top < Application.Top + POPUP_MARGIN Then
        Popup.Top = Application.Top + POPUP_MARGIN
    End If

End Sub


Private Function GetCellScreenLeft(ByVal Target As Range) As Double

    GetCellScreenLeft = ActiveWindow.ActivePane.PointsToScreenPixelsX(Target.Left) * GetPointsPerPixelX

End Function


Private Function GetCellScreenTop(ByVal Target As Range) As Double

    GetCellScreenTop = ActiveWindow.ActivePane.PointsToScreenPixelsY(Target.Top) * GetPointsPerPixelY

End Function


Private Function GetZoomAdjustedCalibrationPoints() As Double

    GetZoomAdjustedCalibrationPoints = CALIBRATION_POINTS_BASE / (ActiveWindow.Zoom / 100#)

End Function


Private Function GetPointsPerPixelX() As Double

    Dim CalibrationPoints As Double

    CalibrationPoints = GetZoomAdjustedCalibrationPoints()

    GetPointsPerPixelX = _
        1# / _
        ( _
            (ActiveWindow.ActivePane.PointsToScreenPixelsX(CalibrationPoints) - _
              ActiveWindow.ActivePane.PointsToScreenPixelsX(0)) / _
            CALIBRATION_POINTS_BASE _
        )

End Function


Private Function GetPointsPerPixelY() As Double

    Dim CalibrationPoints As Double

    CalibrationPoints = GetZoomAdjustedCalibrationPoints()

    GetPointsPerPixelY = _
        1# / _
        ( _
            (ActiveWindow.ActivePane.PointsToScreenPixelsY(CalibrationPoints) - _
              ActiveWindow.ActivePane.PointsToScreenPixelsY(0)) / _
            CALIBRATION_POINTS_BASE _
        )

End Function

